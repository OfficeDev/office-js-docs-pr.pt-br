---
title: Office. Context. Mailbox – conjunto de requisitos 1,6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 9b91a61d301434886723a55eca9608f004f598eb
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451777"
---
# <a name="mailbox"></a><span data-ttu-id="e000d-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="e000d-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="e000d-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="e000d-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="e000d-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="e000d-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e000d-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-105">Requirements</span></span>

|<span data-ttu-id="e000d-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-106">Requirement</span></span>| <span data-ttu-id="e000d-107">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e000d-109">1.0</span></span>|
|[<span data-ttu-id="e000d-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="e000d-111">Restricted</span></span>|
|[<span data-ttu-id="e000d-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e000d-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="e000d-114">Members and methods</span></span>

| <span data-ttu-id="e000d-115">Membro</span><span class="sxs-lookup"><span data-stu-id="e000d-115">Member</span></span> | <span data-ttu-id="e000d-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e000d-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="e000d-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="e000d-118">Membro</span><span class="sxs-lookup"><span data-stu-id="e000d-118">Member</span></span> |
| [<span data-ttu-id="e000d-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="e000d-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="e000d-120">Membro</span><span class="sxs-lookup"><span data-stu-id="e000d-120">Member</span></span> |
| [<span data-ttu-id="e000d-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e000d-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="e000d-122">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-122">Method</span></span> |
| [<span data-ttu-id="e000d-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="e000d-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="e000d-124">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-124">Method</span></span> |
| [<span data-ttu-id="e000d-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e000d-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="e000d-126">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-126">Method</span></span> |
| [<span data-ttu-id="e000d-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="e000d-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="e000d-128">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-128">Method</span></span> |
| [<span data-ttu-id="e000d-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="e000d-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="e000d-130">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-130">Method</span></span> |
| [<span data-ttu-id="e000d-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e000d-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="e000d-132">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-132">Method</span></span> |
| [<span data-ttu-id="e000d-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="e000d-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="e000d-134">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-134">Method</span></span> |
| [<span data-ttu-id="e000d-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e000d-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="e000d-136">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-136">Method</span></span> |
| [<span data-ttu-id="e000d-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="e000d-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="e000d-138">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-138">Method</span></span> |
| [<span data-ttu-id="e000d-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e000d-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="e000d-140">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-140">Method</span></span> |
| [<span data-ttu-id="e000d-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e000d-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="e000d-142">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-142">Method</span></span> |
| [<span data-ttu-id="e000d-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e000d-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="e000d-144">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-144">Method</span></span> |
| [<span data-ttu-id="e000d-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="e000d-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="e000d-146">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-146">Method</span></span> |
| [<span data-ttu-id="e000d-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e000d-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="e000d-148">Método</span><span class="sxs-lookup"><span data-stu-id="e000d-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="e000d-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="e000d-149">Namespaces</span></span>

<span data-ttu-id="e000d-150">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e000d-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="e000d-151">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e000d-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="e000d-152">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e000d-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="e000d-153">Membros</span><span class="sxs-lookup"><span data-stu-id="e000d-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="e000d-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="e000d-154">ewsUrl :String</span></span>

<span data-ttu-id="e000d-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="e000d-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e000d-157">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e000d-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e000d-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="e000d-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e000d-160">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e000d-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="e000d-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="e000d-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="e000d-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-163">Type</span></span>

*   <span data-ttu-id="e000d-164">String</span><span class="sxs-lookup"><span data-stu-id="e000d-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e000d-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-165">Requirements</span></span>

|<span data-ttu-id="e000d-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-166">Requirement</span></span>| <span data-ttu-id="e000d-167">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-169">1.0</span><span class="sxs-lookup"><span data-stu-id="e000d-169">1.0</span></span>|
|[<span data-ttu-id="e000d-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-171">ReadItem</span></span>|
|[<span data-ttu-id="e000d-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-173">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-173">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="e000d-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="e000d-174">restUrl :String</span></span>

<span data-ttu-id="e000d-175">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="e000d-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="e000d-176">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="e000d-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="e000d-177">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e000d-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="e000d-p104">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="e000d-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="e000d-180">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-180">Type</span></span>

*   <span data-ttu-id="e000d-181">String</span><span class="sxs-lookup"><span data-stu-id="e000d-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e000d-182">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-182">Requirements</span></span>

|<span data-ttu-id="e000d-183">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-183">Requirement</span></span>| <span data-ttu-id="e000d-184">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-185">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-186">1,5</span><span class="sxs-lookup"><span data-stu-id="e000d-186">1.5</span></span> |
|[<span data-ttu-id="e000d-187">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-188">ReadItem</span></span>|
|[<span data-ttu-id="e000d-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="e000d-191">Métodos</span><span class="sxs-lookup"><span data-stu-id="e000d-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="e000d-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e000d-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="e000d-193">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="e000d-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="e000d-194">No momento, o único tipo de evento compatível é `Office.EventType.ItemChanged`, que é invocado quando o usuário seleciona um novo item.</span><span class="sxs-lookup"><span data-stu-id="e000d-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="e000d-195">Este evento é usado por suplementos que implementam um painel de tarefas fixável e permite que o suplemento atualize a interface do usuário do painel de tarefas com base no item selecionado no momento.</span><span class="sxs-lookup"><span data-stu-id="e000d-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-196">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-196">Parameters</span></span>

| <span data-ttu-id="e000d-197">Nome</span><span class="sxs-lookup"><span data-stu-id="e000d-197">Name</span></span> | <span data-ttu-id="e000d-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-198">Type</span></span> | <span data-ttu-id="e000d-199">Atributos</span><span class="sxs-lookup"><span data-stu-id="e000d-199">Attributes</span></span> | <span data-ttu-id="e000d-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e000d-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e000d-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e000d-202">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="e000d-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="e000d-203">Função</span><span class="sxs-lookup"><span data-stu-id="e000d-203">Function</span></span> || <span data-ttu-id="e000d-p106">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="e000d-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="e000d-207">Objeto</span><span class="sxs-lookup"><span data-stu-id="e000d-207">Object</span></span> | <span data-ttu-id="e000d-208">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-208">&lt;optional&gt;</span></span> | <span data-ttu-id="e000d-209">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e000d-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e000d-210">Objeto</span><span class="sxs-lookup"><span data-stu-id="e000d-210">Object</span></span> | <span data-ttu-id="e000d-211">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-211">&lt;optional&gt;</span></span> | <span data-ttu-id="e000d-212">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e000d-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e000d-213">function</span><span class="sxs-lookup"><span data-stu-id="e000d-213">function</span></span>| <span data-ttu-id="e000d-214">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-214">&lt;optional&gt;</span></span>|<span data-ttu-id="e000d-215">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e000d-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e000d-216">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-216">Requirements</span></span>

|<span data-ttu-id="e000d-217">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-217">Requirement</span></span>| <span data-ttu-id="e000d-218">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-219">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-220">1,5</span><span class="sxs-lookup"><span data-stu-id="e000d-220">1.5</span></span> |
|[<span data-ttu-id="e000d-221">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-222">ReadItem</span></span> |
|[<span data-ttu-id="e000d-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-224">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e000d-225">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e000d-225">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="e000d-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e000d-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e000d-227">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="e000d-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="e000d-228">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e000d-228">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e000d-p107">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="e000d-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-231">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-231">Parameters</span></span>

|<span data-ttu-id="e000d-232">Nome</span><span class="sxs-lookup"><span data-stu-id="e000d-232">Name</span></span>| <span data-ttu-id="e000d-233">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-233">Type</span></span>| <span data-ttu-id="e000d-234">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e000d-235">String</span><span class="sxs-lookup"><span data-stu-id="e000d-235">String</span></span>|<span data-ttu-id="e000d-236">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="e000d-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="e000d-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e000d-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="e000d-238">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="e000d-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e000d-239">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-239">Requirements</span></span>

|<span data-ttu-id="e000d-240">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-240">Requirement</span></span>| <span data-ttu-id="e000d-241">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-242">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-243">1.3</span><span class="sxs-lookup"><span data-stu-id="e000d-243">1.3</span></span>|
|[<span data-ttu-id="e000d-244">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-245">Restrito</span><span class="sxs-lookup"><span data-stu-id="e000d-245">Restricted</span></span>|
|[<span data-ttu-id="e000d-246">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-247">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e000d-248">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e000d-248">Returns:</span></span>

<span data-ttu-id="e000d-249">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="e000d-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e000d-250">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e000d-250">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="e000d-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="e000d-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="e000d-252">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="e000d-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="e000d-p108">As datas e horas usadas por um aplicativo de email para o Outlook ou o Outlook Web App podem usar fusos horários diferentes. O Outlook usa o fuso horário do computador cliente; o Outlook Web App usa o fuso horário definido na Centro de administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="e000d-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="e000d-p109">Se o aplicativo de email estiver sendo executado no Outlook, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador do cliente. Se o aplicativo de email estiver sendo executado no Outlook Web App, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="e000d-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-258">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-258">Parameters</span></span>

|<span data-ttu-id="e000d-259">Nome</span><span class="sxs-lookup"><span data-stu-id="e000d-259">Name</span></span>| <span data-ttu-id="e000d-260">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-260">Type</span></span>| <span data-ttu-id="e000d-261">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="e000d-262">Date</span><span class="sxs-lookup"><span data-stu-id="e000d-262">Date</span></span>|<span data-ttu-id="e000d-263">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="e000d-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e000d-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-264">Requirements</span></span>

|<span data-ttu-id="e000d-265">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-265">Requirement</span></span>| <span data-ttu-id="e000d-266">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-267">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-268">1.0</span><span class="sxs-lookup"><span data-stu-id="e000d-268">1.0</span></span>|
|[<span data-ttu-id="e000d-269">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-270">ReadItem</span></span>|
|[<span data-ttu-id="e000d-271">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-272">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e000d-273">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e000d-273">Returns:</span></span>

<span data-ttu-id="e000d-274">Tipo: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="e000d-274">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="e000d-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e000d-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e000d-276">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="e000d-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="e000d-277">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e000d-277">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e000d-p110">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="e000d-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-280">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-280">Parameters</span></span>

|<span data-ttu-id="e000d-281">Nome</span><span class="sxs-lookup"><span data-stu-id="e000d-281">Name</span></span>| <span data-ttu-id="e000d-282">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-282">Type</span></span>| <span data-ttu-id="e000d-283">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e000d-284">String</span><span class="sxs-lookup"><span data-stu-id="e000d-284">String</span></span>|<span data-ttu-id="e000d-285">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="e000d-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="e000d-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e000d-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="e000d-287">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="e000d-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e000d-288">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-288">Requirements</span></span>

|<span data-ttu-id="e000d-289">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-289">Requirement</span></span>| <span data-ttu-id="e000d-290">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-291">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-292">1.3</span><span class="sxs-lookup"><span data-stu-id="e000d-292">1.3</span></span>|
|[<span data-ttu-id="e000d-293">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-294">Restrito</span><span class="sxs-lookup"><span data-stu-id="e000d-294">Restricted</span></span>|
|[<span data-ttu-id="e000d-295">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-296">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e000d-297">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e000d-297">Returns:</span></span>

<span data-ttu-id="e000d-298">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="e000d-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e000d-299">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e000d-299">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="e000d-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="e000d-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="e000d-301">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="e000d-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="e000d-302">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="e000d-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-303">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-303">Parameters</span></span>

|<span data-ttu-id="e000d-304">Nome</span><span class="sxs-lookup"><span data-stu-id="e000d-304">Name</span></span>| <span data-ttu-id="e000d-305">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-305">Type</span></span>| <span data-ttu-id="e000d-306">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="e000d-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e000d-307">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="e000d-308">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="e000d-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e000d-309">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-309">Requirements</span></span>

|<span data-ttu-id="e000d-310">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-310">Requirement</span></span>| <span data-ttu-id="e000d-311">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-312">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-313">1.0</span><span class="sxs-lookup"><span data-stu-id="e000d-313">1.0</span></span>|
|[<span data-ttu-id="e000d-314">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-315">ReadItem</span></span>|
|[<span data-ttu-id="e000d-316">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-317">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e000d-318">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e000d-318">Returns:</span></span>

<span data-ttu-id="e000d-319">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="e000d-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="e000d-320">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="e000d-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e000d-321">Date</span><span class="sxs-lookup"><span data-stu-id="e000d-321">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="e000d-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e000d-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="e000d-323">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="e000d-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e000d-324">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e000d-324">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e000d-325">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="e000d-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e000d-p111">No Outlook para Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir uma instância da série. Isso ocorre porque no Outlook para Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="e000d-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="e000d-328">No Outlook Web App, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e000d-328">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="e000d-329">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="e000d-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-330">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-330">Parameters</span></span>

|<span data-ttu-id="e000d-331">Nome</span><span class="sxs-lookup"><span data-stu-id="e000d-331">Name</span></span>| <span data-ttu-id="e000d-332">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-332">Type</span></span>| <span data-ttu-id="e000d-333">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e000d-334">String</span><span class="sxs-lookup"><span data-stu-id="e000d-334">String</span></span>|<span data-ttu-id="e000d-335">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="e000d-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e000d-336">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-336">Requirements</span></span>

|<span data-ttu-id="e000d-337">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-337">Requirement</span></span>| <span data-ttu-id="e000d-338">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-339">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-340">1.0</span><span class="sxs-lookup"><span data-stu-id="e000d-340">1.0</span></span>|
|[<span data-ttu-id="e000d-341">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-341">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-342">ReadItem</span></span>|
|[<span data-ttu-id="e000d-343">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-343">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-344">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e000d-345">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e000d-345">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="e000d-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e000d-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="e000d-347">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="e000d-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="e000d-348">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e000d-348">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e000d-349">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="e000d-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e000d-350">No Outlook Web App, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e000d-350">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="e000d-351">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="e000d-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="e000d-p112">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="e000d-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-354">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-354">Parameters</span></span>

|<span data-ttu-id="e000d-355">Nome</span><span class="sxs-lookup"><span data-stu-id="e000d-355">Name</span></span>| <span data-ttu-id="e000d-356">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-356">Type</span></span>| <span data-ttu-id="e000d-357">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e000d-358">String</span><span class="sxs-lookup"><span data-stu-id="e000d-358">String</span></span>|<span data-ttu-id="e000d-359">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="e000d-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e000d-360">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-360">Requirements</span></span>

|<span data-ttu-id="e000d-361">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-361">Requirement</span></span>| <span data-ttu-id="e000d-362">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-363">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-364">1.0</span><span class="sxs-lookup"><span data-stu-id="e000d-364">1.0</span></span>|
|[<span data-ttu-id="e000d-365">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-366">ReadItem</span></span>|
|[<span data-ttu-id="e000d-367">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e000d-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-368">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e000d-369">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e000d-369">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="e000d-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="e000d-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="e000d-371">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="e000d-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e000d-372">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="e000d-372">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e000d-p113">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="e000d-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e000d-p114">No Outlook Web App e no OWA para Dispositivos, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="e000d-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="e000d-p115">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="e000d-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="e000d-380">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="e000d-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-381">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="e000d-382">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="e000d-382">All parameters are optional.</span></span>

|<span data-ttu-id="e000d-383">Name</span><span class="sxs-lookup"><span data-stu-id="e000d-383">Name</span></span>| <span data-ttu-id="e000d-384">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-384">Type</span></span>| <span data-ttu-id="e000d-385">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e000d-386">Object</span><span class="sxs-lookup"><span data-stu-id="e000d-386">Object</span></span> | <span data-ttu-id="e000d-387">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="e000d-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="e000d-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e000d-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="e000d-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="e000d-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e000d-p117">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="e000d-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="e000d-394">Data</span><span class="sxs-lookup"><span data-stu-id="e000d-394">Date</span></span> | <span data-ttu-id="e000d-395">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="e000d-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="e000d-396">Data</span><span class="sxs-lookup"><span data-stu-id="e000d-396">Date</span></span> | <span data-ttu-id="e000d-397">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="e000d-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="e000d-398">String</span><span class="sxs-lookup"><span data-stu-id="e000d-398">String</span></span> | <span data-ttu-id="e000d-p118">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e000d-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="e000d-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="e000d-p119">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="e000d-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e000d-404">String</span><span class="sxs-lookup"><span data-stu-id="e000d-404">String</span></span> | <span data-ttu-id="e000d-p120">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e000d-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="e000d-407">String</span><span class="sxs-lookup"><span data-stu-id="e000d-407">String</span></span> | <span data-ttu-id="e000d-p121">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e000d-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e000d-410">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-410">Requirements</span></span>

|<span data-ttu-id="e000d-411">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-411">Requirement</span></span>| <span data-ttu-id="e000d-412">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-413">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-414">1.0</span><span class="sxs-lookup"><span data-stu-id="e000d-414">1.0</span></span>|
|[<span data-ttu-id="e000d-415">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-416">ReadItem</span></span>|
|[<span data-ttu-id="e000d-417">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-418">Read</span><span class="sxs-lookup"><span data-stu-id="e000d-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e000d-419">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e000d-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="e000d-420">displayNewMessageForm (parâmetros)</span><span class="sxs-lookup"><span data-stu-id="e000d-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="e000d-421">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="e000d-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="e000d-422">O `displayNewMessageForm` método abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="e000d-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="e000d-423">Se os parâmetros forem especificados, os campos de formulário da mensagem serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="e000d-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e000d-424">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="e000d-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-425">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="e000d-426">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="e000d-426">All parameters are optional.</span></span>

|<span data-ttu-id="e000d-427">Name</span><span class="sxs-lookup"><span data-stu-id="e000d-427">Name</span></span>| <span data-ttu-id="e000d-428">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-428">Type</span></span>| <span data-ttu-id="e000d-429">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e000d-430">Object</span><span class="sxs-lookup"><span data-stu-id="e000d-430">Object</span></span> | <span data-ttu-id="e000d-431">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="e000d-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="e000d-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e000d-433">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha para.</span><span class="sxs-lookup"><span data-stu-id="e000d-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="e000d-434">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="e000d-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="e000d-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e000d-436">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha CC.</span><span class="sxs-lookup"><span data-stu-id="e000d-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="e000d-437">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="e000d-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="e000d-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="e000d-439">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="e000d-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="e000d-440">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="e000d-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e000d-441">String</span><span class="sxs-lookup"><span data-stu-id="e000d-441">String</span></span> | <span data-ttu-id="e000d-442">Uma cadeia de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="e000d-442">A string containing the subject of the message.</span></span> <span data-ttu-id="e000d-443">A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e000d-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="e000d-444">String</span><span class="sxs-lookup"><span data-stu-id="e000d-444">String</span></span> | <span data-ttu-id="e000d-445">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="e000d-445">The HTML body of the message.</span></span> <span data-ttu-id="e000d-446">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e000d-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="e000d-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="e000d-448">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="e000d-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="e000d-449">String</span><span class="sxs-lookup"><span data-stu-id="e000d-449">String</span></span> | <span data-ttu-id="e000d-p128">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="e000d-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="e000d-452">String</span><span class="sxs-lookup"><span data-stu-id="e000d-452">String</span></span> | <span data-ttu-id="e000d-453">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="e000d-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="e000d-454">String</span><span class="sxs-lookup"><span data-stu-id="e000d-454">String</span></span> | <span data-ttu-id="e000d-p129">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="e000d-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="e000d-457">Booliano</span><span class="sxs-lookup"><span data-stu-id="e000d-457">Boolean</span></span> | <span data-ttu-id="e000d-p130">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="e000d-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="e000d-460">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e000d-460">String</span></span> | <span data-ttu-id="e000d-461">Usado somente se `type` estiver definido como `item`.</span><span class="sxs-lookup"><span data-stu-id="e000d-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="e000d-462">A ID do item do EWS do email existente que você deseja anexar à nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="e000d-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="e000d-463">Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e000d-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="e000d-464">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-464">Requirements</span></span>

|<span data-ttu-id="e000d-465">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-465">Requirement</span></span>| <span data-ttu-id="e000d-466">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-467">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-468">1.6</span><span class="sxs-lookup"><span data-stu-id="e000d-468">1.6</span></span> |
|[<span data-ttu-id="e000d-469">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-470">ReadItem</span></span>|
|[<span data-ttu-id="e000d-471">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-472">Read</span><span class="sxs-lookup"><span data-stu-id="e000d-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e000d-473">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e000d-473">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="e000d-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="e000d-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="e000d-475">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="e000d-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="e000d-p132">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="e000d-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="e000d-478">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="e000d-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="e000d-479">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="e000d-479">**REST Tokens**</span></span>

<span data-ttu-id="e000d-p133">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="e000d-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="e000d-483">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="e000d-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="e000d-484">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="e000d-484">**EWS Tokens**</span></span>

<span data-ttu-id="e000d-p134">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="e000d-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="e000d-487">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="e000d-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-488">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-488">Parameters</span></span>

|<span data-ttu-id="e000d-489">Nome</span><span class="sxs-lookup"><span data-stu-id="e000d-489">Name</span></span>| <span data-ttu-id="e000d-490">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-490">Type</span></span>| <span data-ttu-id="e000d-491">Atributos</span><span class="sxs-lookup"><span data-stu-id="e000d-491">Attributes</span></span>| <span data-ttu-id="e000d-492">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="e000d-493">Objeto</span><span class="sxs-lookup"><span data-stu-id="e000d-493">Object</span></span> | <span data-ttu-id="e000d-494">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-494">&lt;optional&gt;</span></span> | <span data-ttu-id="e000d-495">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e000d-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="e000d-496">Booliano</span><span class="sxs-lookup"><span data-stu-id="e000d-496">Boolean</span></span> |  <span data-ttu-id="e000d-497">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-497">&lt;optional&gt;</span></span> | <span data-ttu-id="e000d-p135">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="e000d-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e000d-500">Objeto</span><span class="sxs-lookup"><span data-stu-id="e000d-500">Object</span></span> |  <span data-ttu-id="e000d-501">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-501">&lt;optional&gt;</span></span> | <span data-ttu-id="e000d-502">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="e000d-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="e000d-503">function</span><span class="sxs-lookup"><span data-stu-id="e000d-503">function</span></span>||<span data-ttu-id="e000d-p136">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e000d-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e000d-506">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-506">Requirements</span></span>

|<span data-ttu-id="e000d-507">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-507">Requirement</span></span>| <span data-ttu-id="e000d-508">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-509">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-510">1,5</span><span class="sxs-lookup"><span data-stu-id="e000d-510">1.5</span></span> |
|[<span data-ttu-id="e000d-511">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-512">ReadItem</span></span>|
|[<span data-ttu-id="e000d-513">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-514">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="e000d-514">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="e000d-515">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e000d-515">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="e000d-516">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e000d-516">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e000d-517">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="e000d-517">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="e000d-p137">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="e000d-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="e000d-p138">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="e000d-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e000d-523">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e000d-523">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="e000d-p139">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="e000d-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-526">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-526">Parameters</span></span>

|<span data-ttu-id="e000d-527">Nome</span><span class="sxs-lookup"><span data-stu-id="e000d-527">Name</span></span>| <span data-ttu-id="e000d-528">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-528">Type</span></span>| <span data-ttu-id="e000d-529">Atributos</span><span class="sxs-lookup"><span data-stu-id="e000d-529">Attributes</span></span>| <span data-ttu-id="e000d-530">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-530">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e000d-531">function</span><span class="sxs-lookup"><span data-stu-id="e000d-531">function</span></span>||<span data-ttu-id="e000d-p140">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e000d-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="e000d-534">Objeto</span><span class="sxs-lookup"><span data-stu-id="e000d-534">Object</span></span>| <span data-ttu-id="e000d-535">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-535">&lt;optional&gt;</span></span>|<span data-ttu-id="e000d-536">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="e000d-536">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e000d-537">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-537">Requirements</span></span>

|<span data-ttu-id="e000d-538">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-538">Requirement</span></span>| <span data-ttu-id="e000d-539">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-540">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-541">1.3</span><span class="sxs-lookup"><span data-stu-id="e000d-541">1.3</span></span>|
|[<span data-ttu-id="e000d-542">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-542">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-543">ReadItem</span></span>|
|[<span data-ttu-id="e000d-544">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-545">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="e000d-545">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="e000d-546">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e000d-546">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="e000d-547">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e000d-547">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e000d-548">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="e000d-548">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="e000d-549">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="e000d-549">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-550">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-550">Parameters</span></span>

|<span data-ttu-id="e000d-551">Nome</span><span class="sxs-lookup"><span data-stu-id="e000d-551">Name</span></span>| <span data-ttu-id="e000d-552">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-552">Type</span></span>| <span data-ttu-id="e000d-553">Atributos</span><span class="sxs-lookup"><span data-stu-id="e000d-553">Attributes</span></span>| <span data-ttu-id="e000d-554">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-554">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e000d-555">function</span><span class="sxs-lookup"><span data-stu-id="e000d-555">function</span></span>||<span data-ttu-id="e000d-556">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e000d-556">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e000d-557">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e000d-557">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="e000d-558">Object</span><span class="sxs-lookup"><span data-stu-id="e000d-558">Object</span></span>| <span data-ttu-id="e000d-559">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-559">&lt;optional&gt;</span></span>|<span data-ttu-id="e000d-560">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="e000d-560">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e000d-561">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-561">Requirements</span></span>

|<span data-ttu-id="e000d-562">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-562">Requirement</span></span>| <span data-ttu-id="e000d-563">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-564">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-565">1.0</span><span class="sxs-lookup"><span data-stu-id="e000d-565">1.0</span></span>|
|[<span data-ttu-id="e000d-566">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-567">ReadItem</span></span>|
|[<span data-ttu-id="e000d-568">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e000d-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-569">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-569">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e000d-570">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e000d-570">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="e000d-571">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e000d-571">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="e000d-572">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="e000d-572">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="e000d-573">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="e000d-573">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="e000d-574">No Outlook para iOS ou no Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="e000d-574">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="e000d-575">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="e000d-575">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="e000d-576">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="e000d-576">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="e000d-577">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="e000d-577">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="e000d-578">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="e000d-578">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="e000d-579">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="e000d-579">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="e000d-580">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="e000d-580">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="e000d-p142">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="e000d-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="e000d-583">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="e000d-583">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="e000d-584">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="e000d-584">Version differences</span></span>

<span data-ttu-id="e000d-585">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="e000d-585">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="e000d-p143">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="e000d-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-589">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-589">Parameters</span></span>

|<span data-ttu-id="e000d-590">Nome</span><span class="sxs-lookup"><span data-stu-id="e000d-590">Name</span></span>| <span data-ttu-id="e000d-591">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-591">Type</span></span>| <span data-ttu-id="e000d-592">Atributos</span><span class="sxs-lookup"><span data-stu-id="e000d-592">Attributes</span></span>| <span data-ttu-id="e000d-593">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-593">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="e000d-594">String</span><span class="sxs-lookup"><span data-stu-id="e000d-594">String</span></span>||<span data-ttu-id="e000d-595">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="e000d-595">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="e000d-596">function</span><span class="sxs-lookup"><span data-stu-id="e000d-596">function</span></span>||<span data-ttu-id="e000d-597">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e000d-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e000d-598">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e000d-598">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="e000d-599">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="e000d-599">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="e000d-600">Objeto</span><span class="sxs-lookup"><span data-stu-id="e000d-600">Object</span></span>| <span data-ttu-id="e000d-601">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-601">&lt;optional&gt;</span></span>|<span data-ttu-id="e000d-602">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="e000d-602">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e000d-603">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-603">Requirements</span></span>

|<span data-ttu-id="e000d-604">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-604">Requirement</span></span>| <span data-ttu-id="e000d-605">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-606">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-607">1.0</span><span class="sxs-lookup"><span data-stu-id="e000d-607">1.0</span></span>|
|[<span data-ttu-id="e000d-608">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-609">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="e000d-609">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="e000d-610">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-611">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-611">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e000d-612">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e000d-612">Example</span></span>

<span data-ttu-id="e000d-613">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="e000d-613">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="e000d-614">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e000d-614">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="e000d-615">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="e000d-615">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="e000d-616">Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="e000d-616">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e000d-617">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e000d-617">Parameters</span></span>

| <span data-ttu-id="e000d-618">Nome</span><span class="sxs-lookup"><span data-stu-id="e000d-618">Name</span></span> | <span data-ttu-id="e000d-619">Tipo</span><span class="sxs-lookup"><span data-stu-id="e000d-619">Type</span></span> | <span data-ttu-id="e000d-620">Atributos</span><span class="sxs-lookup"><span data-stu-id="e000d-620">Attributes</span></span> | <span data-ttu-id="e000d-621">Descrição</span><span class="sxs-lookup"><span data-stu-id="e000d-621">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e000d-622">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e000d-622">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e000d-623">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="e000d-623">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="e000d-624">Objeto</span><span class="sxs-lookup"><span data-stu-id="e000d-624">Object</span></span> | <span data-ttu-id="e000d-625">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-625">&lt;optional&gt;</span></span> | <span data-ttu-id="e000d-626">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="e000d-626">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e000d-627">Objeto</span><span class="sxs-lookup"><span data-stu-id="e000d-627">Object</span></span> | <span data-ttu-id="e000d-628">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-628">&lt;optional&gt;</span></span> | <span data-ttu-id="e000d-629">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="e000d-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e000d-630">function</span><span class="sxs-lookup"><span data-stu-id="e000d-630">function</span></span>| <span data-ttu-id="e000d-631">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e000d-631">&lt;optional&gt;</span></span>|<span data-ttu-id="e000d-632">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e000d-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e000d-633">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e000d-633">Requirements</span></span>

|<span data-ttu-id="e000d-634">Requisito</span><span class="sxs-lookup"><span data-stu-id="e000d-634">Requirement</span></span>| <span data-ttu-id="e000d-635">Valor</span><span class="sxs-lookup"><span data-stu-id="e000d-635">Value</span></span>|
|---|---|
|[<span data-ttu-id="e000d-636">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e000d-636">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e000d-637">1,5</span><span class="sxs-lookup"><span data-stu-id="e000d-637">1.5</span></span> |
|[<span data-ttu-id="e000d-638">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e000d-638">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e000d-639">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e000d-639">ReadItem</span></span> |
|[<span data-ttu-id="e000d-640">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e000d-640">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e000d-641">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e000d-641">Compose or Read</span></span>|
