---
title: 'Office.context.mailbox: conjunto de requisitos da versão 1.7'
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 449030924b3fcd3c1a134bc3a676bb87a10fe894
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068326"
---
# <a name="mailbox"></a><span data-ttu-id="35a13-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="35a13-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="35a13-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="35a13-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="35a13-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="35a13-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="35a13-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-105">Requirements</span></span>

|<span data-ttu-id="35a13-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-106">Requirement</span></span>| <span data-ttu-id="35a13-107">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-109">1.0</span><span class="sxs-lookup"><span data-stu-id="35a13-109">1.0</span></span>|
|[<span data-ttu-id="35a13-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="35a13-111">Restricted</span></span>|
|[<span data-ttu-id="35a13-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-113">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="35a13-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="35a13-114">Members and methods</span></span>

| <span data-ttu-id="35a13-115">Membro</span><span class="sxs-lookup"><span data-stu-id="35a13-115">Member</span></span> | <span data-ttu-id="35a13-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="35a13-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="35a13-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="35a13-118">Membro</span><span class="sxs-lookup"><span data-stu-id="35a13-118">Member</span></span> |
| [<span data-ttu-id="35a13-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="35a13-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="35a13-120">Membro</span><span class="sxs-lookup"><span data-stu-id="35a13-120">Member</span></span> |
| [<span data-ttu-id="35a13-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="35a13-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="35a13-122">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-122">Method</span></span> |
| [<span data-ttu-id="35a13-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="35a13-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="35a13-124">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-124">Method</span></span> |
| [<span data-ttu-id="35a13-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="35a13-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) | <span data-ttu-id="35a13-126">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-126">Method</span></span> |
| [<span data-ttu-id="35a13-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="35a13-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="35a13-128">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-128">Method</span></span> |
| [<span data-ttu-id="35a13-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="35a13-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="35a13-130">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-130">Method</span></span> |
| [<span data-ttu-id="35a13-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="35a13-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="35a13-132">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-132">Method</span></span> |
| [<span data-ttu-id="35a13-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="35a13-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="35a13-134">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-134">Method</span></span> |
| [<span data-ttu-id="35a13-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="35a13-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="35a13-136">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-136">Method</span></span> |
| [<span data-ttu-id="35a13-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="35a13-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="35a13-138">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-138">Method</span></span> |
| [<span data-ttu-id="35a13-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="35a13-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="35a13-140">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-140">Method</span></span> |
| [<span data-ttu-id="35a13-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="35a13-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="35a13-142">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-142">Method</span></span> |
| [<span data-ttu-id="35a13-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="35a13-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="35a13-144">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-144">Method</span></span> |
| [<span data-ttu-id="35a13-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="35a13-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="35a13-146">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-146">Method</span></span> |
| [<span data-ttu-id="35a13-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="35a13-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="35a13-148">Método</span><span class="sxs-lookup"><span data-stu-id="35a13-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="35a13-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="35a13-149">Namespaces</span></span>

<span data-ttu-id="35a13-150">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="35a13-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="35a13-151">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="35a13-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="35a13-152">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="35a13-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="35a13-153">Membros</span><span class="sxs-lookup"><span data-stu-id="35a13-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="35a13-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="35a13-154">ewsUrl :String</span></span>

<span data-ttu-id="35a13-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="35a13-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="35a13-157">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="35a13-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35a13-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="35a13-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="35a13-160">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="35a13-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="35a13-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="35a13-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="35a13-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-163">Type</span></span>

*   <span data-ttu-id="35a13-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="35a13-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35a13-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-165">Requirements</span></span>

|<span data-ttu-id="35a13-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-166">Requirement</span></span>| <span data-ttu-id="35a13-167">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-169">1.0</span><span class="sxs-lookup"><span data-stu-id="35a13-169">1.0</span></span>|
|[<span data-ttu-id="35a13-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-171">ReadItem</span></span>|
|[<span data-ttu-id="35a13-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-173">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-173">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="35a13-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="35a13-174">restUrl :String</span></span>

<span data-ttu-id="35a13-175">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="35a13-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="35a13-176">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](https://docs.microsoft.com/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="35a13-176">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="35a13-177">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="35a13-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="35a13-p104">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="35a13-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="35a13-180">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-180">Type</span></span>

*   <span data-ttu-id="35a13-181">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="35a13-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="35a13-182">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-182">Requirements</span></span>

|<span data-ttu-id="35a13-183">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-183">Requirement</span></span>| <span data-ttu-id="35a13-184">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-185">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-186">1,5</span><span class="sxs-lookup"><span data-stu-id="35a13-186">1.5</span></span> |
|[<span data-ttu-id="35a13-187">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-187">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-188">ReadItem</span></span>|
|[<span data-ttu-id="35a13-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-189">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-190">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="35a13-191">Métodos</span><span class="sxs-lookup"><span data-stu-id="35a13-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="35a13-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="35a13-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="35a13-193">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="35a13-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="35a13-194">Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="35a13-194">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-195">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-195">Parameters</span></span>

| <span data-ttu-id="35a13-196">Nome</span><span class="sxs-lookup"><span data-stu-id="35a13-196">Name</span></span> | <span data-ttu-id="35a13-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-197">Type</span></span> | <span data-ttu-id="35a13-198">Atributos</span><span class="sxs-lookup"><span data-stu-id="35a13-198">Attributes</span></span> | <span data-ttu-id="35a13-199">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="35a13-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="35a13-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="35a13-201">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="35a13-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="35a13-202">Função</span><span class="sxs-lookup"><span data-stu-id="35a13-202">Function</span></span> || <span data-ttu-id="35a13-p105">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="35a13-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="35a13-206">Objeto</span><span class="sxs-lookup"><span data-stu-id="35a13-206">Object</span></span> | <span data-ttu-id="35a13-207">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-207">&lt;optional&gt;</span></span> | <span data-ttu-id="35a13-208">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="35a13-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="35a13-209">Objeto</span><span class="sxs-lookup"><span data-stu-id="35a13-209">Object</span></span> | <span data-ttu-id="35a13-210">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-210">&lt;optional&gt;</span></span> | <span data-ttu-id="35a13-211">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="35a13-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="35a13-212">function</span><span class="sxs-lookup"><span data-stu-id="35a13-212">function</span></span>| <span data-ttu-id="35a13-213">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-213">&lt;optional&gt;</span></span>|<span data-ttu-id="35a13-214">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35a13-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a13-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-215">Requirements</span></span>

|<span data-ttu-id="35a13-216">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-216">Requirement</span></span>| <span data-ttu-id="35a13-217">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-218">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-219">1,5</span><span class="sxs-lookup"><span data-stu-id="35a13-219">1.5</span></span> |
|[<span data-ttu-id="35a13-220">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-221">ReadItem</span></span> |
|[<span data-ttu-id="35a13-222">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-223">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35a13-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35a13-224">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="35a13-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="35a13-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="35a13-226">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="35a13-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="35a13-227">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="35a13-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35a13-p106">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="35a13-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-230">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-230">Parameters</span></span>

|<span data-ttu-id="35a13-231">Nome</span><span class="sxs-lookup"><span data-stu-id="35a13-231">Name</span></span>| <span data-ttu-id="35a13-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-232">Type</span></span>| <span data-ttu-id="35a13-233">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="35a13-234">String</span><span class="sxs-lookup"><span data-stu-id="35a13-234">String</span></span>|<span data-ttu-id="35a13-235">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="35a13-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="35a13-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="35a13-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="35a13-237">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="35a13-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a13-238">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-238">Requirements</span></span>

|<span data-ttu-id="35a13-239">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-239">Requirement</span></span>| <span data-ttu-id="35a13-240">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-241">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-242">1.3</span><span class="sxs-lookup"><span data-stu-id="35a13-242">1.3</span></span>|
|[<span data-ttu-id="35a13-243">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-244">Restrito</span><span class="sxs-lookup"><span data-stu-id="35a13-244">Restricted</span></span>|
|[<span data-ttu-id="35a13-245">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-246">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35a13-247">Retorna:</span><span class="sxs-lookup"><span data-stu-id="35a13-247">Returns:</span></span>

<span data-ttu-id="35a13-248">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="35a13-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="35a13-249">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35a13-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a><span data-ttu-id="35a13-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="35a13-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span></span>

<span data-ttu-id="35a13-251">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="35a13-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="35a13-p107">As datas e horas usadas por um aplicativo de email para o Outlook ou o Outlook Web App podem usar fusos horários diferentes. O Outlook usa o fuso horário do computador cliente; o Outlook Web App usa o fuso horário definido na Centro de administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="35a13-p107">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="35a13-p108">Se o aplicativo de email estiver sendo executado no Outlook, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador do cliente. Se o aplicativo de email estiver sendo executado no Outlook Web App, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="35a13-p108">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-257">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-257">Parameters</span></span>

|<span data-ttu-id="35a13-258">Nome</span><span class="sxs-lookup"><span data-stu-id="35a13-258">Name</span></span>| <span data-ttu-id="35a13-259">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-259">Type</span></span>| <span data-ttu-id="35a13-260">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="35a13-261">Data</span><span class="sxs-lookup"><span data-stu-id="35a13-261">Date</span></span>|<span data-ttu-id="35a13-262">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="35a13-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a13-263">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-263">Requirements</span></span>

|<span data-ttu-id="35a13-264">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-264">Requirement</span></span>| <span data-ttu-id="35a13-265">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-266">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-267">1.0</span><span class="sxs-lookup"><span data-stu-id="35a13-267">1.0</span></span>|
|[<span data-ttu-id="35a13-268">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-269">ReadItem</span></span>|
|[<span data-ttu-id="35a13-270">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-271">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35a13-272">Retorna:</span><span class="sxs-lookup"><span data-stu-id="35a13-272">Returns:</span></span>

<span data-ttu-id="35a13-273">Tipo: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="35a13-273">Type: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="35a13-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="35a13-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="35a13-275">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="35a13-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="35a13-276">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="35a13-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35a13-p109">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="35a13-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-279">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-279">Parameters</span></span>

|<span data-ttu-id="35a13-280">Nome</span><span class="sxs-lookup"><span data-stu-id="35a13-280">Name</span></span>| <span data-ttu-id="35a13-281">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-281">Type</span></span>| <span data-ttu-id="35a13-282">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="35a13-283">String</span><span class="sxs-lookup"><span data-stu-id="35a13-283">String</span></span>|<span data-ttu-id="35a13-284">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="35a13-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="35a13-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="35a13-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="35a13-286">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="35a13-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a13-287">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-287">Requirements</span></span>

|<span data-ttu-id="35a13-288">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-288">Requirement</span></span>| <span data-ttu-id="35a13-289">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-290">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-291">1.3</span><span class="sxs-lookup"><span data-stu-id="35a13-291">1.3</span></span>|
|[<span data-ttu-id="35a13-292">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-292">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-293">Restrito</span><span class="sxs-lookup"><span data-stu-id="35a13-293">Restricted</span></span>|
|[<span data-ttu-id="35a13-294">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-294">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-295">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35a13-296">Retorna:</span><span class="sxs-lookup"><span data-stu-id="35a13-296">Returns:</span></span>

<span data-ttu-id="35a13-297">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="35a13-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="35a13-298">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35a13-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="35a13-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="35a13-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="35a13-300">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="35a13-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="35a13-301">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="35a13-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-302">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-302">Parameters</span></span>

|<span data-ttu-id="35a13-303">Nome</span><span class="sxs-lookup"><span data-stu-id="35a13-303">Name</span></span>| <span data-ttu-id="35a13-304">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-304">Type</span></span>| <span data-ttu-id="35a13-305">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="35a13-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="35a13-306">LocalClientTime</span></span>](/javascript/api/outlook_1_7/office.LocalClientTime)|<span data-ttu-id="35a13-307">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="35a13-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a13-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-308">Requirements</span></span>

|<span data-ttu-id="35a13-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-309">Requirement</span></span>| <span data-ttu-id="35a13-310">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-312">1.0</span><span class="sxs-lookup"><span data-stu-id="35a13-312">1.0</span></span>|
|[<span data-ttu-id="35a13-313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-314">ReadItem</span></span>|
|[<span data-ttu-id="35a13-315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-316">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="35a13-317">Retorna:</span><span class="sxs-lookup"><span data-stu-id="35a13-317">Returns:</span></span>

<span data-ttu-id="35a13-318">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="35a13-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="35a13-319">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="35a13-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="35a13-320">Date</span><span class="sxs-lookup"><span data-stu-id="35a13-320">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="35a13-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="35a13-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="35a13-322">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="35a13-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="35a13-323">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="35a13-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35a13-324">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="35a13-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="35a13-p110">No Outlook para Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir uma instância da série. Isso ocorre porque no Outlook para Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="35a13-p110">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="35a13-327">No Outlook Web App, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="35a13-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="35a13-328">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="35a13-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-329">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-329">Parameters</span></span>

|<span data-ttu-id="35a13-330">Nome</span><span class="sxs-lookup"><span data-stu-id="35a13-330">Name</span></span>| <span data-ttu-id="35a13-331">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-331">Type</span></span>| <span data-ttu-id="35a13-332">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="35a13-333">String</span><span class="sxs-lookup"><span data-stu-id="35a13-333">String</span></span>|<span data-ttu-id="35a13-334">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="35a13-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a13-335">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-335">Requirements</span></span>

|<span data-ttu-id="35a13-336">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-336">Requirement</span></span>| <span data-ttu-id="35a13-337">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-338">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-339">1.0</span><span class="sxs-lookup"><span data-stu-id="35a13-339">1.0</span></span>|
|[<span data-ttu-id="35a13-340">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-340">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-341">ReadItem</span></span>|
|[<span data-ttu-id="35a13-342">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-342">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-343">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35a13-344">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35a13-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="35a13-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="35a13-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="35a13-346">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="35a13-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="35a13-347">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="35a13-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35a13-348">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="35a13-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="35a13-349">No Outlook Web App, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="35a13-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="35a13-350">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="35a13-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="35a13-p111">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="35a13-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-353">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-353">Parameters</span></span>

|<span data-ttu-id="35a13-354">Nome</span><span class="sxs-lookup"><span data-stu-id="35a13-354">Name</span></span>| <span data-ttu-id="35a13-355">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-355">Type</span></span>| <span data-ttu-id="35a13-356">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="35a13-357">String</span><span class="sxs-lookup"><span data-stu-id="35a13-357">String</span></span>|<span data-ttu-id="35a13-358">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="35a13-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a13-359">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-359">Requirements</span></span>

|<span data-ttu-id="35a13-360">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-360">Requirement</span></span>| <span data-ttu-id="35a13-361">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-362">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-363">1.0</span><span class="sxs-lookup"><span data-stu-id="35a13-363">1.0</span></span>|
|[<span data-ttu-id="35a13-364">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-365">ReadItem</span></span>|
|[<span data-ttu-id="35a13-366">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-367">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35a13-368">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35a13-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="35a13-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="35a13-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="35a13-370">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="35a13-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="35a13-371">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="35a13-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="35a13-p112">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="35a13-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="35a13-p113">No Outlook Web App e no OWA para Dispositivos, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="35a13-p113">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="35a13-p114">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="35a13-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="35a13-379">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="35a13-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-380">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="35a13-381">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="35a13-381">All parameters are optional.</span></span>

|<span data-ttu-id="35a13-382">Name</span><span class="sxs-lookup"><span data-stu-id="35a13-382">Name</span></span>| <span data-ttu-id="35a13-383">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-383">Type</span></span>| <span data-ttu-id="35a13-384">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="35a13-385">Objeto</span><span class="sxs-lookup"><span data-stu-id="35a13-385">Object</span></span> | <span data-ttu-id="35a13-386">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="35a13-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="35a13-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="35a13-p115">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="35a13-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="35a13-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="35a13-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="35a13-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="35a13-393">Data</span><span class="sxs-lookup"><span data-stu-id="35a13-393">Date</span></span> | <span data-ttu-id="35a13-394">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="35a13-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="35a13-395">Data</span><span class="sxs-lookup"><span data-stu-id="35a13-395">Date</span></span> | <span data-ttu-id="35a13-396">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="35a13-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="35a13-397">String</span><span class="sxs-lookup"><span data-stu-id="35a13-397">String</span></span> | <span data-ttu-id="35a13-p117">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="35a13-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="35a13-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="35a13-p118">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="35a13-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="35a13-403">String</span><span class="sxs-lookup"><span data-stu-id="35a13-403">String</span></span> | <span data-ttu-id="35a13-p119">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="35a13-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="35a13-406">String</span><span class="sxs-lookup"><span data-stu-id="35a13-406">String</span></span> | <span data-ttu-id="35a13-p120">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="35a13-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="35a13-409">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-409">Requirements</span></span>

|<span data-ttu-id="35a13-410">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-410">Requirement</span></span>| <span data-ttu-id="35a13-411">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-412">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-413">1.0</span><span class="sxs-lookup"><span data-stu-id="35a13-413">1.0</span></span>|
|[<span data-ttu-id="35a13-414">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-414">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-415">ReadItem</span></span>|
|[<span data-ttu-id="35a13-416">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-416">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-417">Read</span><span class="sxs-lookup"><span data-stu-id="35a13-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35a13-418">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35a13-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="35a13-419">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="35a13-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="35a13-420">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="35a13-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="35a13-421">O método `displayNewMessageForm` abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="35a13-421">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="35a13-422">Quando os parâmetros são especificados, os campos do formulário de mensagem são preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="35a13-422">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="35a13-423">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="35a13-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-424">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="35a13-425">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="35a13-425">All parameters are optional.</span></span>

|<span data-ttu-id="35a13-426">Name</span><span class="sxs-lookup"><span data-stu-id="35a13-426">Name</span></span>| <span data-ttu-id="35a13-427">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-427">Type</span></span>| <span data-ttu-id="35a13-428">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="35a13-429">Objeto</span><span class="sxs-lookup"><span data-stu-id="35a13-429">Object</span></span> | <span data-ttu-id="35a13-430">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="35a13-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="35a13-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="35a13-432">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz que contém um objeto `EmailAddressDetails` para cada um dos destinatários na linha Para.</span><span class="sxs-lookup"><span data-stu-id="35a13-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="35a13-433">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="35a13-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="35a13-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="35a13-435">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz que contém um objeto `EmailAddressDetails` para cada um dos destinatários na linha Cc.</span><span class="sxs-lookup"><span data-stu-id="35a13-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="35a13-436">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="35a13-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="35a13-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="35a13-438">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz que contém um objeto `EmailAddressDetails` para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="35a13-438">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="35a13-439">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="35a13-439">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="35a13-440">String</span><span class="sxs-lookup"><span data-stu-id="35a13-440">String</span></span> | <span data-ttu-id="35a13-441">Uma cadeia de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="35a13-441">A string containing the subject of the message.</span></span> <span data-ttu-id="35a13-442">A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="35a13-442">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="35a13-443">String</span><span class="sxs-lookup"><span data-stu-id="35a13-443">String</span></span> | <span data-ttu-id="35a13-444">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="35a13-444">The HTML body of the message.</span></span> <span data-ttu-id="35a13-445">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="35a13-445">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="35a13-446">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="35a13-447">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="35a13-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="35a13-448">String</span><span class="sxs-lookup"><span data-stu-id="35a13-448">String</span></span> | <span data-ttu-id="35a13-p127">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="35a13-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="35a13-451">String</span><span class="sxs-lookup"><span data-stu-id="35a13-451">String</span></span> | <span data-ttu-id="35a13-452">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="35a13-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="35a13-453">String</span><span class="sxs-lookup"><span data-stu-id="35a13-453">String</span></span> | <span data-ttu-id="35a13-p128">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="35a13-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="35a13-456">Booliano</span><span class="sxs-lookup"><span data-stu-id="35a13-456">Boolean</span></span> | <span data-ttu-id="35a13-p129">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="35a13-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="35a13-459">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="35a13-459">String</span></span> | <span data-ttu-id="35a13-460">Usado somente se `type` estiver definido como `item`.</span><span class="sxs-lookup"><span data-stu-id="35a13-460">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="35a13-461">A ID do item do EWS para o email existente que você deseja anexar à nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="35a13-461">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="35a13-462">Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="35a13-462">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="35a13-463">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-463">Requirements</span></span>

|<span data-ttu-id="35a13-464">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-464">Requirement</span></span>| <span data-ttu-id="35a13-465">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-466">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-467">1.6</span><span class="sxs-lookup"><span data-stu-id="35a13-467">1.6</span></span> |
|[<span data-ttu-id="35a13-468">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-468">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-469">ReadItem</span></span>|
|[<span data-ttu-id="35a13-470">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-470">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-471">Read</span><span class="sxs-lookup"><span data-stu-id="35a13-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35a13-472">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35a13-472">Example</span></span>

```javascript
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="35a13-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="35a13-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="35a13-474">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="35a13-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="35a13-p131">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="35a13-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="35a13-477">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="35a13-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="35a13-478">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="35a13-478">**REST Tokens**</span></span>

<span data-ttu-id="35a13-p132">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="35a13-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="35a13-482">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="35a13-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="35a13-483">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="35a13-483">**EWS Tokens**</span></span>

<span data-ttu-id="35a13-p133">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="35a13-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="35a13-486">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="35a13-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-487">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-487">Parameters</span></span>

|<span data-ttu-id="35a13-488">Nome</span><span class="sxs-lookup"><span data-stu-id="35a13-488">Name</span></span>| <span data-ttu-id="35a13-489">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-489">Type</span></span>| <span data-ttu-id="35a13-490">Atributos</span><span class="sxs-lookup"><span data-stu-id="35a13-490">Attributes</span></span>| <span data-ttu-id="35a13-491">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="35a13-492">Object</span><span class="sxs-lookup"><span data-stu-id="35a13-492">Object</span></span> | <span data-ttu-id="35a13-493">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-493">&lt;optional&gt;</span></span> | <span data-ttu-id="35a13-494">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="35a13-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="35a13-495">Booliano</span><span class="sxs-lookup"><span data-stu-id="35a13-495">Boolean</span></span> |  <span data-ttu-id="35a13-496">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-496">&lt;optional&gt;</span></span> | <span data-ttu-id="35a13-p134">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="35a13-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="35a13-499">Objeto</span><span class="sxs-lookup"><span data-stu-id="35a13-499">Object</span></span> |  <span data-ttu-id="35a13-500">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-500">&lt;optional&gt;</span></span> | <span data-ttu-id="35a13-501">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="35a13-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="35a13-502">function</span><span class="sxs-lookup"><span data-stu-id="35a13-502">function</span></span>||<span data-ttu-id="35a13-p135">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="35a13-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a13-505">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-505">Requirements</span></span>

|<span data-ttu-id="35a13-506">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-506">Requirement</span></span>| <span data-ttu-id="35a13-507">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-508">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-509">1,5</span><span class="sxs-lookup"><span data-stu-id="35a13-509">1.5</span></span> |
|[<span data-ttu-id="35a13-510">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-511">ReadItem</span></span>|
|[<span data-ttu-id="35a13-512">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-513">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-513">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="35a13-514">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35a13-514">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="35a13-515">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="35a13-515">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="35a13-516">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="35a13-516">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="35a13-p136">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="35a13-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="35a13-p137">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="35a13-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="35a13-522">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="35a13-522">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="35a13-p138">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="35a13-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-525">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-525">Parameters</span></span>

|<span data-ttu-id="35a13-526">Nome</span><span class="sxs-lookup"><span data-stu-id="35a13-526">Name</span></span>| <span data-ttu-id="35a13-527">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-527">Type</span></span>| <span data-ttu-id="35a13-528">Atributos</span><span class="sxs-lookup"><span data-stu-id="35a13-528">Attributes</span></span>| <span data-ttu-id="35a13-529">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-529">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="35a13-530">function</span><span class="sxs-lookup"><span data-stu-id="35a13-530">function</span></span>||<span data-ttu-id="35a13-p139">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="35a13-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="35a13-533">Objeto</span><span class="sxs-lookup"><span data-stu-id="35a13-533">Object</span></span>| <span data-ttu-id="35a13-534">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-534">&lt;optional&gt;</span></span>|<span data-ttu-id="35a13-535">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="35a13-535">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a13-536">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-536">Requirements</span></span>

|<span data-ttu-id="35a13-537">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-537">Requirement</span></span>| <span data-ttu-id="35a13-538">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-539">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-540">1.3</span><span class="sxs-lookup"><span data-stu-id="35a13-540">1.3</span></span>|
|[<span data-ttu-id="35a13-541">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-542">ReadItem</span></span>|
|[<span data-ttu-id="35a13-543">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-544">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="35a13-545">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35a13-545">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="35a13-546">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="35a13-546">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="35a13-547">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="35a13-547">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="35a13-548">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="35a13-548">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-549">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-549">Parameters</span></span>

|<span data-ttu-id="35a13-550">Nome</span><span class="sxs-lookup"><span data-stu-id="35a13-550">Name</span></span>| <span data-ttu-id="35a13-551">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-551">Type</span></span>| <span data-ttu-id="35a13-552">Atributos</span><span class="sxs-lookup"><span data-stu-id="35a13-552">Attributes</span></span>| <span data-ttu-id="35a13-553">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-553">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="35a13-554">função</span><span class="sxs-lookup"><span data-stu-id="35a13-554">function</span></span>||<span data-ttu-id="35a13-555">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35a13-555">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="35a13-556">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="35a13-556">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="35a13-557">Object</span><span class="sxs-lookup"><span data-stu-id="35a13-557">Object</span></span>| <span data-ttu-id="35a13-558">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-558">&lt;optional&gt;</span></span>|<span data-ttu-id="35a13-559">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="35a13-559">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a13-560">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-560">Requirements</span></span>

|<span data-ttu-id="35a13-561">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-561">Requirement</span></span>| <span data-ttu-id="35a13-562">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-563">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-564">1.0</span><span class="sxs-lookup"><span data-stu-id="35a13-564">1.0</span></span>|
|[<span data-ttu-id="35a13-565">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-566">ReadItem</span></span>|
|[<span data-ttu-id="35a13-567">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-568">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-568">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35a13-569">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35a13-569">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="35a13-570">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="35a13-570">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="35a13-571">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="35a13-571">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="35a13-572">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="35a13-572">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="35a13-573">No Outlook para iOS ou no Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="35a13-573">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="35a13-574">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="35a13-574">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="35a13-575">Nesses casos, os suplementos devem [usar as APIs REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="35a13-575">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="35a13-576">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="35a13-576">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="35a13-577">Consulte [Chamar serviços Web de um suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="35a13-577">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="35a13-578">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="35a13-578">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="35a13-579">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="35a13-579">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="35a13-p141">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="35a13-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="35a13-582">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="35a13-582">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="35a13-583">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="35a13-583">Version differences</span></span>

<span data-ttu-id="35a13-584">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="35a13-584">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="35a13-p142">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="35a13-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-588">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-588">Parameters</span></span>

|<span data-ttu-id="35a13-589">Nome</span><span class="sxs-lookup"><span data-stu-id="35a13-589">Name</span></span>| <span data-ttu-id="35a13-590">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-590">Type</span></span>| <span data-ttu-id="35a13-591">Atributos</span><span class="sxs-lookup"><span data-stu-id="35a13-591">Attributes</span></span>| <span data-ttu-id="35a13-592">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-592">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="35a13-593">String</span><span class="sxs-lookup"><span data-stu-id="35a13-593">String</span></span>||<span data-ttu-id="35a13-594">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="35a13-594">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="35a13-595">função</span><span class="sxs-lookup"><span data-stu-id="35a13-595">function</span></span>||<span data-ttu-id="35a13-596">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35a13-596">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="35a13-597">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="35a13-597">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="35a13-598">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="35a13-598">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="35a13-599">Objeto</span><span class="sxs-lookup"><span data-stu-id="35a13-599">Object</span></span>| <span data-ttu-id="35a13-600">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-600">&lt;optional&gt;</span></span>|<span data-ttu-id="35a13-601">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="35a13-601">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a13-602">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-602">Requirements</span></span>

|<span data-ttu-id="35a13-603">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-603">Requirement</span></span>| <span data-ttu-id="35a13-604">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-605">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-606">1.0</span><span class="sxs-lookup"><span data-stu-id="35a13-606">1.0</span></span>|
|[<span data-ttu-id="35a13-607">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-607">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-608">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="35a13-608">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="35a13-609">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-609">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-610">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-610">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="35a13-611">Exemplo</span><span class="sxs-lookup"><span data-stu-id="35a13-611">Example</span></span>

<span data-ttu-id="35a13-612">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="35a13-612">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="35a13-613">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="35a13-613">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="35a13-614">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="35a13-614">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="35a13-615">Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="35a13-615">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="35a13-616">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="35a13-616">Parameters</span></span>

| <span data-ttu-id="35a13-617">Nome</span><span class="sxs-lookup"><span data-stu-id="35a13-617">Name</span></span> | <span data-ttu-id="35a13-618">Tipo</span><span class="sxs-lookup"><span data-stu-id="35a13-618">Type</span></span> | <span data-ttu-id="35a13-619">Atributos</span><span class="sxs-lookup"><span data-stu-id="35a13-619">Attributes</span></span> | <span data-ttu-id="35a13-620">Descrição</span><span class="sxs-lookup"><span data-stu-id="35a13-620">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="35a13-621">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="35a13-621">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="35a13-622">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="35a13-622">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="35a13-623">Objeto</span><span class="sxs-lookup"><span data-stu-id="35a13-623">Object</span></span> | <span data-ttu-id="35a13-624">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-624">&lt;optional&gt;</span></span> | <span data-ttu-id="35a13-625">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="35a13-625">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="35a13-626">Objeto</span><span class="sxs-lookup"><span data-stu-id="35a13-626">Object</span></span> | <span data-ttu-id="35a13-627">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-627">&lt;optional&gt;</span></span> | <span data-ttu-id="35a13-628">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="35a13-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="35a13-629">function</span><span class="sxs-lookup"><span data-stu-id="35a13-629">function</span></span>| <span data-ttu-id="35a13-630">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="35a13-630">&lt;optional&gt;</span></span>|<span data-ttu-id="35a13-631">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="35a13-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="35a13-632">Requisitos</span><span class="sxs-lookup"><span data-stu-id="35a13-632">Requirements</span></span>

|<span data-ttu-id="35a13-633">Requisito</span><span class="sxs-lookup"><span data-stu-id="35a13-633">Requirement</span></span>| <span data-ttu-id="35a13-634">Valor</span><span class="sxs-lookup"><span data-stu-id="35a13-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="35a13-635">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="35a13-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="35a13-636">1,5</span><span class="sxs-lookup"><span data-stu-id="35a13-636">1.5</span></span> |
|[<span data-ttu-id="35a13-637">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="35a13-637">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="35a13-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="35a13-638">ReadItem</span></span> |
|[<span data-ttu-id="35a13-639">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="35a13-639">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="35a13-640">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="35a13-640">Compose or Read</span></span>|
