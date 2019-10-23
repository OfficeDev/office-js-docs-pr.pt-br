---
title: Office. Context. Mailbox – conjunto de requisitos 1,7
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 87e5334879bb4b5fa84700a03f6da86d4c72e7d2
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627073"
---
# <a name="mailbox"></a><span data-ttu-id="4f1ea-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="4f1ea-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="4f1ea-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="4f1ea-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="4f1ea-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4f1ea-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-105">Requirements</span></span>

|<span data-ttu-id="4f1ea-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-106">Requirement</span></span>| <span data-ttu-id="4f1ea-107">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4f1ea-109">1.0</span></span>|
|[<span data-ttu-id="4f1ea-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-111">Restricted</span></span>|
|[<span data-ttu-id="4f1ea-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4f1ea-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-114">Members and methods</span></span>

| <span data-ttu-id="4f1ea-115">Membro</span><span class="sxs-lookup"><span data-stu-id="4f1ea-115">Member</span></span> | <span data-ttu-id="4f1ea-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4f1ea-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="4f1ea-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="4f1ea-118">Membro</span><span class="sxs-lookup"><span data-stu-id="4f1ea-118">Member</span></span> |
| [<span data-ttu-id="4f1ea-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="4f1ea-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="4f1ea-120">Membro</span><span class="sxs-lookup"><span data-stu-id="4f1ea-120">Member</span></span> |
| [<span data-ttu-id="4f1ea-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4f1ea-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="4f1ea-122">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-122">Method</span></span> |
| [<span data-ttu-id="4f1ea-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="4f1ea-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="4f1ea-124">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-124">Method</span></span> |
| [<span data-ttu-id="4f1ea-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4f1ea-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="4f1ea-126">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-126">Method</span></span> |
| [<span data-ttu-id="4f1ea-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="4f1ea-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="4f1ea-128">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-128">Method</span></span> |
| [<span data-ttu-id="4f1ea-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="4f1ea-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="4f1ea-130">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-130">Method</span></span> |
| [<span data-ttu-id="4f1ea-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="4f1ea-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="4f1ea-132">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-132">Method</span></span> |
| [<span data-ttu-id="4f1ea-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="4f1ea-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="4f1ea-134">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-134">Method</span></span> |
| [<span data-ttu-id="4f1ea-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="4f1ea-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="4f1ea-136">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-136">Method</span></span> |
| [<span data-ttu-id="4f1ea-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="4f1ea-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="4f1ea-138">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-138">Method</span></span> |
| [<span data-ttu-id="4f1ea-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4f1ea-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="4f1ea-140">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-140">Method</span></span> |
| [<span data-ttu-id="4f1ea-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4f1ea-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="4f1ea-142">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-142">Method</span></span> |
| [<span data-ttu-id="4f1ea-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4f1ea-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="4f1ea-144">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-144">Method</span></span> |
| [<span data-ttu-id="4f1ea-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="4f1ea-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="4f1ea-146">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-146">Method</span></span> |
| [<span data-ttu-id="4f1ea-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4f1ea-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="4f1ea-148">Método</span><span class="sxs-lookup"><span data-stu-id="4f1ea-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4f1ea-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="4f1ea-149">Namespaces</span></span>

<span data-ttu-id="4f1ea-150">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="4f1ea-151">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="4f1ea-152">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="4f1ea-153">Members</span><span class="sxs-lookup"><span data-stu-id="4f1ea-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="4f1ea-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-154">ewsUrl: String</span></span>

<span data-ttu-id="4f1ea-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4f1ea-157">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f1ea-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4f1ea-160">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="4f1ea-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="4f1ea-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-163">Type</span></span>

*   <span data-ttu-id="4f1ea-164">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4f1ea-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-165">Requirements</span></span>

|<span data-ttu-id="4f1ea-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-166">Requirement</span></span>| <span data-ttu-id="4f1ea-167">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-169">1.0</span><span class="sxs-lookup"><span data-stu-id="4f1ea-169">1.0</span></span>|
|[<span data-ttu-id="4f1ea-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-171">ReadItem</span></span>|
|[<span data-ttu-id="4f1ea-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-173">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="4f1ea-174">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-174">restUrl: String</span></span>

<span data-ttu-id="4f1ea-175">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="4f1ea-176">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="4f1ea-177">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="4f1ea-p104">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="4f1ea-180">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-180">Type</span></span>

*   <span data-ttu-id="4f1ea-181">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4f1ea-182">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-182">Requirements</span></span>

|<span data-ttu-id="4f1ea-183">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-183">Requirement</span></span>| <span data-ttu-id="4f1ea-184">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-185">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-186">1,5</span><span class="sxs-lookup"><span data-stu-id="4f1ea-186">1.5</span></span> |
|[<span data-ttu-id="4f1ea-187">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-188">ReadItem</span></span>|
|[<span data-ttu-id="4f1ea-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4f1ea-191">Métodos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="4f1ea-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4f1ea-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="4f1ea-193">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="4f1ea-194">Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-194">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-195">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-195">Parameters</span></span>

| <span data-ttu-id="4f1ea-196">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-196">Name</span></span> | <span data-ttu-id="4f1ea-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-197">Type</span></span> | <span data-ttu-id="4f1ea-198">Atributos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-198">Attributes</span></span> | <span data-ttu-id="4f1ea-199">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4f1ea-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4f1ea-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4f1ea-201">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="4f1ea-202">Função</span><span class="sxs-lookup"><span data-stu-id="4f1ea-202">Function</span></span> || <span data-ttu-id="4f1ea-p105">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="4f1ea-206">Objeto</span><span class="sxs-lookup"><span data-stu-id="4f1ea-206">Object</span></span> | <span data-ttu-id="4f1ea-207">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-207">&lt;optional&gt;</span></span> | <span data-ttu-id="4f1ea-208">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4f1ea-209">Objeto</span><span class="sxs-lookup"><span data-stu-id="4f1ea-209">Object</span></span> | <span data-ttu-id="4f1ea-210">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-210">&lt;optional&gt;</span></span> | <span data-ttu-id="4f1ea-211">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4f1ea-212">function</span><span class="sxs-lookup"><span data-stu-id="4f1ea-212">function</span></span>| <span data-ttu-id="4f1ea-213">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-213">&lt;optional&gt;</span></span>|<span data-ttu-id="4f1ea-214">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4f1ea-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f1ea-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-215">Requirements</span></span>

|<span data-ttu-id="4f1ea-216">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-216">Requirement</span></span>| <span data-ttu-id="4f1ea-217">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-218">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-219">1,5</span><span class="sxs-lookup"><span data-stu-id="4f1ea-219">1.5</span></span> |
|[<span data-ttu-id="4f1ea-220">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-221">ReadItem</span></span> |
|[<span data-ttu-id="4f1ea-222">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-223">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f1ea-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="4f1ea-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="4f1ea-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="4f1ea-226">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="4f1ea-227">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-227">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f1ea-p106">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-230">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-230">Parameters</span></span>

|<span data-ttu-id="4f1ea-231">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-231">Name</span></span>| <span data-ttu-id="4f1ea-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-232">Type</span></span>| <span data-ttu-id="4f1ea-233">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4f1ea-234">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-234">String</span></span>|<span data-ttu-id="4f1ea-235">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="4f1ea-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="4f1ea-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="4f1ea-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="4f1ea-237">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f1ea-238">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-238">Requirements</span></span>

|<span data-ttu-id="4f1ea-239">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-239">Requirement</span></span>| <span data-ttu-id="4f1ea-240">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-241">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-242">1.3</span><span class="sxs-lookup"><span data-stu-id="4f1ea-242">1.3</span></span>|
|[<span data-ttu-id="4f1ea-243">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-244">Restrito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-244">Restricted</span></span>|
|[<span data-ttu-id="4f1ea-245">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4f1ea-247">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4f1ea-247">Returns:</span></span>

<span data-ttu-id="4f1ea-248">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4f1ea-249">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-249">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-17"></a><span data-ttu-id="4f1ea-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="4f1ea-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="4f1ea-251">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="4f1ea-p107">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para as datas e horas. O Outlook em uma área de trabalho usa o fuso horário do computador cliente; o Outlook na Web usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="4f1ea-p108">Se o aplicativo de email estiver sendo executado no Outlook em um cliente da área de trabalho, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador cliente. Se o aplicativo de email estiver sendo executado no Outlook na Web, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-257">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-257">Parameters</span></span>

|<span data-ttu-id="4f1ea-258">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-258">Name</span></span>| <span data-ttu-id="4f1ea-259">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-259">Type</span></span>| <span data-ttu-id="4f1ea-260">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="4f1ea-261">Date</span><span class="sxs-lookup"><span data-stu-id="4f1ea-261">Date</span></span>|<span data-ttu-id="4f1ea-262">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="4f1ea-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f1ea-263">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-263">Requirements</span></span>

|<span data-ttu-id="4f1ea-264">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-264">Requirement</span></span>| <span data-ttu-id="4f1ea-265">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-266">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-267">1.0</span><span class="sxs-lookup"><span data-stu-id="4f1ea-267">1.0</span></span>|
|[<span data-ttu-id="4f1ea-268">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-269">ReadItem</span></span>|
|[<span data-ttu-id="4f1ea-270">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-271">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4f1ea-272">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4f1ea-272">Returns:</span></span>

<span data-ttu-id="4f1ea-273">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4f1ea-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="4f1ea-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="4f1ea-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="4f1ea-275">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="4f1ea-276">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-276">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f1ea-p109">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-279">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-279">Parameters</span></span>

|<span data-ttu-id="4f1ea-280">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-280">Name</span></span>| <span data-ttu-id="4f1ea-281">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-281">Type</span></span>| <span data-ttu-id="4f1ea-282">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4f1ea-283">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-283">String</span></span>|<span data-ttu-id="4f1ea-284">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="4f1ea-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="4f1ea-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="4f1ea-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="4f1ea-286">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f1ea-287">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-287">Requirements</span></span>

|<span data-ttu-id="4f1ea-288">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-288">Requirement</span></span>| <span data-ttu-id="4f1ea-289">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-290">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-291">1.3</span><span class="sxs-lookup"><span data-stu-id="4f1ea-291">1.3</span></span>|
|[<span data-ttu-id="4f1ea-292">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-293">Restrito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-293">Restricted</span></span>|
|[<span data-ttu-id="4f1ea-294">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-295">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4f1ea-296">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4f1ea-296">Returns:</span></span>

<span data-ttu-id="4f1ea-297">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4f1ea-298">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-298">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="4f1ea-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="4f1ea-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="4f1ea-300">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="4f1ea-301">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-302">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-302">Parameters</span></span>

|<span data-ttu-id="4f1ea-303">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-303">Name</span></span>| <span data-ttu-id="4f1ea-304">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-304">Type</span></span>| <span data-ttu-id="4f1ea-305">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="4f1ea-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4f1ea-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)|<span data-ttu-id="4f1ea-307">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f1ea-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-308">Requirements</span></span>

|<span data-ttu-id="4f1ea-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-309">Requirement</span></span>| <span data-ttu-id="4f1ea-310">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-312">1.0</span><span class="sxs-lookup"><span data-stu-id="4f1ea-312">1.0</span></span>|
|[<span data-ttu-id="4f1ea-313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-314">ReadItem</span></span>|
|[<span data-ttu-id="4f1ea-315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-316">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4f1ea-317">Retorna:</span><span class="sxs-lookup"><span data-stu-id="4f1ea-317">Returns:</span></span>

<span data-ttu-id="4f1ea-318">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-318">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="4f1ea-319">Tipo: Data</span><span class="sxs-lookup"><span data-stu-id="4f1ea-319">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="4f1ea-320">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-320">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="4f1ea-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4f1ea-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="4f1ea-322">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4f1ea-323">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-323">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f1ea-324">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4f1ea-p110">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir um instância da série. Isso ocorre porque no Outlook no Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="4f1ea-327">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-327">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="4f1ea-328">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-329">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-329">Parameters</span></span>

|<span data-ttu-id="4f1ea-330">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-330">Name</span></span>| <span data-ttu-id="4f1ea-331">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-331">Type</span></span>| <span data-ttu-id="4f1ea-332">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4f1ea-333">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-333">String</span></span>|<span data-ttu-id="4f1ea-334">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f1ea-335">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-335">Requirements</span></span>

|<span data-ttu-id="4f1ea-336">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-336">Requirement</span></span>| <span data-ttu-id="4f1ea-337">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-338">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-339">1.0</span><span class="sxs-lookup"><span data-stu-id="4f1ea-339">1.0</span></span>|
|[<span data-ttu-id="4f1ea-340">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-341">ReadItem</span></span>|
|[<span data-ttu-id="4f1ea-342">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-343">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f1ea-344">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="4f1ea-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4f1ea-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="4f1ea-346">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="4f1ea-347">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-347">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f1ea-348">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4f1ea-349">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-349">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="4f1ea-350">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="4f1ea-p111">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-353">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-353">Parameters</span></span>

|<span data-ttu-id="4f1ea-354">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-354">Name</span></span>| <span data-ttu-id="4f1ea-355">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-355">Type</span></span>| <span data-ttu-id="4f1ea-356">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4f1ea-357">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-357">String</span></span>|<span data-ttu-id="4f1ea-358">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f1ea-359">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-359">Requirements</span></span>

|<span data-ttu-id="4f1ea-360">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-360">Requirement</span></span>| <span data-ttu-id="4f1ea-361">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-362">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-363">1.0</span><span class="sxs-lookup"><span data-stu-id="4f1ea-363">1.0</span></span>|
|[<span data-ttu-id="4f1ea-364">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-365">ReadItem</span></span>|
|[<span data-ttu-id="4f1ea-366">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4f1ea-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-367">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f1ea-368">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="4f1ea-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="4f1ea-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="4f1ea-370">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4f1ea-371">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-371">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f1ea-p112">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="4f1ea-p113">No Outlook na Web e em dispositivos móveis, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="4f1ea-p114">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="4f1ea-379">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-380">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="4f1ea-381">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-381">All parameters are optional.</span></span>

|<span data-ttu-id="4f1ea-382">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-382">Name</span></span>| <span data-ttu-id="4f1ea-383">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-383">Type</span></span>| <span data-ttu-id="4f1ea-384">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="4f1ea-385">Object</span><span class="sxs-lookup"><span data-stu-id="4f1ea-385">Object</span></span> | <span data-ttu-id="4f1ea-386">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="4f1ea-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="4f1ea-p115">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="4f1ea-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="4f1ea-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="4f1ea-393">Data</span><span class="sxs-lookup"><span data-stu-id="4f1ea-393">Date</span></span> | <span data-ttu-id="4f1ea-394">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="4f1ea-395">Data</span><span class="sxs-lookup"><span data-stu-id="4f1ea-395">Date</span></span> | <span data-ttu-id="4f1ea-396">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="4f1ea-397">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-397">String</span></span> | <span data-ttu-id="4f1ea-p117">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="4f1ea-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="4f1ea-p118">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="4f1ea-403">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-403">String</span></span> | <span data-ttu-id="4f1ea-p119">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="4f1ea-406">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-406">String</span></span> | <span data-ttu-id="4f1ea-p120">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4f1ea-409">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-409">Requirements</span></span>

|<span data-ttu-id="4f1ea-410">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-410">Requirement</span></span>| <span data-ttu-id="4f1ea-411">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-412">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-413">1.0</span><span class="sxs-lookup"><span data-stu-id="4f1ea-413">1.0</span></span>|
|[<span data-ttu-id="4f1ea-414">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-415">ReadItem</span></span>|
|[<span data-ttu-id="4f1ea-416">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-417">Read</span><span class="sxs-lookup"><span data-stu-id="4f1ea-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f1ea-418">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="4f1ea-419">displayNewMessageForm (parâmetros)</span><span class="sxs-lookup"><span data-stu-id="4f1ea-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="4f1ea-420">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="4f1ea-421">O `displayNewMessageForm` método abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-421">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="4f1ea-422">Se os parâmetros forem especificados, os campos de formulário da mensagem serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-422">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="4f1ea-423">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-424">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="4f1ea-425">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-425">All parameters are optional.</span></span>

|<span data-ttu-id="4f1ea-426">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-426">Name</span></span>| <span data-ttu-id="4f1ea-427">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-427">Type</span></span>| <span data-ttu-id="4f1ea-428">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="4f1ea-429">Objeto</span><span class="sxs-lookup"><span data-stu-id="4f1ea-429">Object</span></span> | <span data-ttu-id="4f1ea-430">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="4f1ea-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="4f1ea-432">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha para.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="4f1ea-433">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="4f1ea-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="4f1ea-435">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha CC.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="4f1ea-436">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="4f1ea-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="4f1ea-438">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-438">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="4f1ea-439">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-439">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="4f1ea-440">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-440">String</span></span> | <span data-ttu-id="4f1ea-441">Uma cadeia de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-441">A string containing the subject of the message.</span></span> <span data-ttu-id="4f1ea-442">A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-442">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="4f1ea-443">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-443">String</span></span> | <span data-ttu-id="4f1ea-444">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-444">The HTML body of the message.</span></span> <span data-ttu-id="4f1ea-445">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-445">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="4f1ea-446">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4f1ea-447">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="4f1ea-448">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-448">String</span></span> | <span data-ttu-id="4f1ea-p127">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="4f1ea-451">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-451">String</span></span> | <span data-ttu-id="4f1ea-452">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="4f1ea-453">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-453">String</span></span> | <span data-ttu-id="4f1ea-p128">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="4f1ea-456">Booliano</span><span class="sxs-lookup"><span data-stu-id="4f1ea-456">Boolean</span></span> | <span data-ttu-id="4f1ea-p129">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="4f1ea-459">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4f1ea-459">String</span></span> | <span data-ttu-id="4f1ea-460">Usado somente se `type` estiver definido como `item`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-460">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="4f1ea-461">A ID do item do EWS do email existente que você deseja anexar à nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-461">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="4f1ea-462">Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-462">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="4f1ea-463">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-463">Requirements</span></span>

|<span data-ttu-id="4f1ea-464">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-464">Requirement</span></span>| <span data-ttu-id="4f1ea-465">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-466">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-467">1.6</span><span class="sxs-lookup"><span data-stu-id="4f1ea-467">1.6</span></span> |
|[<span data-ttu-id="4f1ea-468">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-469">ReadItem</span></span>|
|[<span data-ttu-id="4f1ea-470">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-471">Read</span><span class="sxs-lookup"><span data-stu-id="4f1ea-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f1ea-472">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-472">Example</span></span>

```js
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

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="4f1ea-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4f1ea-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="4f1ea-474">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="4f1ea-p131">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="4f1ea-477">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="4f1ea-478">Chamar o `getCallbackTokenAsync` método no modo de leitura requer um nível mínimo de permissão de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-478">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="4f1ea-479">A `getCallbackTokenAsync` chamada no modo de redação requer que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-479">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="4f1ea-480">O [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) método requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-480">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="4f1ea-481">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="4f1ea-481">**REST Tokens**</span></span>

<span data-ttu-id="4f1ea-p133">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="4f1ea-485">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-485">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="4f1ea-486">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="4f1ea-486">**EWS Tokens**</span></span>

<span data-ttu-id="4f1ea-p134">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="4f1ea-489">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-489">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="4f1ea-490">Você pode passar o token e um identificador de anexo ou identificador de item para um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-490">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="4f1ea-491">O sistema de terceiros usa o token como um token de autorização de portador para chamar a operação [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) dos serviços Web do Exchange (EWS) ou a operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) para retornar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-491">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="4f1ea-492">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4f1ea-492">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-493">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-493">Parameters</span></span>

|<span data-ttu-id="4f1ea-494">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-494">Name</span></span>| <span data-ttu-id="4f1ea-495">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-495">Type</span></span>| <span data-ttu-id="4f1ea-496">Atributos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-496">Attributes</span></span>| <span data-ttu-id="4f1ea-497">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-497">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="4f1ea-498">Object</span><span class="sxs-lookup"><span data-stu-id="4f1ea-498">Object</span></span> | <span data-ttu-id="4f1ea-499">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-499">&lt;optional&gt;</span></span> | <span data-ttu-id="4f1ea-500">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-500">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="4f1ea-501">Booliano</span><span class="sxs-lookup"><span data-stu-id="4f1ea-501">Boolean</span></span> |  <span data-ttu-id="4f1ea-502">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-502">&lt;optional&gt;</span></span> | <span data-ttu-id="4f1ea-p136">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4f1ea-505">Objeto</span><span class="sxs-lookup"><span data-stu-id="4f1ea-505">Object</span></span> |  <span data-ttu-id="4f1ea-506">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-506">&lt;optional&gt;</span></span> | <span data-ttu-id="4f1ea-507">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-507">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="4f1ea-508">function</span><span class="sxs-lookup"><span data-stu-id="4f1ea-508">function</span></span>||<span data-ttu-id="4f1ea-509">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4f1ea-509">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4f1ea-510">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-510">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4f1ea-511">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-511">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4f1ea-512">Erros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-512">Errors</span></span>

|<span data-ttu-id="4f1ea-513">Código de erro</span><span class="sxs-lookup"><span data-stu-id="4f1ea-513">Error code</span></span>|<span data-ttu-id="4f1ea-514">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-514">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4f1ea-515">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-515">The request has failed.</span></span> <span data-ttu-id="4f1ea-516">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-516">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4f1ea-517">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-517">The Exchange server returned an error.</span></span> <span data-ttu-id="4f1ea-518">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-518">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4f1ea-519">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-519">The user is no longer connected to the network.</span></span> <span data-ttu-id="4f1ea-520">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-520">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f1ea-521">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-521">Requirements</span></span>

|<span data-ttu-id="4f1ea-522">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-522">Requirement</span></span>| <span data-ttu-id="4f1ea-523">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-524">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-525">1,5</span><span class="sxs-lookup"><span data-stu-id="4f1ea-525">1.5</span></span> |
|[<span data-ttu-id="4f1ea-526">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-527">ReadItem</span></span>|
|[<span data-ttu-id="4f1ea-528">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-529">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="4f1ea-529">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f1ea-530">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-530">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="4f1ea-531">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4f1ea-531">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4f1ea-532">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-532">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="4f1ea-p140">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p140">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="4f1ea-535">Você pode passar o token e um identificador de anexo ou identificador de item para um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-535">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="4f1ea-536">O sistema de terceiros usa o token como um token de autorização de portador para chamar a operação [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) dos serviços Web do Exchange (EWS) ou a operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) para retornar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-536">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="4f1ea-537">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4f1ea-537">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4f1ea-538">Chamar o `getCallbackTokenAsync` método no modo de leitura requer um nível mínimo de permissão de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-538">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="4f1ea-539">A `getCallbackTokenAsync` chamada no modo de redação requer que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-539">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="4f1ea-540">O [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) método requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-540">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-541">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-541">Parameters</span></span>

|<span data-ttu-id="4f1ea-542">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-542">Name</span></span>| <span data-ttu-id="4f1ea-543">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-543">Type</span></span>| <span data-ttu-id="4f1ea-544">Atributos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-544">Attributes</span></span>| <span data-ttu-id="4f1ea-545">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-545">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4f1ea-546">function</span><span class="sxs-lookup"><span data-stu-id="4f1ea-546">function</span></span>||<span data-ttu-id="4f1ea-547">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4f1ea-547">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4f1ea-548">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-548">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4f1ea-549">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-549">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="4f1ea-550">Objeto</span><span class="sxs-lookup"><span data-stu-id="4f1ea-550">Object</span></span>| <span data-ttu-id="4f1ea-551">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-551">&lt;optional&gt;</span></span>|<span data-ttu-id="4f1ea-552">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-552">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4f1ea-553">Erros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-553">Errors</span></span>

|<span data-ttu-id="4f1ea-554">Código de erro</span><span class="sxs-lookup"><span data-stu-id="4f1ea-554">Error code</span></span>|<span data-ttu-id="4f1ea-555">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-555">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4f1ea-556">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-556">The request has failed.</span></span> <span data-ttu-id="4f1ea-557">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-557">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4f1ea-558">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-558">The Exchange server returned an error.</span></span> <span data-ttu-id="4f1ea-559">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-559">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4f1ea-560">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-560">The user is no longer connected to the network.</span></span> <span data-ttu-id="4f1ea-561">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-561">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f1ea-562">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-562">Requirements</span></span>

|<span data-ttu-id="4f1ea-563">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-563">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4f1ea-564">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-565">1.0</span><span class="sxs-lookup"><span data-stu-id="4f1ea-565">1.0</span></span> | <span data-ttu-id="4f1ea-566">1.3</span><span class="sxs-lookup"><span data-stu-id="4f1ea-566">1.3</span></span> |
|[<span data-ttu-id="4f1ea-567">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-567">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-568">ReadItem</span></span> | <span data-ttu-id="4f1ea-569">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-569">ReadItem</span></span> |
|[<span data-ttu-id="4f1ea-570">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-570">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-571">Read</span><span class="sxs-lookup"><span data-stu-id="4f1ea-571">Read</span></span> | <span data-ttu-id="4f1ea-572">Escrever</span><span class="sxs-lookup"><span data-stu-id="4f1ea-572">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="4f1ea-573">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-573">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="4f1ea-574">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4f1ea-574">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4f1ea-575">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-575">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="4f1ea-576">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="4f1ea-576">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-577">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-577">Parameters</span></span>

|<span data-ttu-id="4f1ea-578">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-578">Name</span></span>| <span data-ttu-id="4f1ea-579">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-579">Type</span></span>| <span data-ttu-id="4f1ea-580">Atributos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-580">Attributes</span></span>| <span data-ttu-id="4f1ea-581">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-581">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4f1ea-582">function</span><span class="sxs-lookup"><span data-stu-id="4f1ea-582">function</span></span>||<span data-ttu-id="4f1ea-583">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4f1ea-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4f1ea-584">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-584">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4f1ea-585">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-585">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="4f1ea-586">Objeto</span><span class="sxs-lookup"><span data-stu-id="4f1ea-586">Object</span></span>| <span data-ttu-id="4f1ea-587">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-587">&lt;optional&gt;</span></span>|<span data-ttu-id="4f1ea-588">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-588">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4f1ea-589">Erros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-589">Errors</span></span>

|<span data-ttu-id="4f1ea-590">Código de erro</span><span class="sxs-lookup"><span data-stu-id="4f1ea-590">Error code</span></span>|<span data-ttu-id="4f1ea-591">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-591">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4f1ea-592">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-592">The request has failed.</span></span> <span data-ttu-id="4f1ea-593">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-593">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4f1ea-594">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-594">The Exchange server returned an error.</span></span> <span data-ttu-id="4f1ea-595">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-595">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4f1ea-596">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-596">The user is no longer connected to the network.</span></span> <span data-ttu-id="4f1ea-597">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-597">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f1ea-598">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-598">Requirements</span></span>

|<span data-ttu-id="4f1ea-599">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-599">Requirement</span></span>| <span data-ttu-id="4f1ea-600">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-601">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-602">1.0</span><span class="sxs-lookup"><span data-stu-id="4f1ea-602">1.0</span></span>|
|[<span data-ttu-id="4f1ea-603">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-604">ReadItem</span></span>|
|[<span data-ttu-id="4f1ea-605">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4f1ea-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-606">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-606">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f1ea-607">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-607">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="4f1ea-608">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4f1ea-608">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="4f1ea-609">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-609">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="4f1ea-610">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-610">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="4f1ea-611">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="4f1ea-611">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="4f1ea-612">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="4f1ea-612">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="4f1ea-613">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-613">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="4f1ea-614">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-614">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="4f1ea-615">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-615">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="4f1ea-616">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-616">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="4f1ea-617">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-617">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="4f1ea-p150">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p150">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="4f1ea-620">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-620">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="4f1ea-621">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="4f1ea-621">Version differences</span></span>

<span data-ttu-id="4f1ea-622">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-622">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="4f1ea-p151">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-p151">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-626">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-626">Parameters</span></span>

|<span data-ttu-id="4f1ea-627">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-627">Name</span></span>| <span data-ttu-id="4f1ea-628">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-628">Type</span></span>| <span data-ttu-id="4f1ea-629">Atributos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-629">Attributes</span></span>| <span data-ttu-id="4f1ea-630">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-630">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4f1ea-631">String</span><span class="sxs-lookup"><span data-stu-id="4f1ea-631">String</span></span>||<span data-ttu-id="4f1ea-632">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-632">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="4f1ea-633">function</span><span class="sxs-lookup"><span data-stu-id="4f1ea-633">function</span></span>||<span data-ttu-id="4f1ea-634">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4f1ea-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4f1ea-635">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-635">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="4f1ea-636">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-636">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="4f1ea-637">Objeto</span><span class="sxs-lookup"><span data-stu-id="4f1ea-637">Object</span></span>| <span data-ttu-id="4f1ea-638">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-638">&lt;optional&gt;</span></span>|<span data-ttu-id="4f1ea-639">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-639">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f1ea-640">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-640">Requirements</span></span>

|<span data-ttu-id="4f1ea-641">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-641">Requirement</span></span>| <span data-ttu-id="4f1ea-642">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-643">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-644">1.0</span><span class="sxs-lookup"><span data-stu-id="4f1ea-644">1.0</span></span>|
|[<span data-ttu-id="4f1ea-645">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-646">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="4f1ea-646">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="4f1ea-647">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4f1ea-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-648">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-648">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f1ea-649">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-649">Example</span></span>

<span data-ttu-id="4f1ea-650">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-650">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="4f1ea-651">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4f1ea-651">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="4f1ea-652">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-652">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="4f1ea-653">Atualmente, o único tipo de evento compatível é `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-653">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f1ea-654">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="4f1ea-654">Parameters</span></span>

| <span data-ttu-id="4f1ea-655">Nome</span><span class="sxs-lookup"><span data-stu-id="4f1ea-655">Name</span></span> | <span data-ttu-id="4f1ea-656">Tipo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-656">Type</span></span> | <span data-ttu-id="4f1ea-657">Atributos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-657">Attributes</span></span> | <span data-ttu-id="4f1ea-658">Descrição</span><span class="sxs-lookup"><span data-stu-id="4f1ea-658">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4f1ea-659">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4f1ea-659">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4f1ea-660">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-660">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="4f1ea-661">Objeto</span><span class="sxs-lookup"><span data-stu-id="4f1ea-661">Object</span></span> | <span data-ttu-id="4f1ea-662">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-662">&lt;optional&gt;</span></span> | <span data-ttu-id="4f1ea-663">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-663">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4f1ea-664">Objeto</span><span class="sxs-lookup"><span data-stu-id="4f1ea-664">Object</span></span> | <span data-ttu-id="4f1ea-665">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-665">&lt;optional&gt;</span></span> | <span data-ttu-id="4f1ea-666">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="4f1ea-666">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4f1ea-667">function</span><span class="sxs-lookup"><span data-stu-id="4f1ea-667">function</span></span>| <span data-ttu-id="4f1ea-668">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f1ea-668">&lt;optional&gt;</span></span>|<span data-ttu-id="4f1ea-669">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4f1ea-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f1ea-670">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4f1ea-670">Requirements</span></span>

|<span data-ttu-id="4f1ea-671">Requisito</span><span class="sxs-lookup"><span data-stu-id="4f1ea-671">Requirement</span></span>| <span data-ttu-id="4f1ea-672">Valor</span><span class="sxs-lookup"><span data-stu-id="4f1ea-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f1ea-673">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4f1ea-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f1ea-674">1,5</span><span class="sxs-lookup"><span data-stu-id="4f1ea-674">1.5</span></span> |
|[<span data-ttu-id="4f1ea-675">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="4f1ea-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f1ea-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f1ea-676">ReadItem</span></span> |
|[<span data-ttu-id="4f1ea-677">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="4f1ea-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f1ea-678">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4f1ea-678">Compose or Read</span></span>|
