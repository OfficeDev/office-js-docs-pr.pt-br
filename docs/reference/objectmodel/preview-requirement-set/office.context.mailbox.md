---
title: Office. Context. Mailbox-visualização do conjunto de requisitos
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: d19cb7c664cda42469cf7cde31d69f87101278c8
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838533"
---
# <a name="mailbox"></a><span data-ttu-id="a10c1-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="a10c1-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="a10c1-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="a10c1-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="a10c1-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="a10c1-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a10c1-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-105">Requirements</span></span>

|<span data-ttu-id="a10c1-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-106">Requirement</span></span>| <span data-ttu-id="a10c1-107">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a10c1-109">1.0</span></span>|
|[<span data-ttu-id="a10c1-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="a10c1-111">Restricted</span></span>|
|[<span data-ttu-id="a10c1-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a10c1-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="a10c1-114">Members and methods</span></span>

| <span data-ttu-id="a10c1-115">Membro</span><span class="sxs-lookup"><span data-stu-id="a10c1-115">Member</span></span> | <span data-ttu-id="a10c1-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a10c1-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="a10c1-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="a10c1-118">Membro</span><span class="sxs-lookup"><span data-stu-id="a10c1-118">Member</span></span> |
| [<span data-ttu-id="a10c1-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="a10c1-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="a10c1-120">Membro</span><span class="sxs-lookup"><span data-stu-id="a10c1-120">Member</span></span> |
| [<span data-ttu-id="a10c1-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a10c1-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="a10c1-122">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-122">Method</span></span> |
| [<span data-ttu-id="a10c1-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="a10c1-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="a10c1-124">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-124">Method</span></span> |
| [<span data-ttu-id="a10c1-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="a10c1-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="a10c1-126">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-126">Method</span></span> |
| [<span data-ttu-id="a10c1-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="a10c1-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="a10c1-128">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-128">Method</span></span> |
| [<span data-ttu-id="a10c1-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="a10c1-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="a10c1-130">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-130">Method</span></span> |
| [<span data-ttu-id="a10c1-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="a10c1-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="a10c1-132">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-132">Method</span></span> |
| [<span data-ttu-id="a10c1-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="a10c1-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="a10c1-134">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-134">Method</span></span> |
| [<span data-ttu-id="a10c1-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="a10c1-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="a10c1-136">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-136">Method</span></span> |
| [<span data-ttu-id="a10c1-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="a10c1-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="a10c1-138">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-138">Method</span></span> |
| [<span data-ttu-id="a10c1-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="a10c1-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="a10c1-140">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-140">Method</span></span> |
| [<span data-ttu-id="a10c1-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="a10c1-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="a10c1-142">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-142">Method</span></span> |
| [<span data-ttu-id="a10c1-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="a10c1-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="a10c1-144">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-144">Method</span></span> |
| [<span data-ttu-id="a10c1-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="a10c1-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="a10c1-146">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-146">Method</span></span> |
| [<span data-ttu-id="a10c1-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="a10c1-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="a10c1-148">Método</span><span class="sxs-lookup"><span data-stu-id="a10c1-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a10c1-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="a10c1-149">Namespaces</span></span>

<span data-ttu-id="a10c1-150">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a10c1-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="a10c1-151">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a10c1-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="a10c1-152">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a10c1-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="a10c1-153">Membros</span><span class="sxs-lookup"><span data-stu-id="a10c1-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="a10c1-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="a10c1-154">ewsUrl :String</span></span>

<span data-ttu-id="a10c1-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a10c1-157">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a10c1-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a10c1-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="a10c1-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="a10c1-160">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a10c1-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="a10c1-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="a10c1-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-163">Type</span></span>

*   <span data-ttu-id="a10c1-164">String</span><span class="sxs-lookup"><span data-stu-id="a10c1-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a10c1-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-165">Requirements</span></span>

|<span data-ttu-id="a10c1-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-166">Requirement</span></span>| <span data-ttu-id="a10c1-167">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-169">1.0</span><span class="sxs-lookup"><span data-stu-id="a10c1-169">1.0</span></span>|
|[<span data-ttu-id="a10c1-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-171">ReadItem</span></span>|
|[<span data-ttu-id="a10c1-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-173">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-173">Compose or Read</span></span>|

---
---

#### <a name="resturl-string"></a><span data-ttu-id="a10c1-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="a10c1-174">restUrl :String</span></span>

<span data-ttu-id="a10c1-175">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="a10c1-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="a10c1-176">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="a10c1-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="a10c1-177">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a10c1-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="a10c1-p104">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="a10c1-180">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-180">Type</span></span>

*   <span data-ttu-id="a10c1-181">String</span><span class="sxs-lookup"><span data-stu-id="a10c1-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a10c1-182">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-182">Requirements</span></span>

|<span data-ttu-id="a10c1-183">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-183">Requirement</span></span>| <span data-ttu-id="a10c1-184">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-185">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-186">1,5</span><span class="sxs-lookup"><span data-stu-id="a10c1-186">1.5</span></span> |
|[<span data-ttu-id="a10c1-187">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-188">ReadItem</span></span>|
|[<span data-ttu-id="a10c1-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="a10c1-191">Métodos</span><span class="sxs-lookup"><span data-stu-id="a10c1-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="a10c1-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a10c1-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="a10c1-193">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="a10c1-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="a10c1-194">Atualmente, os tipos de eventos com `Office.EventType.ItemChanged` suporte `Office.EventType.OfficeThemeChanged`são e.</span><span class="sxs-lookup"><span data-stu-id="a10c1-194">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-195">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-195">Parameters</span></span>

| <span data-ttu-id="a10c1-196">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-196">Name</span></span> | <span data-ttu-id="a10c1-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-197">Type</span></span> | <span data-ttu-id="a10c1-198">Atributos</span><span class="sxs-lookup"><span data-stu-id="a10c1-198">Attributes</span></span> | <span data-ttu-id="a10c1-199">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a10c1-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a10c1-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a10c1-201">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="a10c1-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="a10c1-202">Função</span><span class="sxs-lookup"><span data-stu-id="a10c1-202">Function</span></span> || <span data-ttu-id="a10c1-p105">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="a10c1-206">Objeto</span><span class="sxs-lookup"><span data-stu-id="a10c1-206">Object</span></span> | <span data-ttu-id="a10c1-207">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-207">&lt;optional&gt;</span></span> | <span data-ttu-id="a10c1-208">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a10c1-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a10c1-209">Objeto</span><span class="sxs-lookup"><span data-stu-id="a10c1-209">Object</span></span> | <span data-ttu-id="a10c1-210">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-210">&lt;optional&gt;</span></span> | <span data-ttu-id="a10c1-211">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a10c1-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a10c1-212">function</span><span class="sxs-lookup"><span data-stu-id="a10c1-212">function</span></span>| <span data-ttu-id="a10c1-213">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-213">&lt;optional&gt;</span></span>|<span data-ttu-id="a10c1-214">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a10c1-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a10c1-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-215">Requirements</span></span>

|<span data-ttu-id="a10c1-216">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-216">Requirement</span></span>| <span data-ttu-id="a10c1-217">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-218">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-219">1,5</span><span class="sxs-lookup"><span data-stu-id="a10c1-219">1.5</span></span> |
|[<span data-ttu-id="a10c1-220">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-221">ReadItem</span></span> |
|[<span data-ttu-id="a10c1-222">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-223">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a10c1-224">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a10c1-224">Example</span></span>

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
}
```

---
---

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="a10c1-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="a10c1-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="a10c1-226">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="a10c1-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="a10c1-227">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a10c1-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a10c1-p106">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-230">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-230">Parameters</span></span>

|<span data-ttu-id="a10c1-231">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-231">Name</span></span>| <span data-ttu-id="a10c1-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-232">Type</span></span>| <span data-ttu-id="a10c1-233">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="a10c1-234">String</span><span class="sxs-lookup"><span data-stu-id="a10c1-234">String</span></span>|<span data-ttu-id="a10c1-235">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="a10c1-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="a10c1-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="a10c1-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="a10c1-237">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="a10c1-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a10c1-238">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-238">Requirements</span></span>

|<span data-ttu-id="a10c1-239">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-239">Requirement</span></span>| <span data-ttu-id="a10c1-240">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-241">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-242">1.3</span><span class="sxs-lookup"><span data-stu-id="a10c1-242">1.3</span></span>|
|[<span data-ttu-id="a10c1-243">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-244">Restrito</span><span class="sxs-lookup"><span data-stu-id="a10c1-244">Restricted</span></span>|
|[<span data-ttu-id="a10c1-245">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a10c1-247">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a10c1-247">Returns:</span></span>

<span data-ttu-id="a10c1-248">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="a10c1-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="a10c1-249">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a10c1-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="a10c1-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="a10c1-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="a10c1-251">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="a10c1-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="a10c1-p107">As datas e horas usadas por um aplicativo de email para o Outlook ou o Outlook Web App podem usar fusos horários diferentes. O Outlook usa o fuso horário do computador cliente; o Outlook Web App usa o fuso horário definido na Centro de administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p107">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="a10c1-p108">Se o aplicativo de email estiver sendo executado no Outlook, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador do cliente. Se o aplicativo de email estiver sendo executado no Outlook Web App, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p108">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-257">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-257">Parameters</span></span>

|<span data-ttu-id="a10c1-258">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-258">Name</span></span>| <span data-ttu-id="a10c1-259">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-259">Type</span></span>| <span data-ttu-id="a10c1-260">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="a10c1-261">Date</span><span class="sxs-lookup"><span data-stu-id="a10c1-261">Date</span></span>|<span data-ttu-id="a10c1-262">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="a10c1-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a10c1-263">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-263">Requirements</span></span>

|<span data-ttu-id="a10c1-264">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-264">Requirement</span></span>| <span data-ttu-id="a10c1-265">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-266">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-267">1.0</span><span class="sxs-lookup"><span data-stu-id="a10c1-267">1.0</span></span>|
|[<span data-ttu-id="a10c1-268">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-269">ReadItem</span></span>|
|[<span data-ttu-id="a10c1-270">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-271">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a10c1-272">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a10c1-272">Returns:</span></span>

<span data-ttu-id="a10c1-273">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="a10c1-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

---
---

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="a10c1-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="a10c1-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="a10c1-275">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="a10c1-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="a10c1-276">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a10c1-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a10c1-p109">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-279">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-279">Parameters</span></span>

|<span data-ttu-id="a10c1-280">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-280">Name</span></span>| <span data-ttu-id="a10c1-281">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-281">Type</span></span>| <span data-ttu-id="a10c1-282">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="a10c1-283">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a10c1-283">String</span></span>|<span data-ttu-id="a10c1-284">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="a10c1-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="a10c1-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="a10c1-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="a10c1-286">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="a10c1-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a10c1-287">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-287">Requirements</span></span>

|<span data-ttu-id="a10c1-288">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-288">Requirement</span></span>| <span data-ttu-id="a10c1-289">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-290">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-291">1.3</span><span class="sxs-lookup"><span data-stu-id="a10c1-291">1.3</span></span>|
|[<span data-ttu-id="a10c1-292">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-293">Restrito</span><span class="sxs-lookup"><span data-stu-id="a10c1-293">Restricted</span></span>|
|[<span data-ttu-id="a10c1-294">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-295">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a10c1-296">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a10c1-296">Returns:</span></span>

<span data-ttu-id="a10c1-297">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="a10c1-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="a10c1-298">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a10c1-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="a10c1-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="a10c1-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="a10c1-300">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="a10c1-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="a10c1-301">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="a10c1-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-302">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-302">Parameters</span></span>

|<span data-ttu-id="a10c1-303">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-303">Name</span></span>| <span data-ttu-id="a10c1-304">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-304">Type</span></span>| <span data-ttu-id="a10c1-305">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="a10c1-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="a10c1-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="a10c1-307">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="a10c1-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a10c1-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-308">Requirements</span></span>

|<span data-ttu-id="a10c1-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-309">Requirement</span></span>| <span data-ttu-id="a10c1-310">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-312">1.0</span><span class="sxs-lookup"><span data-stu-id="a10c1-312">1.0</span></span>|
|[<span data-ttu-id="a10c1-313">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-314">ReadItem</span></span>|
|[<span data-ttu-id="a10c1-315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-316">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a10c1-317">Retorna:</span><span class="sxs-lookup"><span data-stu-id="a10c1-317">Returns:</span></span>

<span data-ttu-id="a10c1-318">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="a10c1-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="a10c1-319">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="a10c1-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a10c1-320">Date</span><span class="sxs-lookup"><span data-stu-id="a10c1-320">Date</span></span></dd>

</dl>

---
---

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="a10c1-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="a10c1-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="a10c1-322">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="a10c1-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a10c1-323">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a10c1-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a10c1-324">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="a10c1-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="a10c1-p110">No Outlook para Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir uma instância da série. Isso ocorre porque no Outlook para Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p110">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="a10c1-327">No Outlook Web App, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a10c1-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="a10c1-328">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="a10c1-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-329">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-329">Parameters</span></span>

|<span data-ttu-id="a10c1-330">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-330">Name</span></span>| <span data-ttu-id="a10c1-331">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-331">Type</span></span>| <span data-ttu-id="a10c1-332">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="a10c1-333">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a10c1-333">String</span></span>|<span data-ttu-id="a10c1-334">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="a10c1-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a10c1-335">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-335">Requirements</span></span>

|<span data-ttu-id="a10c1-336">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-336">Requirement</span></span>| <span data-ttu-id="a10c1-337">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-338">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-339">1.0</span><span class="sxs-lookup"><span data-stu-id="a10c1-339">1.0</span></span>|
|[<span data-ttu-id="a10c1-340">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-341">ReadItem</span></span>|
|[<span data-ttu-id="a10c1-342">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-343">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a10c1-344">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a10c1-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

####  <a name="displaymessageformitemid"></a><span data-ttu-id="a10c1-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="a10c1-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="a10c1-346">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="a10c1-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="a10c1-347">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a10c1-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a10c1-348">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="a10c1-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="a10c1-349">No Outlook Web App, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a10c1-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="a10c1-350">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="a10c1-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="a10c1-p111">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-353">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-353">Parameters</span></span>

|<span data-ttu-id="a10c1-354">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-354">Name</span></span>| <span data-ttu-id="a10c1-355">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-355">Type</span></span>| <span data-ttu-id="a10c1-356">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="a10c1-357">String</span><span class="sxs-lookup"><span data-stu-id="a10c1-357">String</span></span>|<span data-ttu-id="a10c1-358">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="a10c1-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a10c1-359">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-359">Requirements</span></span>

|<span data-ttu-id="a10c1-360">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-360">Requirement</span></span>| <span data-ttu-id="a10c1-361">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-362">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-363">1.0</span><span class="sxs-lookup"><span data-stu-id="a10c1-363">1.0</span></span>|
|[<span data-ttu-id="a10c1-364">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-365">ReadItem</span></span>|
|[<span data-ttu-id="a10c1-366">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a10c1-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-367">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a10c1-368">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a10c1-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="a10c1-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="a10c1-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="a10c1-370">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="a10c1-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a10c1-371">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="a10c1-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a10c1-p112">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="a10c1-p113">No Outlook Web App e no OWA para Dispositivos, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p113">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="a10c1-p114">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="a10c1-379">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="a10c1-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-380">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="a10c1-381">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="a10c1-381">All parameters are optional.</span></span>

|<span data-ttu-id="a10c1-382">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-382">Name</span></span>| <span data-ttu-id="a10c1-383">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-383">Type</span></span>| <span data-ttu-id="a10c1-384">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="a10c1-385">Object</span><span class="sxs-lookup"><span data-stu-id="a10c1-385">Object</span></span> | <span data-ttu-id="a10c1-386">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="a10c1-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="a10c1-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="a10c1-p115">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="a10c1-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="a10c1-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="a10c1-393">Data</span><span class="sxs-lookup"><span data-stu-id="a10c1-393">Date</span></span> | <span data-ttu-id="a10c1-394">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="a10c1-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="a10c1-395">Data</span><span class="sxs-lookup"><span data-stu-id="a10c1-395">Date</span></span> | <span data-ttu-id="a10c1-396">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="a10c1-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="a10c1-397">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a10c1-397">String</span></span> | <span data-ttu-id="a10c1-p117">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="a10c1-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="a10c1-p118">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="a10c1-403">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a10c1-403">String</span></span> | <span data-ttu-id="a10c1-p119">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="a10c1-406">String</span><span class="sxs-lookup"><span data-stu-id="a10c1-406">String</span></span> | <span data-ttu-id="a10c1-p120">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a10c1-409">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-409">Requirements</span></span>

|<span data-ttu-id="a10c1-410">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-410">Requirement</span></span>| <span data-ttu-id="a10c1-411">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-412">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-413">1.0</span><span class="sxs-lookup"><span data-stu-id="a10c1-413">1.0</span></span>|
|[<span data-ttu-id="a10c1-414">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-415">ReadItem</span></span>|
|[<span data-ttu-id="a10c1-416">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-417">Read</span><span class="sxs-lookup"><span data-stu-id="a10c1-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a10c1-418">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a10c1-418">Example</span></span>

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

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="a10c1-419">displayNewMessageForm (parâmetros)</span><span class="sxs-lookup"><span data-stu-id="a10c1-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="a10c1-420">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="a10c1-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="a10c1-421">O `displayNewMessageForm` método abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="a10c1-421">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="a10c1-422">Se os parâmetros forem especificados, os campos de formulário da mensagem serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="a10c1-422">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="a10c1-423">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="a10c1-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-424">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="a10c1-425">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="a10c1-425">All parameters are optional.</span></span>

|<span data-ttu-id="a10c1-426">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-426">Name</span></span>| <span data-ttu-id="a10c1-427">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-427">Type</span></span>| <span data-ttu-id="a10c1-428">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="a10c1-429">Object</span><span class="sxs-lookup"><span data-stu-id="a10c1-429">Object</span></span> | <span data-ttu-id="a10c1-430">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="a10c1-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="a10c1-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="a10c1-432">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha para.</span><span class="sxs-lookup"><span data-stu-id="a10c1-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="a10c1-433">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="a10c1-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="a10c1-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="a10c1-435">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha CC.</span><span class="sxs-lookup"><span data-stu-id="a10c1-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="a10c1-436">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="a10c1-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="a10c1-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="a10c1-438">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="a10c1-438">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="a10c1-439">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="a10c1-439">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="a10c1-440">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a10c1-440">String</span></span> | <span data-ttu-id="a10c1-441">Uma cadeia de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="a10c1-441">A string containing the subject of the message.</span></span> <span data-ttu-id="a10c1-442">A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a10c1-442">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="a10c1-443">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a10c1-443">String</span></span> | <span data-ttu-id="a10c1-444">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="a10c1-444">The HTML body of the message.</span></span> <span data-ttu-id="a10c1-445">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="a10c1-445">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="a10c1-446">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="a10c1-447">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="a10c1-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="a10c1-448">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a10c1-448">String</span></span> | <span data-ttu-id="a10c1-p127">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="a10c1-451">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a10c1-451">String</span></span> | <span data-ttu-id="a10c1-452">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="a10c1-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="a10c1-453">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a10c1-453">String</span></span> | <span data-ttu-id="a10c1-p128">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="a10c1-456">Booliano</span><span class="sxs-lookup"><span data-stu-id="a10c1-456">Boolean</span></span> | <span data-ttu-id="a10c1-p129">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="a10c1-459">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a10c1-459">String</span></span> | <span data-ttu-id="a10c1-460">Usado somente se `type` estiver definido como `item`.</span><span class="sxs-lookup"><span data-stu-id="a10c1-460">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="a10c1-461">A ID do item do EWS do email existente que você deseja anexar à nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="a10c1-461">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="a10c1-462">Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="a10c1-462">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="a10c1-463">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-463">Requirements</span></span>

|<span data-ttu-id="a10c1-464">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-464">Requirement</span></span>| <span data-ttu-id="a10c1-465">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-466">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-467">1.6</span><span class="sxs-lookup"><span data-stu-id="a10c1-467">1.6</span></span> |
|[<span data-ttu-id="a10c1-468">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-469">ReadItem</span></span>|
|[<span data-ttu-id="a10c1-470">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-471">Read</span><span class="sxs-lookup"><span data-stu-id="a10c1-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a10c1-472">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a10c1-472">Example</span></span>

```javascript
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to,
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

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="a10c1-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="a10c1-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="a10c1-474">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="a10c1-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="a10c1-p131">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="a10c1-477">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="a10c1-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="a10c1-478">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="a10c1-478">**REST Tokens**</span></span>

<span data-ttu-id="a10c1-p132">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="a10c1-482">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="a10c1-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="a10c1-483">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="a10c1-483">**EWS Tokens**</span></span>

<span data-ttu-id="a10c1-p133">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="a10c1-486">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="a10c1-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-487">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-487">Parameters</span></span>

|<span data-ttu-id="a10c1-488">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-488">Name</span></span>| <span data-ttu-id="a10c1-489">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-489">Type</span></span>| <span data-ttu-id="a10c1-490">Atributos</span><span class="sxs-lookup"><span data-stu-id="a10c1-490">Attributes</span></span>| <span data-ttu-id="a10c1-491">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="a10c1-492">Objeto</span><span class="sxs-lookup"><span data-stu-id="a10c1-492">Object</span></span> | <span data-ttu-id="a10c1-493">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-493">&lt;optional&gt;</span></span> | <span data-ttu-id="a10c1-494">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a10c1-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="a10c1-495">Booliano</span><span class="sxs-lookup"><span data-stu-id="a10c1-495">Boolean</span></span> |  <span data-ttu-id="a10c1-496">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-496">&lt;optional&gt;</span></span> | <span data-ttu-id="a10c1-p134">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a10c1-499">Objeto</span><span class="sxs-lookup"><span data-stu-id="a10c1-499">Object</span></span> |  <span data-ttu-id="a10c1-500">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-500">&lt;optional&gt;</span></span> | <span data-ttu-id="a10c1-501">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="a10c1-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="a10c1-502">function</span><span class="sxs-lookup"><span data-stu-id="a10c1-502">function</span></span>||<span data-ttu-id="a10c1-p135">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a10c1-505">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-505">Requirements</span></span>

|<span data-ttu-id="a10c1-506">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-506">Requirement</span></span>| <span data-ttu-id="a10c1-507">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-508">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-509">1,5</span><span class="sxs-lookup"><span data-stu-id="a10c1-509">1.5</span></span> |
|[<span data-ttu-id="a10c1-510">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-511">ReadItem</span></span>|
|[<span data-ttu-id="a10c1-512">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-513">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="a10c1-513">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="a10c1-514">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a10c1-514">Example</span></span>

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

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="a10c1-515">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a10c1-515">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="a10c1-516">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="a10c1-516">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="a10c1-p136">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="a10c1-p137">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="a10c1-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="a10c1-522">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="a10c1-522">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="a10c1-p138">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-525">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-525">Parameters</span></span>

|<span data-ttu-id="a10c1-526">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-526">Name</span></span>| <span data-ttu-id="a10c1-527">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-527">Type</span></span>| <span data-ttu-id="a10c1-528">Atributos</span><span class="sxs-lookup"><span data-stu-id="a10c1-528">Attributes</span></span>| <span data-ttu-id="a10c1-529">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-529">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="a10c1-530">function</span><span class="sxs-lookup"><span data-stu-id="a10c1-530">function</span></span>||<span data-ttu-id="a10c1-p139">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="a10c1-533">Objeto</span><span class="sxs-lookup"><span data-stu-id="a10c1-533">Object</span></span>| <span data-ttu-id="a10c1-534">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-534">&lt;optional&gt;</span></span>|<span data-ttu-id="a10c1-535">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="a10c1-535">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a10c1-536">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-536">Requirements</span></span>

|<span data-ttu-id="a10c1-537">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-537">Requirement</span></span>| <span data-ttu-id="a10c1-538">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-539">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-540">1.3</span><span class="sxs-lookup"><span data-stu-id="a10c1-540">1.3</span></span>|
|[<span data-ttu-id="a10c1-541">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-542">ReadItem</span></span>|
|[<span data-ttu-id="a10c1-543">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-544">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="a10c1-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="a10c1-545">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a10c1-545">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="a10c1-546">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a10c1-546">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="a10c1-547">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="a10c1-547">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="a10c1-548">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="a10c1-548">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-549">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-549">Parameters</span></span>

|<span data-ttu-id="a10c1-550">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-550">Name</span></span>| <span data-ttu-id="a10c1-551">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-551">Type</span></span>| <span data-ttu-id="a10c1-552">Atributos</span><span class="sxs-lookup"><span data-stu-id="a10c1-552">Attributes</span></span>| <span data-ttu-id="a10c1-553">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-553">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="a10c1-554">function</span><span class="sxs-lookup"><span data-stu-id="a10c1-554">function</span></span>||<span data-ttu-id="a10c1-555">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a10c1-555">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a10c1-556">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a10c1-556">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="a10c1-557">Object</span><span class="sxs-lookup"><span data-stu-id="a10c1-557">Object</span></span>| <span data-ttu-id="a10c1-558">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-558">&lt;optional&gt;</span></span>|<span data-ttu-id="a10c1-559">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="a10c1-559">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a10c1-560">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-560">Requirements</span></span>

|<span data-ttu-id="a10c1-561">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-561">Requirement</span></span>| <span data-ttu-id="a10c1-562">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-563">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-564">1.0</span><span class="sxs-lookup"><span data-stu-id="a10c1-564">1.0</span></span>|
|[<span data-ttu-id="a10c1-565">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-566">ReadItem</span></span>|
|[<span data-ttu-id="a10c1-567">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="a10c1-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-568">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-568">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a10c1-569">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a10c1-569">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="a10c1-570">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a10c1-570">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="a10c1-571">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="a10c1-571">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="a10c1-572">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="a10c1-572">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="a10c1-573">No Outlook para iOS ou no Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="a10c1-573">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="a10c1-574">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="a10c1-574">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="a10c1-575">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="a10c1-575">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="a10c1-576">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="a10c1-576">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="a10c1-577">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="a10c1-577">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="a10c1-578">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="a10c1-578">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="a10c1-579">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="a10c1-579">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="a10c1-p141">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="a10c1-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="a10c1-582">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="a10c1-582">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="a10c1-583">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="a10c1-583">Version differences</span></span>

<span data-ttu-id="a10c1-584">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="a10c1-584">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="a10c1-p142">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="a10c1-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-588">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-588">Parameters</span></span>

|<span data-ttu-id="a10c1-589">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-589">Name</span></span>| <span data-ttu-id="a10c1-590">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-590">Type</span></span>| <span data-ttu-id="a10c1-591">Atributos</span><span class="sxs-lookup"><span data-stu-id="a10c1-591">Attributes</span></span>| <span data-ttu-id="a10c1-592">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-592">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="a10c1-593">String</span><span class="sxs-lookup"><span data-stu-id="a10c1-593">String</span></span>||<span data-ttu-id="a10c1-594">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="a10c1-594">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="a10c1-595">function</span><span class="sxs-lookup"><span data-stu-id="a10c1-595">function</span></span>||<span data-ttu-id="a10c1-596">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a10c1-596">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a10c1-597">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a10c1-597">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="a10c1-598">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="a10c1-598">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="a10c1-599">Objeto</span><span class="sxs-lookup"><span data-stu-id="a10c1-599">Object</span></span>| <span data-ttu-id="a10c1-600">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-600">&lt;optional&gt;</span></span>|<span data-ttu-id="a10c1-601">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="a10c1-601">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a10c1-602">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-602">Requirements</span></span>

|<span data-ttu-id="a10c1-603">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-603">Requirement</span></span>| <span data-ttu-id="a10c1-604">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-605">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-606">1.0</span><span class="sxs-lookup"><span data-stu-id="a10c1-606">1.0</span></span>|
|[<span data-ttu-id="a10c1-607">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-608">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="a10c1-608">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="a10c1-609">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-610">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-610">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a10c1-611">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a10c1-611">Example</span></span>

<span data-ttu-id="a10c1-612">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="a10c1-612">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

---
---

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="a10c1-613">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a10c1-613">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="a10c1-614">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="a10c1-614">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="a10c1-615">Atualmente, os tipos de eventos com `Office.EventType.ItemChanged` suporte `Office.EventType.OfficeThemeChanged`são e.</span><span class="sxs-lookup"><span data-stu-id="a10c1-615">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a10c1-616">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="a10c1-616">Parameters</span></span>

| <span data-ttu-id="a10c1-617">Nome</span><span class="sxs-lookup"><span data-stu-id="a10c1-617">Name</span></span> | <span data-ttu-id="a10c1-618">Tipo</span><span class="sxs-lookup"><span data-stu-id="a10c1-618">Type</span></span> | <span data-ttu-id="a10c1-619">Atributos</span><span class="sxs-lookup"><span data-stu-id="a10c1-619">Attributes</span></span> | <span data-ttu-id="a10c1-620">Descrição</span><span class="sxs-lookup"><span data-stu-id="a10c1-620">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="a10c1-621">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="a10c1-621">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="a10c1-622">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="a10c1-622">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="a10c1-623">Objeto</span><span class="sxs-lookup"><span data-stu-id="a10c1-623">Object</span></span> | <span data-ttu-id="a10c1-624">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-624">&lt;optional&gt;</span></span> | <span data-ttu-id="a10c1-625">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="a10c1-625">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="a10c1-626">Objeto</span><span class="sxs-lookup"><span data-stu-id="a10c1-626">Object</span></span> | <span data-ttu-id="a10c1-627">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-627">&lt;optional&gt;</span></span> | <span data-ttu-id="a10c1-628">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="a10c1-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="a10c1-629">function</span><span class="sxs-lookup"><span data-stu-id="a10c1-629">function</span></span>| <span data-ttu-id="a10c1-630">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="a10c1-630">&lt;optional&gt;</span></span>|<span data-ttu-id="a10c1-631">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a10c1-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a10c1-632">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a10c1-632">Requirements</span></span>

|<span data-ttu-id="a10c1-633">Requisito</span><span class="sxs-lookup"><span data-stu-id="a10c1-633">Requirement</span></span>| <span data-ttu-id="a10c1-634">Valor</span><span class="sxs-lookup"><span data-stu-id="a10c1-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="a10c1-635">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a10c1-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a10c1-636">1,5</span><span class="sxs-lookup"><span data-stu-id="a10c1-636">1.5</span></span> |
|[<span data-ttu-id="a10c1-637">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a10c1-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a10c1-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a10c1-638">ReadItem</span></span> |
|[<span data-ttu-id="a10c1-639">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a10c1-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a10c1-640">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a10c1-640">Compose or Read</span></span>|
