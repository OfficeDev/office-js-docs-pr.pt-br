---
title: Office. Context. Mailbox – conjunto de requisitos 1,8
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: 908eff7b34e63b62fbe250f1a6f810be69b17627
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629213"
---
# <a name="mailbox"></a><span data-ttu-id="d7935-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="d7935-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="d7935-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="d7935-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="d7935-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="d7935-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7935-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-105">Requirements</span></span>

|<span data-ttu-id="d7935-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-106">Requirement</span></span>| <span data-ttu-id="d7935-107">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d7935-109">1.0</span></span>|
|[<span data-ttu-id="d7935-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="d7935-111">Restricted</span></span>|
|[<span data-ttu-id="d7935-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d7935-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d7935-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="d7935-114">Members and methods</span></span>

| <span data-ttu-id="d7935-115">Membro</span><span class="sxs-lookup"><span data-stu-id="d7935-115">Member</span></span> | <span data-ttu-id="d7935-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d7935-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="d7935-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="d7935-118">Membro</span><span class="sxs-lookup"><span data-stu-id="d7935-118">Member</span></span> |
| [<span data-ttu-id="d7935-119">Nova mastercategories</span><span class="sxs-lookup"><span data-stu-id="d7935-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="d7935-120">Membro</span><span class="sxs-lookup"><span data-stu-id="d7935-120">Member</span></span> |
| [<span data-ttu-id="d7935-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="d7935-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="d7935-122">Membro</span><span class="sxs-lookup"><span data-stu-id="d7935-122">Member</span></span> |
| [<span data-ttu-id="d7935-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d7935-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d7935-124">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-124">Method</span></span> |
| [<span data-ttu-id="d7935-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="d7935-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="d7935-126">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-126">Method</span></span> |
| [<span data-ttu-id="d7935-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d7935-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="d7935-128">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-128">Method</span></span> |
| [<span data-ttu-id="d7935-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="d7935-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="d7935-130">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-130">Method</span></span> |
| [<span data-ttu-id="d7935-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="d7935-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="d7935-132">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-132">Method</span></span> |
| [<span data-ttu-id="d7935-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d7935-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="d7935-134">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-134">Method</span></span> |
| [<span data-ttu-id="d7935-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="d7935-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="d7935-136">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-136">Method</span></span> |
| [<span data-ttu-id="d7935-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d7935-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="d7935-138">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-138">Method</span></span> |
| [<span data-ttu-id="d7935-139">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="d7935-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="d7935-140">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-140">Method</span></span> |
| [<span data-ttu-id="d7935-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d7935-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="d7935-142">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-142">Method</span></span> |
| [<span data-ttu-id="d7935-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d7935-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="d7935-144">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-144">Method</span></span> |
| [<span data-ttu-id="d7935-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d7935-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="d7935-146">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-146">Method</span></span> |
| [<span data-ttu-id="d7935-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="d7935-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="d7935-148">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-148">Method</span></span> |
| [<span data-ttu-id="d7935-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d7935-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="d7935-150">Método</span><span class="sxs-lookup"><span data-stu-id="d7935-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d7935-151">Namespaces</span><span class="sxs-lookup"><span data-stu-id="d7935-151">Namespaces</span></span>

<span data-ttu-id="d7935-152">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d7935-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="d7935-153">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d7935-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="d7935-154">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d7935-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="d7935-155">Members</span><span class="sxs-lookup"><span data-stu-id="d7935-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="d7935-156">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="d7935-156">ewsUrl: String</span></span>

<span data-ttu-id="d7935-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="d7935-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d7935-159">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d7935-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d7935-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d7935-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d7935-162">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="d7935-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="d7935-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="d7935-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d7935-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-165">Type</span></span>

*   <span data-ttu-id="d7935-166">String</span><span class="sxs-lookup"><span data-stu-id="d7935-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7935-167">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-167">Requirements</span></span>

|<span data-ttu-id="d7935-168">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-168">Requirement</span></span>| <span data-ttu-id="d7935-169">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-170">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-171">1.0</span><span class="sxs-lookup"><span data-stu-id="d7935-171">1.0</span></span>|
|[<span data-ttu-id="d7935-172">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-173">ReadItem</span></span>|
|[<span data-ttu-id="d7935-174">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d7935-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-175">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-175">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategoriesviewoutlook-js-18"></a><span data-ttu-id="d7935-176">Nova mastercategories: [nova mastercategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="d7935-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)</span></span>

<span data-ttu-id="d7935-177">Obtém um objeto que fornece métodos para gerenciar a lista mestra de categorias nesta caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="d7935-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="d7935-178">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d7935-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d7935-179">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-179">Type</span></span>

*   [<span data-ttu-id="d7935-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="d7935-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="d7935-181">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-181">Requirements</span></span>

|<span data-ttu-id="d7935-182">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-182">Requirement</span></span>| <span data-ttu-id="d7935-183">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-184">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-185">1,8</span><span class="sxs-lookup"><span data-stu-id="d7935-185">1.8</span></span> |
|[<span data-ttu-id="d7935-186">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="d7935-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="d7935-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d7935-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="d7935-190">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-190">Example</span></span>

<span data-ttu-id="d7935-191">Este exemplo obtém a lista mestra de categorias para esta caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="d7935-191">This example gets the categories master list for this mailbox.</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="d7935-192">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="d7935-192">restUrl: String</span></span>

<span data-ttu-id="d7935-193">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="d7935-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="d7935-194">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="d7935-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="d7935-195">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-195">Type</span></span>

*   <span data-ttu-id="d7935-196">String</span><span class="sxs-lookup"><span data-stu-id="d7935-196">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7935-197">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-197">Requirements</span></span>

|<span data-ttu-id="d7935-198">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-198">Requirement</span></span>| <span data-ttu-id="d7935-199">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-200">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-201">1,5</span><span class="sxs-lookup"><span data-stu-id="d7935-201">1.5</span></span> |
|[<span data-ttu-id="d7935-202">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-203">ReadItem</span></span>|
|[<span data-ttu-id="d7935-204">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d7935-204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-205">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-205">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d7935-206">Métodos</span><span class="sxs-lookup"><span data-stu-id="d7935-206">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d7935-207">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d7935-207">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d7935-208">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="d7935-208">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d7935-209">Atualmente, os tipos de eventos com `Office.EventType.ItemChanged` suporte `Office.EventType.OfficeThemeChanged`são e.</span><span class="sxs-lookup"><span data-stu-id="d7935-209">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-210">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-210">Parameters</span></span>

| <span data-ttu-id="d7935-211">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-211">Name</span></span> | <span data-ttu-id="d7935-212">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-212">Type</span></span> | <span data-ttu-id="d7935-213">Atributos</span><span class="sxs-lookup"><span data-stu-id="d7935-213">Attributes</span></span> | <span data-ttu-id="d7935-214">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-214">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d7935-215">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d7935-215">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d7935-216">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="d7935-216">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d7935-217">Função</span><span class="sxs-lookup"><span data-stu-id="d7935-217">Function</span></span> || <span data-ttu-id="d7935-p104">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="d7935-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d7935-221">Objeto</span><span class="sxs-lookup"><span data-stu-id="d7935-221">Object</span></span> | <span data-ttu-id="d7935-222">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-222">&lt;optional&gt;</span></span> | <span data-ttu-id="d7935-223">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="d7935-223">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d7935-224">Objeto</span><span class="sxs-lookup"><span data-stu-id="d7935-224">Object</span></span> | <span data-ttu-id="d7935-225">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-225">&lt;optional&gt;</span></span> | <span data-ttu-id="d7935-226">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d7935-226">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d7935-227">function</span><span class="sxs-lookup"><span data-stu-id="d7935-227">function</span></span>| <span data-ttu-id="d7935-228">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-228">&lt;optional&gt;</span></span>|<span data-ttu-id="d7935-229">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d7935-229">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7935-230">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-230">Requirements</span></span>

|<span data-ttu-id="d7935-231">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-231">Requirement</span></span>| <span data-ttu-id="d7935-232">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-233">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-234">1,5</span><span class="sxs-lookup"><span data-stu-id="d7935-234">1.5</span></span> |
|[<span data-ttu-id="d7935-235">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-236">ReadItem</span></span> |
|[<span data-ttu-id="d7935-237">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d7935-237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-238">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7935-239">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-239">Example</span></span>

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
}
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="d7935-240">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d7935-240">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d7935-241">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="d7935-241">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="d7935-242">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d7935-242">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d7935-p105">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="d7935-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-245">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-245">Parameters</span></span>

|<span data-ttu-id="d7935-246">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-246">Name</span></span>| <span data-ttu-id="d7935-247">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-247">Type</span></span>| <span data-ttu-id="d7935-248">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-248">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d7935-249">String</span><span class="sxs-lookup"><span data-stu-id="d7935-249">String</span></span>|<span data-ttu-id="d7935-250">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="d7935-250">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="d7935-251">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d7935-251">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.8)|<span data-ttu-id="d7935-252">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="d7935-252">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7935-253">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-253">Requirements</span></span>

|<span data-ttu-id="d7935-254">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-254">Requirement</span></span>| <span data-ttu-id="d7935-255">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-256">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-256">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-257">1.3</span><span class="sxs-lookup"><span data-stu-id="d7935-257">1.3</span></span>|
|[<span data-ttu-id="d7935-258">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-258">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-259">Restrito</span><span class="sxs-lookup"><span data-stu-id="d7935-259">Restricted</span></span>|
|[<span data-ttu-id="d7935-260">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d7935-260">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-261">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-261">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7935-262">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d7935-262">Returns:</span></span>

<span data-ttu-id="d7935-263">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="d7935-263">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d7935-264">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-264">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-18"></a><span data-ttu-id="d7935-265">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="d7935-265">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="d7935-266">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="d7935-266">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="d7935-p106">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para as datas e horas. O Outlook em uma área de trabalho usa o fuso horário do computador cliente; o Outlook na Web usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="d7935-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="d7935-p107">Se o aplicativo de email estiver sendo executado no Outlook em um cliente da área de trabalho, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador cliente. Se o aplicativo de email estiver sendo executado no Outlook na Web, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="d7935-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-272">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-272">Parameters</span></span>

|<span data-ttu-id="d7935-273">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-273">Name</span></span>| <span data-ttu-id="d7935-274">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-274">Type</span></span>| <span data-ttu-id="d7935-275">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-275">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="d7935-276">Date</span><span class="sxs-lookup"><span data-stu-id="d7935-276">Date</span></span>|<span data-ttu-id="d7935-277">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="d7935-277">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7935-278">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-278">Requirements</span></span>

|<span data-ttu-id="d7935-279">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-279">Requirement</span></span>| <span data-ttu-id="d7935-280">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-280">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-281">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-281">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-282">1.0</span><span class="sxs-lookup"><span data-stu-id="d7935-282">1.0</span></span>|
|[<span data-ttu-id="d7935-283">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-283">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-284">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-284">ReadItem</span></span>|
|[<span data-ttu-id="d7935-285">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d7935-285">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-286">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-286">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7935-287">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d7935-287">Returns:</span></span>

<span data-ttu-id="d7935-288">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="d7935-288">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="d7935-289">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d7935-289">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d7935-290">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="d7935-290">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="d7935-291">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d7935-291">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d7935-p108">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="d7935-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-294">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-294">Parameters</span></span>

|<span data-ttu-id="d7935-295">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-295">Name</span></span>| <span data-ttu-id="d7935-296">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-296">Type</span></span>| <span data-ttu-id="d7935-297">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-297">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d7935-298">String</span><span class="sxs-lookup"><span data-stu-id="d7935-298">String</span></span>|<span data-ttu-id="d7935-299">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="d7935-299">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="d7935-300">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d7935-300">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.8)|<span data-ttu-id="d7935-301">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="d7935-301">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7935-302">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-302">Requirements</span></span>

|<span data-ttu-id="d7935-303">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-303">Requirement</span></span>| <span data-ttu-id="d7935-304">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-305">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-306">1.3</span><span class="sxs-lookup"><span data-stu-id="d7935-306">1.3</span></span>|
|[<span data-ttu-id="d7935-307">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-308">Restrito</span><span class="sxs-lookup"><span data-stu-id="d7935-308">Restricted</span></span>|
|[<span data-ttu-id="d7935-309">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d7935-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-310">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-310">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7935-311">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d7935-311">Returns:</span></span>

<span data-ttu-id="d7935-312">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="d7935-312">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d7935-313">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-313">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="d7935-314">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="d7935-314">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="d7935-315">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="d7935-315">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="d7935-316">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="d7935-316">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-317">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-317">Parameters</span></span>

|<span data-ttu-id="d7935-318">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-318">Name</span></span>| <span data-ttu-id="d7935-319">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-319">Type</span></span>| <span data-ttu-id="d7935-320">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-320">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="d7935-321">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d7935-321">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)|<span data-ttu-id="d7935-322">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="d7935-322">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7935-323">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-323">Requirements</span></span>

|<span data-ttu-id="d7935-324">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-324">Requirement</span></span>| <span data-ttu-id="d7935-325">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-326">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-327">1.0</span><span class="sxs-lookup"><span data-stu-id="d7935-327">1.0</span></span>|
|[<span data-ttu-id="d7935-328">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-329">ReadItem</span></span>|
|[<span data-ttu-id="d7935-330">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d7935-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-331">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-331">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7935-332">Retorna:</span><span class="sxs-lookup"><span data-stu-id="d7935-332">Returns:</span></span>

<span data-ttu-id="d7935-333">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="d7935-333">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="d7935-334">Tipo: Data</span><span class="sxs-lookup"><span data-stu-id="d7935-334">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="d7935-335">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-335">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="d7935-336">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d7935-336">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="d7935-337">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="d7935-337">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d7935-338">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d7935-338">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d7935-339">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="d7935-339">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d7935-p109">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir um instância da série. Isso ocorre porque no Outlook no Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="d7935-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="d7935-342">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="d7935-342">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="d7935-343">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="d7935-343">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-344">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-344">Parameters</span></span>

|<span data-ttu-id="d7935-345">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-345">Name</span></span>| <span data-ttu-id="d7935-346">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-346">Type</span></span>| <span data-ttu-id="d7935-347">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-347">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d7935-348">String</span><span class="sxs-lookup"><span data-stu-id="d7935-348">String</span></span>|<span data-ttu-id="d7935-349">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="d7935-349">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7935-350">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-350">Requirements</span></span>

|<span data-ttu-id="d7935-351">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-351">Requirement</span></span>| <span data-ttu-id="d7935-352">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-353">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-354">1.0</span><span class="sxs-lookup"><span data-stu-id="d7935-354">1.0</span></span>|
|[<span data-ttu-id="d7935-355">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-356">ReadItem</span></span>|
|[<span data-ttu-id="d7935-357">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d7935-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-358">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-358">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7935-359">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-359">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="d7935-360">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d7935-360">displayMessageForm(itemId)</span></span>

<span data-ttu-id="d7935-361">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="d7935-361">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="d7935-362">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d7935-362">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d7935-363">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="d7935-363">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d7935-364">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d7935-364">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="d7935-365">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="d7935-365">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="d7935-p110">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="d7935-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-368">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-368">Parameters</span></span>

|<span data-ttu-id="d7935-369">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-369">Name</span></span>| <span data-ttu-id="d7935-370">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-370">Type</span></span>| <span data-ttu-id="d7935-371">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-371">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d7935-372">String</span><span class="sxs-lookup"><span data-stu-id="d7935-372">String</span></span>|<span data-ttu-id="d7935-373">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="d7935-373">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7935-374">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-374">Requirements</span></span>

|<span data-ttu-id="d7935-375">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-375">Requirement</span></span>| <span data-ttu-id="d7935-376">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-377">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-378">1.0</span><span class="sxs-lookup"><span data-stu-id="d7935-378">1.0</span></span>|
|[<span data-ttu-id="d7935-379">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-380">ReadItem</span></span>|
|[<span data-ttu-id="d7935-381">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d7935-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-382">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-382">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7935-383">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-383">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="d7935-384">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="d7935-384">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="d7935-385">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="d7935-385">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d7935-386">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="d7935-386">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d7935-p111">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="d7935-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d7935-p112">No Outlook na Web e em dispositivos móveis, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="d7935-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="d7935-p113">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="d7935-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="d7935-394">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="d7935-394">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-395">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-395">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="d7935-396">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="d7935-396">All parameters are optional.</span></span>

|<span data-ttu-id="d7935-397">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-397">Name</span></span>| <span data-ttu-id="d7935-398">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-398">Type</span></span>| <span data-ttu-id="d7935-399">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-399">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d7935-400">Object</span><span class="sxs-lookup"><span data-stu-id="d7935-400">Object</span></span> | <span data-ttu-id="d7935-401">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="d7935-401">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="d7935-402">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-402">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="d7935-p114">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="d7935-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="d7935-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="d7935-p115">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="d7935-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="d7935-408">Data</span><span class="sxs-lookup"><span data-stu-id="d7935-408">Date</span></span> | <span data-ttu-id="d7935-409">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="d7935-409">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="d7935-410">Data</span><span class="sxs-lookup"><span data-stu-id="d7935-410">Date</span></span> | <span data-ttu-id="d7935-411">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="d7935-411">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="d7935-412">String</span><span class="sxs-lookup"><span data-stu-id="d7935-412">String</span></span> | <span data-ttu-id="d7935-p116">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d7935-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="d7935-415">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-415">Array.&lt;String&gt;</span></span> | <span data-ttu-id="d7935-p117">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="d7935-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d7935-418">String</span><span class="sxs-lookup"><span data-stu-id="d7935-418">String</span></span> | <span data-ttu-id="d7935-p118">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d7935-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="d7935-421">String</span><span class="sxs-lookup"><span data-stu-id="d7935-421">String</span></span> | <span data-ttu-id="d7935-p119">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d7935-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d7935-424">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-424">Requirements</span></span>

|<span data-ttu-id="d7935-425">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-425">Requirement</span></span>| <span data-ttu-id="d7935-426">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-426">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-427">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-427">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-428">1.0</span><span class="sxs-lookup"><span data-stu-id="d7935-428">1.0</span></span>|
|[<span data-ttu-id="d7935-429">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-429">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-430">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-430">ReadItem</span></span>|
|[<span data-ttu-id="d7935-431">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d7935-431">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-432">Read</span><span class="sxs-lookup"><span data-stu-id="d7935-432">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7935-433">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-433">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="d7935-434">displayNewMessageForm (parâmetros)</span><span class="sxs-lookup"><span data-stu-id="d7935-434">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="d7935-435">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="d7935-435">Displays a form for creating a new message.</span></span>

<span data-ttu-id="d7935-436">O `displayNewMessageForm` método abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="d7935-436">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="d7935-437">Se os parâmetros forem especificados, os campos de formulário da mensagem serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="d7935-437">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d7935-438">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="d7935-438">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-439">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-439">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="d7935-440">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="d7935-440">All parameters are optional.</span></span>

|<span data-ttu-id="d7935-441">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-441">Name</span></span>| <span data-ttu-id="d7935-442">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-442">Type</span></span>| <span data-ttu-id="d7935-443">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-443">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d7935-444">Objeto</span><span class="sxs-lookup"><span data-stu-id="d7935-444">Object</span></span> | <span data-ttu-id="d7935-445">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="d7935-445">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="d7935-446">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-446">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="d7935-447">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha para.</span><span class="sxs-lookup"><span data-stu-id="d7935-447">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="d7935-448">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="d7935-448">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="d7935-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="d7935-450">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha CC.</span><span class="sxs-lookup"><span data-stu-id="d7935-450">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="d7935-451">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="d7935-451">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="d7935-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="d7935-453">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="d7935-453">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="d7935-454">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="d7935-454">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d7935-455">String</span><span class="sxs-lookup"><span data-stu-id="d7935-455">String</span></span> | <span data-ttu-id="d7935-456">Uma cadeia de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="d7935-456">A string containing the subject of the message.</span></span> <span data-ttu-id="d7935-457">A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d7935-457">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="d7935-458">String</span><span class="sxs-lookup"><span data-stu-id="d7935-458">String</span></span> | <span data-ttu-id="d7935-459">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="d7935-459">The HTML body of the message.</span></span> <span data-ttu-id="d7935-460">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="d7935-460">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="d7935-461">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-461">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d7935-462">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="d7935-462">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="d7935-463">String</span><span class="sxs-lookup"><span data-stu-id="d7935-463">String</span></span> | <span data-ttu-id="d7935-p126">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="d7935-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="d7935-466">String</span><span class="sxs-lookup"><span data-stu-id="d7935-466">String</span></span> | <span data-ttu-id="d7935-467">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="d7935-467">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="d7935-468">String</span><span class="sxs-lookup"><span data-stu-id="d7935-468">String</span></span> | <span data-ttu-id="d7935-p127">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="d7935-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="d7935-471">Booliano</span><span class="sxs-lookup"><span data-stu-id="d7935-471">Boolean</span></span> | <span data-ttu-id="d7935-p128">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="d7935-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="d7935-474">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d7935-474">String</span></span> | <span data-ttu-id="d7935-475">Usado somente se `type` estiver definido como `item`.</span><span class="sxs-lookup"><span data-stu-id="d7935-475">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="d7935-476">A ID do item do EWS do email existente que você deseja anexar à nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="d7935-476">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="d7935-477">Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="d7935-477">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="d7935-478">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-478">Requirements</span></span>

|<span data-ttu-id="d7935-479">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-479">Requirement</span></span>| <span data-ttu-id="d7935-480">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-481">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-482">1.6</span><span class="sxs-lookup"><span data-stu-id="d7935-482">1.6</span></span> |
|[<span data-ttu-id="d7935-483">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-484">ReadItem</span></span>|
|[<span data-ttu-id="d7935-485">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d7935-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-486">Read</span><span class="sxs-lookup"><span data-stu-id="d7935-486">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7935-487">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-487">Example</span></span>

```js
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

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="d7935-488">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d7935-488">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="d7935-489">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="d7935-489">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="d7935-p130">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="d7935-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="d7935-492">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="d7935-492">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="d7935-493">Chamar o método `getCallbackTokenAsync` no modo de leitura requer um nível de permissão mínimo de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="d7935-493">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="d7935-494">Chamar `getCallbackTokenAsync` no modo redigir exige que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="d7935-494">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="d7935-495">O método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="d7935-495">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="d7935-496">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="d7935-496">**REST Tokens**</span></span>

<span data-ttu-id="d7935-p132">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="d7935-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="d7935-500">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="d7935-500">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="d7935-501">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="d7935-501">**EWS Tokens**</span></span>

<span data-ttu-id="d7935-p133">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="d7935-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="d7935-504">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="d7935-504">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="d7935-505">Você pode passar o token e também um identificador de anexo ou um identificador de item a um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="d7935-505">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="d7935-506">O sistema de terceiros usa o token como um token de autorização de portador para chamar a operação [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) dos serviços Web do Exchange (EWS) ou a operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) para recuperar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="d7935-506">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="d7935-507">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d7935-507">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-508">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-508">Parameters</span></span>

|<span data-ttu-id="d7935-509">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-509">Name</span></span>| <span data-ttu-id="d7935-510">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-510">Type</span></span>| <span data-ttu-id="d7935-511">Atributos</span><span class="sxs-lookup"><span data-stu-id="d7935-511">Attributes</span></span>| <span data-ttu-id="d7935-512">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-512">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="d7935-513">Object</span><span class="sxs-lookup"><span data-stu-id="d7935-513">Object</span></span> | <span data-ttu-id="d7935-514">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-514">&lt;optional&gt;</span></span> | <span data-ttu-id="d7935-515">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="d7935-515">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="d7935-516">Booliano</span><span class="sxs-lookup"><span data-stu-id="d7935-516">Boolean</span></span> |  <span data-ttu-id="d7935-517">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-517">&lt;optional&gt;</span></span> | <span data-ttu-id="d7935-p135">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="d7935-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d7935-520">Objeto</span><span class="sxs-lookup"><span data-stu-id="d7935-520">Object</span></span> |  <span data-ttu-id="d7935-521">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-521">&lt;optional&gt;</span></span> | <span data-ttu-id="d7935-522">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="d7935-522">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="d7935-523">function</span><span class="sxs-lookup"><span data-stu-id="d7935-523">function</span></span>||<span data-ttu-id="d7935-524">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d7935-524">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d7935-525">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d7935-525">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="d7935-526">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="d7935-526">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d7935-527">Erros</span><span class="sxs-lookup"><span data-stu-id="d7935-527">Errors</span></span>

|<span data-ttu-id="d7935-528">Código de erro</span><span class="sxs-lookup"><span data-stu-id="d7935-528">Error code</span></span>|<span data-ttu-id="d7935-529">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-529">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="d7935-530">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="d7935-530">The request has failed.</span></span> <span data-ttu-id="d7935-531">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="d7935-531">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="d7935-532">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="d7935-532">The Exchange server returned an error.</span></span> <span data-ttu-id="d7935-533">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="d7935-533">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="d7935-534">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="d7935-534">The user is no longer connected to the network.</span></span> <span data-ttu-id="d7935-535">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="d7935-535">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7935-536">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-536">Requirements</span></span>

|<span data-ttu-id="d7935-537">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-537">Requirement</span></span>| <span data-ttu-id="d7935-538">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-539">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-540">1,5</span><span class="sxs-lookup"><span data-stu-id="d7935-540">1.5</span></span> |
|[<span data-ttu-id="d7935-541">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-542">ReadItem</span></span>|
|[<span data-ttu-id="d7935-543">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d7935-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-544">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="d7935-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7935-545">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-545">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="d7935-546">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d7935-546">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d7935-547">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="d7935-547">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="d7935-p139">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="d7935-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="d7935-550">Você pode passar o token e também um identificador de anexo ou um identificador de item a um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="d7935-550">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="d7935-551">O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="d7935-551">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="d7935-552">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d7935-552">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d7935-553">Chamar o método `getCallbackTokenAsync` no modo de leitura requer um nível de permissão mínimo de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="d7935-553">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="d7935-554">Chamar `getCallbackTokenAsync` no modo redigir exige que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="d7935-554">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="d7935-555">O método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="d7935-555">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-556">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-556">Parameters</span></span>

|<span data-ttu-id="d7935-557">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-557">Name</span></span>| <span data-ttu-id="d7935-558">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-558">Type</span></span>| <span data-ttu-id="d7935-559">Atributos</span><span class="sxs-lookup"><span data-stu-id="d7935-559">Attributes</span></span>| <span data-ttu-id="d7935-560">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-560">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d7935-561">function</span><span class="sxs-lookup"><span data-stu-id="d7935-561">function</span></span>||<span data-ttu-id="d7935-562">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d7935-562">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d7935-563">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d7935-563">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="d7935-564">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="d7935-564">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="d7935-565">Objeto</span><span class="sxs-lookup"><span data-stu-id="d7935-565">Object</span></span>| <span data-ttu-id="d7935-566">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-566">&lt;optional&gt;</span></span>|<span data-ttu-id="d7935-567">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="d7935-567">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d7935-568">Erros</span><span class="sxs-lookup"><span data-stu-id="d7935-568">Errors</span></span>

|<span data-ttu-id="d7935-569">Código de erro</span><span class="sxs-lookup"><span data-stu-id="d7935-569">Error code</span></span>|<span data-ttu-id="d7935-570">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-570">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="d7935-571">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="d7935-571">The request has failed.</span></span> <span data-ttu-id="d7935-572">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="d7935-572">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="d7935-573">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="d7935-573">The Exchange server returned an error.</span></span> <span data-ttu-id="d7935-574">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="d7935-574">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="d7935-575">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="d7935-575">The user is no longer connected to the network.</span></span> <span data-ttu-id="d7935-576">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="d7935-576">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7935-577">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-577">Requirements</span></span>

|<span data-ttu-id="d7935-578">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-578">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d7935-579">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-579">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-580">1.0</span><span class="sxs-lookup"><span data-stu-id="d7935-580">1.0</span></span> | <span data-ttu-id="d7935-581">1.3</span><span class="sxs-lookup"><span data-stu-id="d7935-581">1.3</span></span> |
|[<span data-ttu-id="d7935-582">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-583">ReadItem</span></span> | <span data-ttu-id="d7935-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-584">ReadItem</span></span> |
|[<span data-ttu-id="d7935-585">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d7935-585">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-586">Read</span><span class="sxs-lookup"><span data-stu-id="d7935-586">Read</span></span> | <span data-ttu-id="d7935-587">Escrever</span><span class="sxs-lookup"><span data-stu-id="d7935-587">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="d7935-588">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-588">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="d7935-589">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d7935-589">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d7935-590">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="d7935-590">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="d7935-591">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="d7935-591">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-592">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-592">Parameters</span></span>

|<span data-ttu-id="d7935-593">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-593">Name</span></span>| <span data-ttu-id="d7935-594">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-594">Type</span></span>| <span data-ttu-id="d7935-595">Atributos</span><span class="sxs-lookup"><span data-stu-id="d7935-595">Attributes</span></span>| <span data-ttu-id="d7935-596">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-596">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d7935-597">function</span><span class="sxs-lookup"><span data-stu-id="d7935-597">function</span></span>||<span data-ttu-id="d7935-598">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d7935-598">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d7935-599">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d7935-599">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="d7935-600">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="d7935-600">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="d7935-601">Objeto</span><span class="sxs-lookup"><span data-stu-id="d7935-601">Object</span></span>| <span data-ttu-id="d7935-602">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-602">&lt;optional&gt;</span></span>|<span data-ttu-id="d7935-603">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="d7935-603">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d7935-604">Erros</span><span class="sxs-lookup"><span data-stu-id="d7935-604">Errors</span></span>

|<span data-ttu-id="d7935-605">Código de erro</span><span class="sxs-lookup"><span data-stu-id="d7935-605">Error code</span></span>|<span data-ttu-id="d7935-606">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-606">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="d7935-607">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="d7935-607">The request has failed.</span></span> <span data-ttu-id="d7935-608">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="d7935-608">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="d7935-609">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="d7935-609">The Exchange server returned an error.</span></span> <span data-ttu-id="d7935-610">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="d7935-610">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="d7935-611">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="d7935-611">The user is no longer connected to the network.</span></span> <span data-ttu-id="d7935-612">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="d7935-612">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7935-613">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-613">Requirements</span></span>

|<span data-ttu-id="d7935-614">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-614">Requirement</span></span>| <span data-ttu-id="d7935-615">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-615">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-616">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-616">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-617">1.0</span><span class="sxs-lookup"><span data-stu-id="d7935-617">1.0</span></span>|
|[<span data-ttu-id="d7935-618">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-618">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-619">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-619">ReadItem</span></span>|
|[<span data-ttu-id="d7935-620">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="d7935-620">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-621">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-621">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7935-622">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-622">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="d7935-623">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d7935-623">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="d7935-624">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="d7935-624">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="d7935-625">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="d7935-625">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="d7935-626">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="d7935-626">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="d7935-627">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="d7935-627">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="d7935-628">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="d7935-628">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="d7935-629">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="d7935-629">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="d7935-630">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="d7935-630">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="d7935-631">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="d7935-631">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="d7935-632">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="d7935-632">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="d7935-p149">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="d7935-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="d7935-635">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="d7935-635">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="d7935-636">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="d7935-636">Version differences</span></span>

<span data-ttu-id="d7935-637">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="d7935-637">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="d7935-638">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="d7935-638">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="d7935-639">Você pode determinar se o seu aplicativo de email está em execução no Outlook na Web ou em um cliente de desktop usando a propriedade Mailbox. Diagnostics. hostName.</span><span class="sxs-lookup"><span data-stu-id="d7935-639">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="d7935-640">Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="d7935-640">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-641">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-641">Parameters</span></span>

|<span data-ttu-id="d7935-642">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-642">Name</span></span>| <span data-ttu-id="d7935-643">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-643">Type</span></span>| <span data-ttu-id="d7935-644">Atributos</span><span class="sxs-lookup"><span data-stu-id="d7935-644">Attributes</span></span>| <span data-ttu-id="d7935-645">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-645">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d7935-646">String</span><span class="sxs-lookup"><span data-stu-id="d7935-646">String</span></span>||<span data-ttu-id="d7935-647">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="d7935-647">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="d7935-648">function</span><span class="sxs-lookup"><span data-stu-id="d7935-648">function</span></span>||<span data-ttu-id="d7935-649">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d7935-649">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d7935-650">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d7935-650">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="d7935-651">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="d7935-651">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="d7935-652">Objeto</span><span class="sxs-lookup"><span data-stu-id="d7935-652">Object</span></span>| <span data-ttu-id="d7935-653">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-653">&lt;optional&gt;</span></span>|<span data-ttu-id="d7935-654">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="d7935-654">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7935-655">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-655">Requirements</span></span>

|<span data-ttu-id="d7935-656">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-656">Requirement</span></span>| <span data-ttu-id="d7935-657">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-657">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-658">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-658">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-659">1.0</span><span class="sxs-lookup"><span data-stu-id="d7935-659">1.0</span></span>|
|[<span data-ttu-id="d7935-660">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-660">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-661">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="d7935-661">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="d7935-662">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d7935-662">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-663">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-663">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7935-664">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d7935-664">Example</span></span>

<span data-ttu-id="d7935-665">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="d7935-665">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="d7935-666">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d7935-666">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="d7935-667">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="d7935-667">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="d7935-668">Atualmente, os tipos de eventos com `Office.EventType.ItemChanged` suporte `Office.EventType.OfficeThemeChanged`são e.</span><span class="sxs-lookup"><span data-stu-id="d7935-668">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7935-669">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="d7935-669">Parameters</span></span>

| <span data-ttu-id="d7935-670">Nome</span><span class="sxs-lookup"><span data-stu-id="d7935-670">Name</span></span> | <span data-ttu-id="d7935-671">Tipo</span><span class="sxs-lookup"><span data-stu-id="d7935-671">Type</span></span> | <span data-ttu-id="d7935-672">Atributos</span><span class="sxs-lookup"><span data-stu-id="d7935-672">Attributes</span></span> | <span data-ttu-id="d7935-673">Descrição</span><span class="sxs-lookup"><span data-stu-id="d7935-673">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d7935-674">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d7935-674">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d7935-675">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="d7935-675">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="d7935-676">Objeto</span><span class="sxs-lookup"><span data-stu-id="d7935-676">Object</span></span> | <span data-ttu-id="d7935-677">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-677">&lt;optional&gt;</span></span> | <span data-ttu-id="d7935-678">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="d7935-678">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d7935-679">Objeto</span><span class="sxs-lookup"><span data-stu-id="d7935-679">Object</span></span> | <span data-ttu-id="d7935-680">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-680">&lt;optional&gt;</span></span> | <span data-ttu-id="d7935-681">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="d7935-681">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d7935-682">function</span><span class="sxs-lookup"><span data-stu-id="d7935-682">function</span></span>| <span data-ttu-id="d7935-683">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7935-683">&lt;optional&gt;</span></span>|<span data-ttu-id="d7935-684">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d7935-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7935-685">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d7935-685">Requirements</span></span>

|<span data-ttu-id="d7935-686">Requisito</span><span class="sxs-lookup"><span data-stu-id="d7935-686">Requirement</span></span>| <span data-ttu-id="d7935-687">Valor</span><span class="sxs-lookup"><span data-stu-id="d7935-687">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7935-688">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d7935-688">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7935-689">1,5</span><span class="sxs-lookup"><span data-stu-id="d7935-689">1.5</span></span> |
|[<span data-ttu-id="d7935-690">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d7935-690">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7935-691">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7935-691">ReadItem</span></span> |
|[<span data-ttu-id="d7935-692">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d7935-692">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7935-693">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d7935-693">Compose or Read</span></span>|
