---
title: Office. Context. Mailbox-visualização do conjunto de requisitos
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 951bb4ff338507f369f23e7c095debdb6e7a945e
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696474"
---
# <a name="mailbox"></a><span data-ttu-id="c7a4e-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="c7a4e-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="c7a4e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="c7a4e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="c7a4e-104">Fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a4e-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-105">Requirements</span></span>

|<span data-ttu-id="c7a4e-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-106">Requirement</span></span>| <span data-ttu-id="c7a4e-107">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a4e-109">1.0</span></span>|
|[<span data-ttu-id="c7a4e-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-111">Restricted</span></span>|
|[<span data-ttu-id="c7a4e-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c7a4e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c7a4e-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-114">Members and methods</span></span>

| <span data-ttu-id="c7a4e-115">Membro</span><span class="sxs-lookup"><span data-stu-id="c7a4e-115">Member</span></span> | <span data-ttu-id="c7a4e-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c7a4e-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="c7a4e-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="c7a4e-118">Membro</span><span class="sxs-lookup"><span data-stu-id="c7a4e-118">Member</span></span> |
| [<span data-ttu-id="c7a4e-119">Nova mastercategories</span><span class="sxs-lookup"><span data-stu-id="c7a4e-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="c7a4e-120">Membro</span><span class="sxs-lookup"><span data-stu-id="c7a4e-120">Member</span></span> |
| [<span data-ttu-id="c7a4e-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="c7a4e-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="c7a4e-122">Membro</span><span class="sxs-lookup"><span data-stu-id="c7a4e-122">Member</span></span> |
| [<span data-ttu-id="c7a4e-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c7a4e-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c7a4e-124">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-124">Method</span></span> |
| [<span data-ttu-id="c7a4e-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="c7a4e-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="c7a4e-126">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-126">Method</span></span> |
| [<span data-ttu-id="c7a4e-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="c7a4e-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="c7a4e-128">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-128">Method</span></span> |
| [<span data-ttu-id="c7a4e-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="c7a4e-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="c7a4e-130">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-130">Method</span></span> |
| [<span data-ttu-id="c7a4e-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="c7a4e-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="c7a4e-132">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-132">Method</span></span> |
| [<span data-ttu-id="c7a4e-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="c7a4e-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="c7a4e-134">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-134">Method</span></span> |
| [<span data-ttu-id="c7a4e-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="c7a4e-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="c7a4e-136">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-136">Method</span></span> |
| [<span data-ttu-id="c7a4e-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="c7a4e-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="c7a4e-138">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-138">Method</span></span> |
| [<span data-ttu-id="c7a4e-139">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="c7a4e-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="c7a4e-140">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-140">Method</span></span> |
| [<span data-ttu-id="c7a4e-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c7a4e-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="c7a4e-142">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-142">Method</span></span> |
| [<span data-ttu-id="c7a4e-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c7a4e-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="c7a4e-144">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-144">Method</span></span> |
| [<span data-ttu-id="c7a4e-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="c7a4e-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="c7a4e-146">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-146">Method</span></span> |
| [<span data-ttu-id="c7a4e-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="c7a4e-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="c7a4e-148">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-148">Method</span></span> |
| [<span data-ttu-id="c7a4e-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c7a4e-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="c7a4e-150">Método</span><span class="sxs-lookup"><span data-stu-id="c7a4e-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c7a4e-151">Namespaces</span><span class="sxs-lookup"><span data-stu-id="c7a4e-151">Namespaces</span></span>

<span data-ttu-id="c7a4e-152">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="c7a4e-153">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="c7a4e-154">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="c7a4e-155">Membros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="c7a4e-156">ewsUrl: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c7a4e-156">ewsUrl: String</span></span>

<span data-ttu-id="c7a4e-157">Obtém a URL do ponto de extremidade dos Serviços Web do Exchange (EWS) para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-157">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="c7a4e-158">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-158">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a4e-159">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c7a4e-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="c7a4e-162">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="c7a4e-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a4e-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-165">Type</span></span>

*   <span data-ttu-id="c7a4e-166">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a4e-167">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-167">Requirements</span></span>

|<span data-ttu-id="c7a4e-168">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-168">Requirement</span></span>| <span data-ttu-id="c7a4e-169">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-170">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-171">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a4e-171">1.0</span></span>|
|[<span data-ttu-id="c7a4e-172">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-173">ReadItem</span></span>|
|[<span data-ttu-id="c7a4e-174">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c7a4e-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-175">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-175">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="c7a4e-176">Nova mastercategories: [nova mastercategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="c7a4e-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="c7a4e-177">Obtém um objeto que fornece métodos para gerenciar a lista mestra de categorias nesta caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a4e-178">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a4e-179">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-179">Type</span></span>

*   [<span data-ttu-id="c7a4e-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="c7a4e-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="c7a4e-181">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-181">Requirements</span></span>

|<span data-ttu-id="c7a4e-182">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-182">Requirement</span></span>| <span data-ttu-id="c7a4e-183">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-184">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-185">Visualização</span><span class="sxs-lookup"><span data-stu-id="c7a4e-185">Preview</span></span> |
|[<span data-ttu-id="c7a4e-186">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="c7a4e-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="c7a4e-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c7a4e-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="c7a4e-190">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-190">Example</span></span>

<span data-ttu-id="c7a4e-191">Este exemplo obtém a lista mestra de categorias para esta caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-191">This example gets the categories master list for this mailbox.</span></span>

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

#### <a name="resturl-string"></a><span data-ttu-id="c7a4e-192">restUrl: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c7a4e-192">restUrl: String</span></span>

<span data-ttu-id="c7a4e-193">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="c7a4e-194">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="c7a4e-195">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-195">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="c7a4e-p104">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a4e-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-198">Type</span></span>

*   <span data-ttu-id="c7a4e-199">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a4e-200">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-200">Requirements</span></span>

|<span data-ttu-id="c7a4e-201">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-201">Requirement</span></span>| <span data-ttu-id="c7a4e-202">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-203">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-204">1,5</span><span class="sxs-lookup"><span data-stu-id="c7a4e-204">1.5</span></span> |
|[<span data-ttu-id="c7a4e-205">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-206">ReadItem</span></span>|
|[<span data-ttu-id="c7a4e-207">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c7a4e-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-208">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-208">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="c7a4e-209">Métodos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-209">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c7a4e-210">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c7a4e-210">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c7a4e-211">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-211">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c7a4e-212">Atualmente, os tipos de eventos com `Office.EventType.ItemChanged` suporte `Office.EventType.OfficeThemeChanged`são e.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-212">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-213">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-213">Parameters</span></span>

| <span data-ttu-id="c7a4e-214">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-214">Name</span></span> | <span data-ttu-id="c7a4e-215">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-215">Type</span></span> | <span data-ttu-id="c7a4e-216">Atributos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-216">Attributes</span></span> | <span data-ttu-id="c7a4e-217">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-217">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c7a4e-218">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c7a4e-218">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c7a4e-219">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-219">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c7a4e-220">Função</span><span class="sxs-lookup"><span data-stu-id="c7a4e-220">Function</span></span> || <span data-ttu-id="c7a4e-p105">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c7a4e-224">Objeto</span><span class="sxs-lookup"><span data-stu-id="c7a4e-224">Object</span></span> | <span data-ttu-id="c7a4e-225">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-225">&lt;optional&gt;</span></span> | <span data-ttu-id="c7a4e-226">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c7a4e-227">Objeto</span><span class="sxs-lookup"><span data-stu-id="c7a4e-227">Object</span></span> | <span data-ttu-id="c7a4e-228">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-228">&lt;optional&gt;</span></span> | <span data-ttu-id="c7a4e-229">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c7a4e-230">function</span><span class="sxs-lookup"><span data-stu-id="c7a4e-230">function</span></span>| <span data-ttu-id="c7a4e-231">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-231">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a4e-232">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c7a4e-232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a4e-233">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-233">Requirements</span></span>

|<span data-ttu-id="c7a4e-234">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-234">Requirement</span></span>| <span data-ttu-id="c7a4e-235">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-236">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-237">1,5</span><span class="sxs-lookup"><span data-stu-id="c7a4e-237">1.5</span></span> |
|[<span data-ttu-id="c7a4e-238">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-239">ReadItem</span></span> |
|[<span data-ttu-id="c7a4e-240">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c7a4e-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-241">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a4e-242">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-242">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="c7a4e-243">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="c7a4e-243">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="c7a4e-244">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-244">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a4e-245">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-245">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c7a4e-p106">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-248">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-248">Parameters</span></span>

|<span data-ttu-id="c7a4e-249">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-249">Name</span></span>| <span data-ttu-id="c7a4e-250">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-250">Type</span></span>| <span data-ttu-id="c7a4e-251">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-251">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c7a4e-252">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-252">String</span></span>|<span data-ttu-id="c7a4e-253">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="c7a4e-253">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="c7a4e-254">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="c7a4e-254">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="c7a4e-255">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-255">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a4e-256">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-256">Requirements</span></span>

|<span data-ttu-id="c7a4e-257">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-257">Requirement</span></span>| <span data-ttu-id="c7a4e-258">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-259">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-260">1.3</span><span class="sxs-lookup"><span data-stu-id="c7a4e-260">1.3</span></span>|
|[<span data-ttu-id="c7a4e-261">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-262">Restrito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-262">Restricted</span></span>|
|[<span data-ttu-id="c7a4e-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c7a4e-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-264">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c7a4e-265">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c7a4e-265">Returns:</span></span>

<span data-ttu-id="c7a4e-266">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-266">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c7a4e-267">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-267">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="c7a4e-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="c7a4e-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="c7a4e-269">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-269">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="c7a4e-270">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para datas e horas.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-270">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="c7a4e-271">O Outlook em uma área de trabalho usa o fuso horário do computador cliente; O Outlook na Web usa o fuso horário definido no centro de administração do Exchange (Eat).</span><span class="sxs-lookup"><span data-stu-id="c7a4e-271">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="c7a4e-272">Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-272">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="c7a4e-273">Se o aplicativo de email estiver em execução no Outlook em um cliente desktop `convertToLocalClientTime` , o método retornará um objeto Dictionary com os valores definidos para o fuso horário do computador cliente.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-273">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="c7a4e-274">Se o aplicativo de email estiver em execução no Outlook na Web, `convertToLocalClientTime` o método retornará um objeto Dictionary com os valores definidos para o fuso horário especificado no Eat.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-274">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-275">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-275">Parameters</span></span>

|<span data-ttu-id="c7a4e-276">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-276">Name</span></span>| <span data-ttu-id="c7a4e-277">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-277">Type</span></span>| <span data-ttu-id="c7a4e-278">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-278">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="c7a4e-279">Date</span><span class="sxs-lookup"><span data-stu-id="c7a4e-279">Date</span></span>|<span data-ttu-id="c7a4e-280">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="c7a4e-280">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a4e-281">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-281">Requirements</span></span>

|<span data-ttu-id="c7a4e-282">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-282">Requirement</span></span>| <span data-ttu-id="c7a4e-283">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-284">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-285">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a4e-285">1.0</span></span>|
|[<span data-ttu-id="c7a4e-286">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-287">ReadItem</span></span>|
|[<span data-ttu-id="c7a4e-288">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c7a4e-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-289">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-289">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c7a4e-290">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c7a4e-290">Returns:</span></span>

<span data-ttu-id="c7a4e-291">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="c7a4e-291">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="c7a4e-292">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="c7a4e-292">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="c7a4e-293">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-293">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a4e-294">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-294">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c7a4e-p109">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-297">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-297">Parameters</span></span>

|<span data-ttu-id="c7a4e-298">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-298">Name</span></span>| <span data-ttu-id="c7a4e-299">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-299">Type</span></span>| <span data-ttu-id="c7a4e-300">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-300">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c7a4e-301">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-301">String</span></span>|<span data-ttu-id="c7a4e-302">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="c7a4e-302">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="c7a4e-303">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="c7a4e-303">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="c7a4e-304">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-304">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a4e-305">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-305">Requirements</span></span>

|<span data-ttu-id="c7a4e-306">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-306">Requirement</span></span>| <span data-ttu-id="c7a4e-307">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-308">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-309">1.3</span><span class="sxs-lookup"><span data-stu-id="c7a4e-309">1.3</span></span>|
|[<span data-ttu-id="c7a4e-310">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-311">Restrito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-311">Restricted</span></span>|
|[<span data-ttu-id="c7a4e-312">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c7a4e-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-313">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c7a4e-314">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c7a4e-314">Returns:</span></span>

<span data-ttu-id="c7a4e-315">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-315">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="c7a4e-316">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-316">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="c7a4e-317">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="c7a4e-317">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="c7a4e-318">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-318">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="c7a4e-319">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-319">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-320">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-320">Parameters</span></span>

|<span data-ttu-id="c7a4e-321">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-321">Name</span></span>| <span data-ttu-id="c7a4e-322">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-322">Type</span></span>| <span data-ttu-id="c7a4e-323">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-323">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="c7a4e-324">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="c7a4e-324">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="c7a4e-325">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-325">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a4e-326">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-326">Requirements</span></span>

|<span data-ttu-id="c7a4e-327">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-327">Requirement</span></span>| <span data-ttu-id="c7a4e-328">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-329">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-330">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a4e-330">1.0</span></span>|
|[<span data-ttu-id="c7a4e-331">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-332">ReadItem</span></span>|
|[<span data-ttu-id="c7a4e-333">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c7a4e-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-334">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-334">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c7a4e-335">Retorna:</span><span class="sxs-lookup"><span data-stu-id="c7a4e-335">Returns:</span></span>

<span data-ttu-id="c7a4e-336">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-336">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="c7a4e-337">Tipo: data</span><span class="sxs-lookup"><span data-stu-id="c7a4e-337">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="c7a4e-338">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-338">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="c7a4e-339">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="c7a4e-339">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="c7a4e-340">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-340">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a4e-341">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-341">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c7a4e-342">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-342">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="c7a4e-343">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente ou o compromisso mestre de uma série recorrente, mas não é possível exibir uma instância da série.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-343">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="c7a4e-344">Isso ocorre porque, no Outlook no Mac, você não pode acessar as propriedades (incluindo a ID do item) de instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-344">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="c7a4e-345">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres 32 KB.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-345">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="c7a4e-346">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-346">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-347">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-347">Parameters</span></span>

|<span data-ttu-id="c7a4e-348">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-348">Name</span></span>| <span data-ttu-id="c7a4e-349">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-349">Type</span></span>| <span data-ttu-id="c7a4e-350">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-350">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c7a4e-351">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-351">String</span></span>|<span data-ttu-id="c7a4e-352">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-352">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a4e-353">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-353">Requirements</span></span>

|<span data-ttu-id="c7a4e-354">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-354">Requirement</span></span>| <span data-ttu-id="c7a4e-355">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-356">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-357">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a4e-357">1.0</span></span>|
|[<span data-ttu-id="c7a4e-358">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-359">ReadItem</span></span>|
|[<span data-ttu-id="c7a4e-360">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c7a4e-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-361">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-361">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a4e-362">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-362">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="c7a4e-363">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="c7a4e-363">displayMessageForm(itemId)</span></span>

<span data-ttu-id="c7a4e-364">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-364">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a4e-365">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-365">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c7a4e-366">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-366">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="c7a4e-367">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-367">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="c7a4e-368">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-368">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="c7a4e-p111">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-371">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-371">Parameters</span></span>

|<span data-ttu-id="c7a4e-372">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-372">Name</span></span>| <span data-ttu-id="c7a4e-373">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-373">Type</span></span>| <span data-ttu-id="c7a4e-374">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-374">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="c7a4e-375">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-375">String</span></span>|<span data-ttu-id="c7a4e-376">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-376">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a4e-377">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-377">Requirements</span></span>

|<span data-ttu-id="c7a4e-378">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-378">Requirement</span></span>| <span data-ttu-id="c7a4e-379">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-380">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-381">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a4e-381">1.0</span></span>|
|[<span data-ttu-id="c7a4e-382">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-382">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-383">ReadItem</span></span>|
|[<span data-ttu-id="c7a4e-384">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c7a4e-384">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-385">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-385">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a4e-386">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-386">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="c7a4e-387">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="c7a4e-387">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="c7a4e-388">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-388">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a4e-389">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-389">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c7a4e-p112">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="c7a4e-392">No Outlook na Web e dispositivos móveis, este método sempre exibe um formulário com um campo participantes.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-392">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="c7a4e-393">Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-393">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="c7a4e-394">Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-394">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="c7a4e-p114">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="c7a4e-397">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-397">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-398">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-398">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="c7a4e-399">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-399">All parameters are optional.</span></span>

|<span data-ttu-id="c7a4e-400">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-400">Name</span></span>| <span data-ttu-id="c7a4e-401">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-401">Type</span></span>| <span data-ttu-id="c7a4e-402">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-402">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="c7a4e-403">Object</span><span class="sxs-lookup"><span data-stu-id="c7a4e-403">Object</span></span> | <span data-ttu-id="c7a4e-404">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-404">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="c7a4e-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="c7a4e-p115">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="c7a4e-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="c7a4e-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="c7a4e-411">Data</span><span class="sxs-lookup"><span data-stu-id="c7a4e-411">Date</span></span> | <span data-ttu-id="c7a4e-412">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-412">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="c7a4e-413">Data</span><span class="sxs-lookup"><span data-stu-id="c7a4e-413">Date</span></span> | <span data-ttu-id="c7a4e-414">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-414">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="c7a4e-415">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-415">String</span></span> | <span data-ttu-id="c7a4e-p117">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="c7a4e-418">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-418">Array.&lt;String&gt;</span></span> | <span data-ttu-id="c7a4e-p118">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="c7a4e-421">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-421">String</span></span> | <span data-ttu-id="c7a4e-p119">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="c7a4e-424">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-424">String</span></span> | <span data-ttu-id="c7a4e-p120">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c7a4e-427">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-427">Requirements</span></span>

|<span data-ttu-id="c7a4e-428">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-428">Requirement</span></span>| <span data-ttu-id="c7a4e-429">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-430">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-431">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a4e-431">1.0</span></span>|
|[<span data-ttu-id="c7a4e-432">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-433">ReadItem</span></span>|
|[<span data-ttu-id="c7a4e-434">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c7a4e-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-435">Read</span><span class="sxs-lookup"><span data-stu-id="c7a4e-435">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a4e-436">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-436">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="c7a4e-437">displayNewMessageForm (parâmetros)</span><span class="sxs-lookup"><span data-stu-id="c7a4e-437">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="c7a4e-438">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-438">Displays a form for creating a new message.</span></span>

<span data-ttu-id="c7a4e-439">O `displayNewMessageForm` método abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-439">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="c7a4e-440">Se os parâmetros forem especificados, os campos de formulário da mensagem serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-440">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="c7a4e-441">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-441">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-442">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-442">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="c7a4e-443">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-443">All parameters are optional.</span></span>

|<span data-ttu-id="c7a4e-444">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-444">Name</span></span>| <span data-ttu-id="c7a4e-445">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-445">Type</span></span>| <span data-ttu-id="c7a4e-446">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-446">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="c7a4e-447">Objeto</span><span class="sxs-lookup"><span data-stu-id="c7a4e-447">Object</span></span> | <span data-ttu-id="c7a4e-448">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-448">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="c7a4e-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="c7a4e-450">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha para.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-450">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="c7a4e-451">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-451">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="c7a4e-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="c7a4e-453">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha CC.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-453">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="c7a4e-454">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-454">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="c7a4e-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="c7a4e-456">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-456">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="c7a4e-457">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-457">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="c7a4e-458">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-458">String</span></span> | <span data-ttu-id="c7a4e-459">Uma cadeia de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-459">A string containing the subject of the message.</span></span> <span data-ttu-id="c7a4e-460">A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-460">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="c7a4e-461">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-461">String</span></span> | <span data-ttu-id="c7a4e-462">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-462">The HTML body of the message.</span></span> <span data-ttu-id="c7a4e-463">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-463">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="c7a4e-464">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-464">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c7a4e-465">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-465">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="c7a4e-466">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-466">String</span></span> | <span data-ttu-id="c7a4e-p127">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="c7a4e-469">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c7a4e-469">String</span></span> | <span data-ttu-id="c7a4e-470">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-470">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="c7a4e-471">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-471">String</span></span> | <span data-ttu-id="c7a4e-p128">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="c7a4e-474">Booliano</span><span class="sxs-lookup"><span data-stu-id="c7a4e-474">Boolean</span></span> | <span data-ttu-id="c7a4e-p129">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="c7a4e-477">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-477">String</span></span> | <span data-ttu-id="c7a4e-478">Usado somente se `type` estiver definido como `item`.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-478">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="c7a4e-479">A ID do item do EWS do email existente que você deseja anexar à nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-479">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="c7a4e-480">Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-480">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="c7a4e-481">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-481">Requirements</span></span>

|<span data-ttu-id="c7a4e-482">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-482">Requirement</span></span>| <span data-ttu-id="c7a4e-483">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-484">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-485">1.6</span><span class="sxs-lookup"><span data-stu-id="c7a4e-485">1.6</span></span> |
|[<span data-ttu-id="c7a4e-486">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-487">ReadItem</span></span>|
|[<span data-ttu-id="c7a4e-488">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c7a4e-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-489">Read</span><span class="sxs-lookup"><span data-stu-id="c7a4e-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a4e-490">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-490">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="c7a4e-491">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c7a4e-491">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="c7a4e-492">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-492">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="c7a4e-p131">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a4e-495">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-495">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="c7a4e-496">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="c7a4e-496">**REST Tokens**</span></span>

<span data-ttu-id="c7a4e-p132">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="c7a4e-500">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-500">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="c7a4e-501">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="c7a4e-501">**EWS Tokens**</span></span>

<span data-ttu-id="c7a4e-p133">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="c7a4e-504">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-504">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-505">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-505">Parameters</span></span>

|<span data-ttu-id="c7a4e-506">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-506">Name</span></span>| <span data-ttu-id="c7a4e-507">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-507">Type</span></span>| <span data-ttu-id="c7a4e-508">Atributos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-508">Attributes</span></span>| <span data-ttu-id="c7a4e-509">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-509">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="c7a4e-510">Object</span><span class="sxs-lookup"><span data-stu-id="c7a4e-510">Object</span></span> | <span data-ttu-id="c7a4e-511">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-511">&lt;optional&gt;</span></span> | <span data-ttu-id="c7a4e-512">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-512">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="c7a4e-513">Booliano</span><span class="sxs-lookup"><span data-stu-id="c7a4e-513">Boolean</span></span> |  <span data-ttu-id="c7a4e-514">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-514">&lt;optional&gt;</span></span> | <span data-ttu-id="c7a4e-p134">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c7a4e-517">Objeto</span><span class="sxs-lookup"><span data-stu-id="c7a4e-517">Object</span></span> |  <span data-ttu-id="c7a4e-518">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-518">&lt;optional&gt;</span></span> | <span data-ttu-id="c7a4e-519">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-519">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="c7a4e-520">function</span><span class="sxs-lookup"><span data-stu-id="c7a4e-520">function</span></span>||<span data-ttu-id="c7a4e-521">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c7a4e-521">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c7a4e-522">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-522">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="c7a4e-523">Se houvesse um erro, as `asyncResult.error` propriedades `asyncResult.diagnostics` e podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-523">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c7a4e-524">Erros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-524">Errors</span></span>

|<span data-ttu-id="c7a4e-525">Código de erro</span><span class="sxs-lookup"><span data-stu-id="c7a4e-525">Error code</span></span>|<span data-ttu-id="c7a4e-526">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-526">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="c7a4e-527">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-527">The request has failed.</span></span> <span data-ttu-id="c7a4e-528">Confira o objeto Diagnostics do código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-528">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="c7a4e-529">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-529">The Exchange server returned an error.</span></span> <span data-ttu-id="c7a4e-530">Confira o objeto Diagnostics para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-530">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="c7a4e-531">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-531">The user is no longer connected to the network.</span></span> <span data-ttu-id="c7a4e-532">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-532">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a4e-533">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-533">Requirements</span></span>

|<span data-ttu-id="c7a4e-534">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-534">Requirement</span></span>| <span data-ttu-id="c7a4e-535">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-535">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-536">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-536">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-537">1,5</span><span class="sxs-lookup"><span data-stu-id="c7a4e-537">1.5</span></span> |
|[<span data-ttu-id="c7a4e-538">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-538">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-539">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-539">ReadItem</span></span>|
|[<span data-ttu-id="c7a4e-540">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c7a4e-540">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-541">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="c7a4e-541">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a4e-542">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-542">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="c7a4e-543">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c7a4e-543">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="c7a4e-544">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-544">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="c7a4e-p138">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="c7a4e-p139">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="c7a4e-550">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-550">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="c7a4e-p140">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-553">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-553">Parameters</span></span>

|<span data-ttu-id="c7a4e-554">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-554">Name</span></span>| <span data-ttu-id="c7a4e-555">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-555">Type</span></span>| <span data-ttu-id="c7a4e-556">Atributos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-556">Attributes</span></span>| <span data-ttu-id="c7a4e-557">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-557">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c7a4e-558">function</span><span class="sxs-lookup"><span data-stu-id="c7a4e-558">function</span></span>||<span data-ttu-id="c7a4e-559">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c7a4e-559">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c7a4e-560">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-560">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="c7a4e-561">Se houvesse um erro, as `asyncResult.error` propriedades `asyncResult.diagnostics` e podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-561">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="c7a4e-562">Objeto</span><span class="sxs-lookup"><span data-stu-id="c7a4e-562">Object</span></span>| <span data-ttu-id="c7a4e-563">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-563">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a4e-564">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-564">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c7a4e-565">Erros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-565">Errors</span></span>

|<span data-ttu-id="c7a4e-566">Código de erro</span><span class="sxs-lookup"><span data-stu-id="c7a4e-566">Error code</span></span>|<span data-ttu-id="c7a4e-567">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-567">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="c7a4e-568">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-568">The request has failed.</span></span> <span data-ttu-id="c7a4e-569">Confira o objeto Diagnostics do código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-569">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="c7a4e-570">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-570">The Exchange server returned an error.</span></span> <span data-ttu-id="c7a4e-571">Confira o objeto Diagnostics para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-571">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="c7a4e-572">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-572">The user is no longer connected to the network.</span></span> <span data-ttu-id="c7a4e-573">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-573">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a4e-574">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-574">Requirements</span></span>

|<span data-ttu-id="c7a4e-575">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-575">Requirement</span></span>| <span data-ttu-id="c7a4e-576">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-576">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-577">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-577">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-578">1.3</span><span class="sxs-lookup"><span data-stu-id="c7a4e-578">1.3</span></span>|
|[<span data-ttu-id="c7a4e-579">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-579">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-580">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-580">ReadItem</span></span>|
|[<span data-ttu-id="c7a4e-581">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c7a4e-581">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-582">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="c7a4e-582">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a4e-583">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-583">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="c7a4e-584">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c7a4e-584">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="c7a4e-585">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-585">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="c7a4e-586">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="c7a4e-586">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-587">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-587">Parameters</span></span>

|<span data-ttu-id="c7a4e-588">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-588">Name</span></span>| <span data-ttu-id="c7a4e-589">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-589">Type</span></span>| <span data-ttu-id="c7a4e-590">Atributos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-590">Attributes</span></span>| <span data-ttu-id="c7a4e-591">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-591">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c7a4e-592">function</span><span class="sxs-lookup"><span data-stu-id="c7a4e-592">function</span></span>||<span data-ttu-id="c7a4e-593">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c7a4e-593">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c7a4e-594">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-594">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="c7a4e-595">Se houvesse um erro, as `asyncResult.error` propriedades `asyncResult.diagnostics` e podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-595">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="c7a4e-596">Objeto</span><span class="sxs-lookup"><span data-stu-id="c7a4e-596">Object</span></span>| <span data-ttu-id="c7a4e-597">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-597">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a4e-598">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-598">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c7a4e-599">Erros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-599">Errors</span></span>

|<span data-ttu-id="c7a4e-600">Código de erro</span><span class="sxs-lookup"><span data-stu-id="c7a4e-600">Error code</span></span>|<span data-ttu-id="c7a4e-601">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-601">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="c7a4e-602">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-602">The request has failed.</span></span> <span data-ttu-id="c7a4e-603">Confira o objeto Diagnostics do código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-603">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="c7a4e-604">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-604">The Exchange server returned an error.</span></span> <span data-ttu-id="c7a4e-605">Confira o objeto Diagnostics para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-605">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="c7a4e-606">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-606">The user is no longer connected to the network.</span></span> <span data-ttu-id="c7a4e-607">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-607">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a4e-608">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-608">Requirements</span></span>

|<span data-ttu-id="c7a4e-609">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-609">Requirement</span></span>| <span data-ttu-id="c7a4e-610">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-610">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-611">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-611">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-612">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a4e-612">1.0</span></span>|
|[<span data-ttu-id="c7a4e-613">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-613">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-614">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-614">ReadItem</span></span>|
|[<span data-ttu-id="c7a4e-615">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c7a4e-615">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-616">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-616">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a4e-617">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-617">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="c7a4e-618">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c7a4e-618">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="c7a4e-619">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-619">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a4e-620">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-620">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="c7a4e-621">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="c7a4e-621">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="c7a4e-622">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="c7a4e-622">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="c7a4e-623">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-623">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="c7a4e-624">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-624">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="c7a4e-625">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-625">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="c7a4e-626">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-626">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="c7a4e-627">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-627">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="c7a4e-p148">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="c7a4e-p148">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="c7a4e-630">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-630">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="c7a4e-631">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="c7a4e-631">Version differences</span></span>

<span data-ttu-id="c7a4e-632">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-632">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="c7a4e-633">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-633">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="c7a4e-634">Você pode determinar se o seu aplicativo de email está em execução no Outlook na Web ou em um cliente de desktop usando a propriedade Mailbox. Diagnostics. hostName.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-634">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="c7a4e-635">Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-635">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-636">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-636">Parameters</span></span>

|<span data-ttu-id="c7a4e-637">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-637">Name</span></span>| <span data-ttu-id="c7a4e-638">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-638">Type</span></span>| <span data-ttu-id="c7a4e-639">Atributos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-639">Attributes</span></span>| <span data-ttu-id="c7a4e-640">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-640">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="c7a4e-641">String</span><span class="sxs-lookup"><span data-stu-id="c7a4e-641">String</span></span>||<span data-ttu-id="c7a4e-642">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-642">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="c7a4e-643">function</span><span class="sxs-lookup"><span data-stu-id="c7a4e-643">function</span></span>||<span data-ttu-id="c7a4e-644">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c7a4e-644">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c7a4e-645">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-645">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="c7a4e-646">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-646">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="c7a4e-647">Objeto</span><span class="sxs-lookup"><span data-stu-id="c7a4e-647">Object</span></span>| <span data-ttu-id="c7a4e-648">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-648">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a4e-649">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-649">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a4e-650">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-650">Requirements</span></span>

|<span data-ttu-id="c7a4e-651">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-651">Requirement</span></span>| <span data-ttu-id="c7a4e-652">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-652">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-653">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-653">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-654">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a4e-654">1.0</span></span>|
|[<span data-ttu-id="c7a4e-655">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-655">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-656">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="c7a4e-656">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="c7a4e-657">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c7a4e-657">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-658">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-658">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a4e-659">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-659">Example</span></span>

<span data-ttu-id="c7a4e-660">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-660">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="c7a4e-661">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c7a4e-661">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="c7a4e-662">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-662">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="c7a4e-663">Atualmente, os tipos de eventos com `Office.EventType.ItemChanged` suporte `Office.EventType.OfficeThemeChanged`são e.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-663">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a4e-664">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="c7a4e-664">Parameters</span></span>

| <span data-ttu-id="c7a4e-665">Nome</span><span class="sxs-lookup"><span data-stu-id="c7a4e-665">Name</span></span> | <span data-ttu-id="c7a4e-666">Tipo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-666">Type</span></span> | <span data-ttu-id="c7a4e-667">Atributos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-667">Attributes</span></span> | <span data-ttu-id="c7a4e-668">Descrição</span><span class="sxs-lookup"><span data-stu-id="c7a4e-668">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c7a4e-669">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c7a4e-669">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c7a4e-670">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-670">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="c7a4e-671">Objeto</span><span class="sxs-lookup"><span data-stu-id="c7a4e-671">Object</span></span> | <span data-ttu-id="c7a4e-672">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-672">&lt;optional&gt;</span></span> | <span data-ttu-id="c7a4e-673">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-673">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c7a4e-674">Objeto</span><span class="sxs-lookup"><span data-stu-id="c7a4e-674">Object</span></span> | <span data-ttu-id="c7a4e-675">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-675">&lt;optional&gt;</span></span> | <span data-ttu-id="c7a4e-676">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="c7a4e-676">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c7a4e-677">function</span><span class="sxs-lookup"><span data-stu-id="c7a4e-677">function</span></span>| <span data-ttu-id="c7a4e-678">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a4e-678">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a4e-679">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c7a4e-679">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a4e-680">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c7a4e-680">Requirements</span></span>

|<span data-ttu-id="c7a4e-681">Requisito</span><span class="sxs-lookup"><span data-stu-id="c7a4e-681">Requirement</span></span>| <span data-ttu-id="c7a4e-682">Valor</span><span class="sxs-lookup"><span data-stu-id="c7a4e-682">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a4e-683">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c7a4e-683">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a4e-684">1,5</span><span class="sxs-lookup"><span data-stu-id="c7a4e-684">1.5</span></span> |
|[<span data-ttu-id="c7a4e-685">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c7a4e-685">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a4e-686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a4e-686">ReadItem</span></span> |
|[<span data-ttu-id="c7a4e-687">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="c7a4e-687">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a4e-688">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c7a4e-688">Compose or Read</span></span>|
