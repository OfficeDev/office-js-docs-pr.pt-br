---
title: Office. Context. Mailbox-visualização do conjunto de requisitos
description: ''
ms.date: 04/17/2019
localization_priority: Normal
ms.openlocfilehash: 557dedf3943be12fbb9e384873d0b9079b251c2f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450517"
---
# <a name="mailbox"></a><span data-ttu-id="7a42c-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="7a42c-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="7a42c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="7a42c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="7a42c-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="7a42c-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a42c-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-105">Requirements</span></span>

|<span data-ttu-id="7a42c-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-106">Requirement</span></span>| <span data-ttu-id="7a42c-107">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="7a42c-109">1.0</span></span>|
|[<span data-ttu-id="7a42c-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="7a42c-111">Restricted</span></span>|
|[<span data-ttu-id="7a42c-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7a42c-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7a42c-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="7a42c-114">Members and methods</span></span>

| <span data-ttu-id="7a42c-115">Membro</span><span class="sxs-lookup"><span data-stu-id="7a42c-115">Member</span></span> | <span data-ttu-id="7a42c-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7a42c-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="7a42c-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="7a42c-118">Membro</span><span class="sxs-lookup"><span data-stu-id="7a42c-118">Member</span></span> |
| [<span data-ttu-id="7a42c-119">Nova mastercategories</span><span class="sxs-lookup"><span data-stu-id="7a42c-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="7a42c-120">Membro</span><span class="sxs-lookup"><span data-stu-id="7a42c-120">Member</span></span> |
| [<span data-ttu-id="7a42c-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="7a42c-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="7a42c-122">Membro</span><span class="sxs-lookup"><span data-stu-id="7a42c-122">Member</span></span> |
| [<span data-ttu-id="7a42c-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="7a42c-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="7a42c-124">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-124">Method</span></span> |
| [<span data-ttu-id="7a42c-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="7a42c-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="7a42c-126">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-126">Method</span></span> |
| [<span data-ttu-id="7a42c-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="7a42c-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="7a42c-128">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-128">Method</span></span> |
| [<span data-ttu-id="7a42c-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="7a42c-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="7a42c-130">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-130">Method</span></span> |
| [<span data-ttu-id="7a42c-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="7a42c-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="7a42c-132">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-132">Method</span></span> |
| [<span data-ttu-id="7a42c-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="7a42c-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="7a42c-134">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-134">Method</span></span> |
| [<span data-ttu-id="7a42c-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="7a42c-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="7a42c-136">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-136">Method</span></span> |
| [<span data-ttu-id="7a42c-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="7a42c-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="7a42c-138">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-138">Method</span></span> |
| [<span data-ttu-id="7a42c-139">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="7a42c-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="7a42c-140">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-140">Method</span></span> |
| [<span data-ttu-id="7a42c-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="7a42c-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="7a42c-142">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-142">Method</span></span> |
| [<span data-ttu-id="7a42c-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="7a42c-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="7a42c-144">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-144">Method</span></span> |
| [<span data-ttu-id="7a42c-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="7a42c-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="7a42c-146">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-146">Method</span></span> |
| [<span data-ttu-id="7a42c-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="7a42c-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="7a42c-148">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-148">Method</span></span> |
| [<span data-ttu-id="7a42c-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="7a42c-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="7a42c-150">Método</span><span class="sxs-lookup"><span data-stu-id="7a42c-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="7a42c-151">Namespaces</span><span class="sxs-lookup"><span data-stu-id="7a42c-151">Namespaces</span></span>

<span data-ttu-id="7a42c-152">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="7a42c-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="7a42c-153">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="7a42c-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="7a42c-154">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="7a42c-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="7a42c-155">Membros</span><span class="sxs-lookup"><span data-stu-id="7a42c-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="7a42c-156">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="7a42c-156">ewsUrl :String</span></span>

<span data-ttu-id="7a42c-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7a42c-159">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="7a42c-159">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7a42c-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="7a42c-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="7a42c-162">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="7a42c-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="7a42c-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="7a42c-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-165">Type</span></span>

*   <span data-ttu-id="7a42c-166">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a42c-167">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-167">Requirements</span></span>

|<span data-ttu-id="7a42c-168">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-168">Requirement</span></span>| <span data-ttu-id="7a42c-169">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-170">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-171">1.0</span><span class="sxs-lookup"><span data-stu-id="7a42c-171">1.0</span></span>|
|[<span data-ttu-id="7a42c-172">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-173">ReadItem</span></span>|
|[<span data-ttu-id="7a42c-174">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7a42c-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-175">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-175">Compose or Read</span></span>|

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="7a42c-176">Nova mastercategories:[nova mastercategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="7a42c-176">masterCategories :[MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="7a42c-177">Obtém um objeto que fornece métodos para gerenciar a lista mestra de categorias nesta caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="7a42c-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="7a42c-178">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="7a42c-178">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="7a42c-179">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-179">Type</span></span>

*   [<span data-ttu-id="7a42c-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="7a42c-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="7a42c-181">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-181">Requirements</span></span>

|<span data-ttu-id="7a42c-182">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-182">Requirement</span></span>| <span data-ttu-id="7a42c-183">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-184">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-185">Visualização</span><span class="sxs-lookup"><span data-stu-id="7a42c-185">Preview</span></span> |
|[<span data-ttu-id="7a42c-186">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="7a42c-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="7a42c-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7a42c-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="7a42c-190">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7a42c-190">Example</span></span>

<span data-ttu-id="7a42c-191">Este exemplo obtém a lista mestra de categorias para esta caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="7a42c-191">This example gets the categories master list for this mailbox.</span></span>

```javascript
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

---
---

#### <a name="resturl-string"></a><span data-ttu-id="7a42c-192">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="7a42c-192">restUrl :String</span></span>

<span data-ttu-id="7a42c-193">Obtém a URL do ponto de extremidade de REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="7a42c-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="7a42c-194">O valor `restUrl` pode ser usado para fazer chamadas da [API REST](/outlook/rest/) para a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="7a42c-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="7a42c-195">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `restUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="7a42c-195">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="7a42c-p104">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `restUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="7a42c-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-198">Type</span></span>

*   <span data-ttu-id="7a42c-199">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a42c-200">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-200">Requirements</span></span>

|<span data-ttu-id="7a42c-201">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-201">Requirement</span></span>| <span data-ttu-id="7a42c-202">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-203">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-204">1,5</span><span class="sxs-lookup"><span data-stu-id="7a42c-204">1.5</span></span> |
|[<span data-ttu-id="7a42c-205">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-206">ReadItem</span></span>|
|[<span data-ttu-id="7a42c-207">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7a42c-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-208">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-208">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="7a42c-209">Métodos</span><span class="sxs-lookup"><span data-stu-id="7a42c-209">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="7a42c-210">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7a42c-210">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="7a42c-211">Adiciona um manipulador de eventos a um evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="7a42c-211">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="7a42c-212">Atualmente, os tipos de eventos com `Office.EventType.ItemChanged` suporte `Office.EventType.OfficeThemeChanged`são e.</span><span class="sxs-lookup"><span data-stu-id="7a42c-212">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-213">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-213">Parameters</span></span>

| <span data-ttu-id="7a42c-214">Nome</span><span class="sxs-lookup"><span data-stu-id="7a42c-214">Name</span></span> | <span data-ttu-id="7a42c-215">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-215">Type</span></span> | <span data-ttu-id="7a42c-216">Atributos</span><span class="sxs-lookup"><span data-stu-id="7a42c-216">Attributes</span></span> | <span data-ttu-id="7a42c-217">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-217">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="7a42c-218">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="7a42c-218">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="7a42c-219">O evento que deve invocar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="7a42c-219">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="7a42c-220">Função</span><span class="sxs-lookup"><span data-stu-id="7a42c-220">Function</span></span> || <span data-ttu-id="7a42c-p105">A função para manipular o evento. A função deve aceitar um parâmetro exclusivo, que é um objeto literal. A propriedade `type` no parâmetro corresponderá ao parâmetro `eventType` passado para `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="7a42c-224">Objeto</span><span class="sxs-lookup"><span data-stu-id="7a42c-224">Object</span></span> | <span data-ttu-id="7a42c-225">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-225">&lt;optional&gt;</span></span> | <span data-ttu-id="7a42c-226">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="7a42c-226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="7a42c-227">Objeto</span><span class="sxs-lookup"><span data-stu-id="7a42c-227">Object</span></span> | <span data-ttu-id="7a42c-228">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-228">&lt;optional&gt;</span></span> | <span data-ttu-id="7a42c-229">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7a42c-229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="7a42c-230">function</span><span class="sxs-lookup"><span data-stu-id="7a42c-230">function</span></span>| <span data-ttu-id="7a42c-231">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-231">&lt;optional&gt;</span></span>|<span data-ttu-id="7a42c-232">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a42c-232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a42c-233">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-233">Requirements</span></span>

|<span data-ttu-id="7a42c-234">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-234">Requirement</span></span>| <span data-ttu-id="7a42c-235">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-236">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-237">1,5</span><span class="sxs-lookup"><span data-stu-id="7a42c-237">1.5</span></span> |
|[<span data-ttu-id="7a42c-238">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-239">ReadItem</span></span> |
|[<span data-ttu-id="7a42c-240">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7a42c-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-241">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a42c-242">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7a42c-242">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="7a42c-243">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="7a42c-243">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="7a42c-244">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="7a42c-244">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="7a42c-245">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="7a42c-245">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7a42c-p106">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-248">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-248">Parameters</span></span>

|<span data-ttu-id="7a42c-249">Nome</span><span class="sxs-lookup"><span data-stu-id="7a42c-249">Name</span></span>| <span data-ttu-id="7a42c-250">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-250">Type</span></span>| <span data-ttu-id="7a42c-251">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-251">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="7a42c-252">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-252">String</span></span>|<span data-ttu-id="7a42c-253">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="7a42c-253">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="7a42c-254">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="7a42c-254">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="7a42c-255">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="7a42c-255">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a42c-256">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-256">Requirements</span></span>

|<span data-ttu-id="7a42c-257">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-257">Requirement</span></span>| <span data-ttu-id="7a42c-258">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-259">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-260">1.3</span><span class="sxs-lookup"><span data-stu-id="7a42c-260">1.3</span></span>|
|[<span data-ttu-id="7a42c-261">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-262">Restrito</span><span class="sxs-lookup"><span data-stu-id="7a42c-262">Restricted</span></span>|
|[<span data-ttu-id="7a42c-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7a42c-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-264">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7a42c-265">Retorna:</span><span class="sxs-lookup"><span data-stu-id="7a42c-265">Returns:</span></span>

<span data-ttu-id="7a42c-266">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="7a42c-266">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="7a42c-267">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7a42c-267">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="7a42c-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="7a42c-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="7a42c-269">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="7a42c-269">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="7a42c-p107">As datas e horas usadas por um aplicativo de email para o Outlook ou o Outlook Web App podem usar fusos horários diferentes. O Outlook usa o fuso horário do computador cliente; o Outlook Web App usa o fuso horário definido na Centro de administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p107">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="7a42c-p108">Se o aplicativo de email estiver sendo executado no Outlook, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador do cliente. Se o aplicativo de email estiver sendo executado no Outlook Web App, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p108">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-275">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-275">Parameters</span></span>

|<span data-ttu-id="7a42c-276">Nome</span><span class="sxs-lookup"><span data-stu-id="7a42c-276">Name</span></span>| <span data-ttu-id="7a42c-277">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-277">Type</span></span>| <span data-ttu-id="7a42c-278">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-278">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="7a42c-279">Date</span><span class="sxs-lookup"><span data-stu-id="7a42c-279">Date</span></span>|<span data-ttu-id="7a42c-280">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="7a42c-280">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a42c-281">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-281">Requirements</span></span>

|<span data-ttu-id="7a42c-282">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-282">Requirement</span></span>| <span data-ttu-id="7a42c-283">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-284">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-285">1.0</span><span class="sxs-lookup"><span data-stu-id="7a42c-285">1.0</span></span>|
|[<span data-ttu-id="7a42c-286">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-287">ReadItem</span></span>|
|[<span data-ttu-id="7a42c-288">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7a42c-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-289">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-289">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7a42c-290">Retorna:</span><span class="sxs-lookup"><span data-stu-id="7a42c-290">Returns:</span></span>

<span data-ttu-id="7a42c-291">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="7a42c-291">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

---
---

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="7a42c-292">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="7a42c-292">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="7a42c-293">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="7a42c-293">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="7a42c-294">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="7a42c-294">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7a42c-p109">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-297">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-297">Parameters</span></span>

|<span data-ttu-id="7a42c-298">Nome</span><span class="sxs-lookup"><span data-stu-id="7a42c-298">Name</span></span>| <span data-ttu-id="7a42c-299">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-299">Type</span></span>| <span data-ttu-id="7a42c-300">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-300">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="7a42c-301">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-301">String</span></span>|<span data-ttu-id="7a42c-302">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="7a42c-302">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="7a42c-303">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="7a42c-303">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="7a42c-304">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="7a42c-304">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a42c-305">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-305">Requirements</span></span>

|<span data-ttu-id="7a42c-306">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-306">Requirement</span></span>| <span data-ttu-id="7a42c-307">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-308">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-309">1.3</span><span class="sxs-lookup"><span data-stu-id="7a42c-309">1.3</span></span>|
|[<span data-ttu-id="7a42c-310">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-311">Restrito</span><span class="sxs-lookup"><span data-stu-id="7a42c-311">Restricted</span></span>|
|[<span data-ttu-id="7a42c-312">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7a42c-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-313">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7a42c-314">Retorna:</span><span class="sxs-lookup"><span data-stu-id="7a42c-314">Returns:</span></span>

<span data-ttu-id="7a42c-315">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="7a42c-315">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="7a42c-316">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7a42c-316">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="7a42c-317">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="7a42c-317">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="7a42c-318">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="7a42c-318">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="7a42c-319">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="7a42c-319">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-320">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-320">Parameters</span></span>

|<span data-ttu-id="7a42c-321">Nome</span><span class="sxs-lookup"><span data-stu-id="7a42c-321">Name</span></span>| <span data-ttu-id="7a42c-322">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-322">Type</span></span>| <span data-ttu-id="7a42c-323">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-323">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="7a42c-324">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="7a42c-324">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="7a42c-325">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="7a42c-325">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a42c-326">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-326">Requirements</span></span>

|<span data-ttu-id="7a42c-327">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-327">Requirement</span></span>| <span data-ttu-id="7a42c-328">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-329">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-330">1.0</span><span class="sxs-lookup"><span data-stu-id="7a42c-330">1.0</span></span>|
|[<span data-ttu-id="7a42c-331">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-332">ReadItem</span></span>|
|[<span data-ttu-id="7a42c-333">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7a42c-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-334">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-334">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7a42c-335">Retorna:</span><span class="sxs-lookup"><span data-stu-id="7a42c-335">Returns:</span></span>

<span data-ttu-id="7a42c-336">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="7a42c-336">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="7a42c-337">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="7a42c-337">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="7a42c-338">Date</span><span class="sxs-lookup"><span data-stu-id="7a42c-338">Date</span></span></dd>

</dl>

---
---

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="7a42c-339">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="7a42c-339">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="7a42c-340">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="7a42c-340">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7a42c-341">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="7a42c-341">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7a42c-342">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="7a42c-342">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="7a42c-p110">No Outlook para Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir uma instância da série. Isso ocorre porque no Outlook para Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p110">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="7a42c-345">No Outlook Web App, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="7a42c-345">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="7a42c-346">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="7a42c-346">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-347">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-347">Parameters</span></span>

|<span data-ttu-id="7a42c-348">Nome</span><span class="sxs-lookup"><span data-stu-id="7a42c-348">Name</span></span>| <span data-ttu-id="7a42c-349">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-349">Type</span></span>| <span data-ttu-id="7a42c-350">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-350">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="7a42c-351">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-351">String</span></span>|<span data-ttu-id="7a42c-352">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="7a42c-352">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a42c-353">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-353">Requirements</span></span>

|<span data-ttu-id="7a42c-354">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-354">Requirement</span></span>| <span data-ttu-id="7a42c-355">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-356">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-357">1.0</span><span class="sxs-lookup"><span data-stu-id="7a42c-357">1.0</span></span>|
|[<span data-ttu-id="7a42c-358">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-359">ReadItem</span></span>|
|[<span data-ttu-id="7a42c-360">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7a42c-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-361">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-361">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a42c-362">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7a42c-362">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

####  <a name="displaymessageformitemid"></a><span data-ttu-id="7a42c-363">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="7a42c-363">displayMessageForm(itemId)</span></span>

<span data-ttu-id="7a42c-364">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="7a42c-364">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="7a42c-365">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="7a42c-365">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7a42c-366">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="7a42c-366">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="7a42c-367">No Outlook Web App, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="7a42c-367">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="7a42c-368">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="7a42c-368">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="7a42c-p111">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-371">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-371">Parameters</span></span>

|<span data-ttu-id="7a42c-372">Nome</span><span class="sxs-lookup"><span data-stu-id="7a42c-372">Name</span></span>| <span data-ttu-id="7a42c-373">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-373">Type</span></span>| <span data-ttu-id="7a42c-374">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-374">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="7a42c-375">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-375">String</span></span>|<span data-ttu-id="7a42c-376">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="7a42c-376">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a42c-377">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-377">Requirements</span></span>

|<span data-ttu-id="7a42c-378">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-378">Requirement</span></span>| <span data-ttu-id="7a42c-379">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-380">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-381">1.0</span><span class="sxs-lookup"><span data-stu-id="7a42c-381">1.0</span></span>|
|[<span data-ttu-id="7a42c-382">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-382">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-383">ReadItem</span></span>|
|[<span data-ttu-id="7a42c-384">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7a42c-384">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-385">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-385">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a42c-386">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7a42c-386">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="7a42c-387">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="7a42c-387">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="7a42c-388">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="7a42c-388">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7a42c-389">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="7a42c-389">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7a42c-p112">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="7a42c-p113">No Outlook Web App e no OWA para Dispositivos, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p113">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="7a42c-p114">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="7a42c-397">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="7a42c-397">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-398">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-398">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="7a42c-399">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="7a42c-399">All parameters are optional.</span></span>

|<span data-ttu-id="7a42c-400">Name</span><span class="sxs-lookup"><span data-stu-id="7a42c-400">Name</span></span>| <span data-ttu-id="7a42c-401">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-401">Type</span></span>| <span data-ttu-id="7a42c-402">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-402">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="7a42c-403">Object</span><span class="sxs-lookup"><span data-stu-id="7a42c-403">Object</span></span> | <span data-ttu-id="7a42c-404">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="7a42c-404">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="7a42c-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="7a42c-p115">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="7a42c-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="7a42c-p116">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="7a42c-411">Data</span><span class="sxs-lookup"><span data-stu-id="7a42c-411">Date</span></span> | <span data-ttu-id="7a42c-412">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="7a42c-412">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="7a42c-413">Data</span><span class="sxs-lookup"><span data-stu-id="7a42c-413">Date</span></span> | <span data-ttu-id="7a42c-414">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="7a42c-414">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="7a42c-415">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-415">String</span></span> | <span data-ttu-id="7a42c-p117">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="7a42c-418">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-418">Array.&lt;String&gt;</span></span> | <span data-ttu-id="7a42c-p118">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="7a42c-421">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-421">String</span></span> | <span data-ttu-id="7a42c-p119">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="7a42c-424">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-424">String</span></span> | <span data-ttu-id="7a42c-p120">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7a42c-427">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-427">Requirements</span></span>

|<span data-ttu-id="7a42c-428">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-428">Requirement</span></span>| <span data-ttu-id="7a42c-429">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-430">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-431">1.0</span><span class="sxs-lookup"><span data-stu-id="7a42c-431">1.0</span></span>|
|[<span data-ttu-id="7a42c-432">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-433">ReadItem</span></span>|
|[<span data-ttu-id="7a42c-434">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7a42c-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-435">Read</span><span class="sxs-lookup"><span data-stu-id="7a42c-435">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a42c-436">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7a42c-436">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="7a42c-437">displayNewMessageForm (parâmetros)</span><span class="sxs-lookup"><span data-stu-id="7a42c-437">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="7a42c-438">Exibe um formulário para criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="7a42c-438">Displays a form for creating a new message.</span></span>

<span data-ttu-id="7a42c-439">O `displayNewMessageForm` método abre um formulário que permite ao usuário criar uma nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="7a42c-439">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="7a42c-440">Se os parâmetros forem especificados, os campos de formulário da mensagem serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="7a42c-440">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="7a42c-441">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="7a42c-441">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-442">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-442">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="7a42c-443">Todos os parâmetros são opcionais.</span><span class="sxs-lookup"><span data-stu-id="7a42c-443">All parameters are optional.</span></span>

|<span data-ttu-id="7a42c-444">Name</span><span class="sxs-lookup"><span data-stu-id="7a42c-444">Name</span></span>| <span data-ttu-id="7a42c-445">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-445">Type</span></span>| <span data-ttu-id="7a42c-446">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-446">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="7a42c-447">Objeto</span><span class="sxs-lookup"><span data-stu-id="7a42c-447">Object</span></span> | <span data-ttu-id="7a42c-448">Um dicionário de parâmetros que descreve a nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="7a42c-448">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="7a42c-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="7a42c-450">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha para.</span><span class="sxs-lookup"><span data-stu-id="7a42c-450">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="7a42c-451">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="7a42c-451">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="7a42c-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="7a42c-453">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha CC.</span><span class="sxs-lookup"><span data-stu-id="7a42c-453">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="7a42c-454">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="7a42c-454">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="7a42c-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="7a42c-456">Uma matriz de cadeias de caracteres que contém os endereços de email `EmailAddressDetails` ou uma matriz que contém um objeto para cada um dos destinatários na linha Cco.</span><span class="sxs-lookup"><span data-stu-id="7a42c-456">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="7a42c-457">A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="7a42c-457">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="7a42c-458">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-458">String</span></span> | <span data-ttu-id="7a42c-459">Uma cadeia de caracteres que contém o assunto da mensagem.</span><span class="sxs-lookup"><span data-stu-id="7a42c-459">A string containing the subject of the message.</span></span> <span data-ttu-id="7a42c-460">A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="7a42c-460">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="7a42c-461">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-461">String</span></span> | <span data-ttu-id="7a42c-462">O corpo HTML da mensagem.</span><span class="sxs-lookup"><span data-stu-id="7a42c-462">The HTML body of the message.</span></span> <span data-ttu-id="7a42c-463">O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="7a42c-463">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="7a42c-464">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-464">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7a42c-465">Uma matriz de objetos JSON que são anexos de arquivo ou item.</span><span class="sxs-lookup"><span data-stu-id="7a42c-465">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="7a42c-466">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-466">String</span></span> | <span data-ttu-id="7a42c-p127">Indica o tipo de anexo. Deve ser `file` para um anexo de arquivo ou `item` para um anexo de item.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="7a42c-469">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-469">String</span></span> | <span data-ttu-id="7a42c-470">Uma cadeia de caracteres que contém o nome do anexo, até 255 caracteres de comprimento.</span><span class="sxs-lookup"><span data-stu-id="7a42c-470">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="7a42c-471">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-471">String</span></span> | <span data-ttu-id="7a42c-p128">Usado somente se `type` estiver definido como `file`. O URI do local para o arquivo.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="7a42c-474">Booliano</span><span class="sxs-lookup"><span data-stu-id="7a42c-474">Boolean</span></span> | <span data-ttu-id="7a42c-p129">Usado somente se `type` estiver definido como `file`. Se for `true`, indicará que o anexo será mostrado embutido no corpo da mensagem e não deverá ser exibido na lista de anexos.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="7a42c-477">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7a42c-477">String</span></span> | <span data-ttu-id="7a42c-478">Usado somente se `type` estiver definido como `item`.</span><span class="sxs-lookup"><span data-stu-id="7a42c-478">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="7a42c-479">A ID do item do EWS do email existente que você deseja anexar à nova mensagem.</span><span class="sxs-lookup"><span data-stu-id="7a42c-479">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="7a42c-480">Isso é uma cadeia de até 100 caracteres.</span><span class="sxs-lookup"><span data-stu-id="7a42c-480">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="7a42c-481">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-481">Requirements</span></span>

|<span data-ttu-id="7a42c-482">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-482">Requirement</span></span>| <span data-ttu-id="7a42c-483">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-484">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-485">1.6</span><span class="sxs-lookup"><span data-stu-id="7a42c-485">1.6</span></span> |
|[<span data-ttu-id="7a42c-486">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-487">ReadItem</span></span>|
|[<span data-ttu-id="7a42c-488">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7a42c-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-489">Read</span><span class="sxs-lookup"><span data-stu-id="7a42c-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a42c-490">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7a42c-490">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="7a42c-491">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="7a42c-491">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="7a42c-492">Obtém uma cadeia de caracteres que contém um token usado para chamar APIs REST ou Serviços Web do Exchange.</span><span class="sxs-lookup"><span data-stu-id="7a42c-492">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="7a42c-p131">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="7a42c-495">É recomendável que suplementos usem as APIs REST em vez dos Serviços Web do Exchange sempre que possível.</span><span class="sxs-lookup"><span data-stu-id="7a42c-495">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="7a42c-496">**Tokens REST**</span><span class="sxs-lookup"><span data-stu-id="7a42c-496">**REST Tokens**</span></span>

<span data-ttu-id="7a42c-p132">Quando um token REST é solicitado (`options.isRest = true`), o token resultante não funcionará para autenticar as chamadas dos Serviços Web do Exchange. O token será limitado em escopo para acesso somente leitura no item atual e seus anexos, a menos que o suplemento tenha especificado a permissão [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) em seu manifesto. Se a permissão `ReadWriteMailbox` tiver sido especificada, o token resultante concederá acesso de leitura/gravação a email, calendário e contatos, incluindo a capacidade de enviar emails.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="7a42c-500">O suplemento deve usar a propriedade `restUrl` para determinar a URL correta a ser usada ao fazer chamadas da API REST.</span><span class="sxs-lookup"><span data-stu-id="7a42c-500">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="7a42c-501">**Tokens EWS**</span><span class="sxs-lookup"><span data-stu-id="7a42c-501">**EWS Tokens**</span></span>

<span data-ttu-id="7a42c-p133">Quando um token EWS é solicitado (`options.isRest = false`), o token resultante não funcionará para autenticar as chamadas de API REST. O token será limitado em escopo para acessar o item atual.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="7a42c-504">O suplemento deve usar a propriedade `ewsUrl` para determinar a URL correta a ser usada ao fazer chamadas de EWS.</span><span class="sxs-lookup"><span data-stu-id="7a42c-504">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-505">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-505">Parameters</span></span>

|<span data-ttu-id="7a42c-506">Nome</span><span class="sxs-lookup"><span data-stu-id="7a42c-506">Name</span></span>| <span data-ttu-id="7a42c-507">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-507">Type</span></span>| <span data-ttu-id="7a42c-508">Atributos</span><span class="sxs-lookup"><span data-stu-id="7a42c-508">Attributes</span></span>| <span data-ttu-id="7a42c-509">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-509">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="7a42c-510">Objeto</span><span class="sxs-lookup"><span data-stu-id="7a42c-510">Object</span></span> | <span data-ttu-id="7a42c-511">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-511">&lt;optional&gt;</span></span> | <span data-ttu-id="7a42c-512">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="7a42c-512">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="7a42c-513">Booliano</span><span class="sxs-lookup"><span data-stu-id="7a42c-513">Boolean</span></span> |  <span data-ttu-id="7a42c-514">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-514">&lt;optional&gt;</span></span> | <span data-ttu-id="7a42c-p134">Determina se o token fornecido será usado para as APIs REST do Outlook ou Serviços Web do Exchange. O valor padrão é `false`.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="7a42c-517">Objeto</span><span class="sxs-lookup"><span data-stu-id="7a42c-517">Object</span></span> |  <span data-ttu-id="7a42c-518">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-518">&lt;optional&gt;</span></span> | <span data-ttu-id="7a42c-519">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="7a42c-519">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="7a42c-520">function</span><span class="sxs-lookup"><span data-stu-id="7a42c-520">function</span></span>||<span data-ttu-id="7a42c-p135">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a42c-523">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-523">Requirements</span></span>

|<span data-ttu-id="7a42c-524">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-524">Requirement</span></span>| <span data-ttu-id="7a42c-525">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-525">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-526">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-526">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-527">1,5</span><span class="sxs-lookup"><span data-stu-id="7a42c-527">1.5</span></span> |
|[<span data-ttu-id="7a42c-528">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-528">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-529">ReadItem</span></span>|
|[<span data-ttu-id="7a42c-530">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7a42c-530">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-531">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="7a42c-531">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a42c-532">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7a42c-532">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="7a42c-533">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7a42c-533">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="7a42c-534">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="7a42c-534">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="7a42c-p136">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="7a42c-p137">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="7a42c-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="7a42c-540">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="7a42c-540">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="7a42c-p138">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-543">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-543">Parameters</span></span>

|<span data-ttu-id="7a42c-544">Nome</span><span class="sxs-lookup"><span data-stu-id="7a42c-544">Name</span></span>| <span data-ttu-id="7a42c-545">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-545">Type</span></span>| <span data-ttu-id="7a42c-546">Atributos</span><span class="sxs-lookup"><span data-stu-id="7a42c-546">Attributes</span></span>| <span data-ttu-id="7a42c-547">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-547">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="7a42c-548">function</span><span class="sxs-lookup"><span data-stu-id="7a42c-548">function</span></span>||<span data-ttu-id="7a42c-p139">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="7a42c-551">Objeto</span><span class="sxs-lookup"><span data-stu-id="7a42c-551">Object</span></span>| <span data-ttu-id="7a42c-552">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-552">&lt;optional&gt;</span></span>|<span data-ttu-id="7a42c-553">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="7a42c-553">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a42c-554">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-554">Requirements</span></span>

|<span data-ttu-id="7a42c-555">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-555">Requirement</span></span>| <span data-ttu-id="7a42c-556">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-557">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-558">1.3</span><span class="sxs-lookup"><span data-stu-id="7a42c-558">1.3</span></span>|
|[<span data-ttu-id="7a42c-559">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-560">ReadItem</span></span>|
|[<span data-ttu-id="7a42c-561">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7a42c-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-562">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="7a42c-562">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a42c-563">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7a42c-563">Example</span></span>

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

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="7a42c-564">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7a42c-564">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="7a42c-565">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="7a42c-565">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="7a42c-566">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="7a42c-566">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-567">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-567">Parameters</span></span>

|<span data-ttu-id="7a42c-568">Nome</span><span class="sxs-lookup"><span data-stu-id="7a42c-568">Name</span></span>| <span data-ttu-id="7a42c-569">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-569">Type</span></span>| <span data-ttu-id="7a42c-570">Atributos</span><span class="sxs-lookup"><span data-stu-id="7a42c-570">Attributes</span></span>| <span data-ttu-id="7a42c-571">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-571">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="7a42c-572">function</span><span class="sxs-lookup"><span data-stu-id="7a42c-572">function</span></span>||<span data-ttu-id="7a42c-573">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a42c-573">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7a42c-574">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7a42c-574">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="7a42c-575">Object</span><span class="sxs-lookup"><span data-stu-id="7a42c-575">Object</span></span>| <span data-ttu-id="7a42c-576">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-576">&lt;optional&gt;</span></span>|<span data-ttu-id="7a42c-577">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="7a42c-577">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a42c-578">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-578">Requirements</span></span>

|<span data-ttu-id="7a42c-579">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-579">Requirement</span></span>| <span data-ttu-id="7a42c-580">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-581">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-582">1.0</span><span class="sxs-lookup"><span data-stu-id="7a42c-582">1.0</span></span>|
|[<span data-ttu-id="7a42c-583">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-583">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-584">ReadItem</span></span>|
|[<span data-ttu-id="7a42c-585">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7a42c-585">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-586">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-586">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a42c-587">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7a42c-587">Example</span></span>

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

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="7a42c-588">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7a42c-588">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="7a42c-589">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="7a42c-589">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="7a42c-590">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="7a42c-590">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="7a42c-591">No Outlook para iOS ou no Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="7a42c-591">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="7a42c-592">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="7a42c-592">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="7a42c-593">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="7a42c-593">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="7a42c-594">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="7a42c-594">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="7a42c-595">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="7a42c-595">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="7a42c-596">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="7a42c-596">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="7a42c-597">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="7a42c-597">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="7a42c-p141">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="7a42c-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="7a42c-600">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="7a42c-600">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="7a42c-601">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="7a42c-601">Version differences</span></span>

<span data-ttu-id="7a42c-602">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="7a42c-602">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="7a42c-p142">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="7a42c-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-606">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-606">Parameters</span></span>

|<span data-ttu-id="7a42c-607">Nome</span><span class="sxs-lookup"><span data-stu-id="7a42c-607">Name</span></span>| <span data-ttu-id="7a42c-608">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-608">Type</span></span>| <span data-ttu-id="7a42c-609">Atributos</span><span class="sxs-lookup"><span data-stu-id="7a42c-609">Attributes</span></span>| <span data-ttu-id="7a42c-610">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-610">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="7a42c-611">String</span><span class="sxs-lookup"><span data-stu-id="7a42c-611">String</span></span>||<span data-ttu-id="7a42c-612">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="7a42c-612">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="7a42c-613">function</span><span class="sxs-lookup"><span data-stu-id="7a42c-613">function</span></span>||<span data-ttu-id="7a42c-614">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a42c-614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7a42c-615">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7a42c-615">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="7a42c-616">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="7a42c-616">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="7a42c-617">Objeto</span><span class="sxs-lookup"><span data-stu-id="7a42c-617">Object</span></span>| <span data-ttu-id="7a42c-618">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-618">&lt;optional&gt;</span></span>|<span data-ttu-id="7a42c-619">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="7a42c-619">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a42c-620">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-620">Requirements</span></span>

|<span data-ttu-id="7a42c-621">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-621">Requirement</span></span>| <span data-ttu-id="7a42c-622">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-623">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-623">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-624">1.0</span><span class="sxs-lookup"><span data-stu-id="7a42c-624">1.0</span></span>|
|[<span data-ttu-id="7a42c-625">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-625">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-626">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="7a42c-626">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="7a42c-627">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7a42c-627">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-628">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-628">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a42c-629">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7a42c-629">Example</span></span>

<span data-ttu-id="7a42c-630">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="7a42c-630">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="7a42c-631">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7a42c-631">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="7a42c-632">Remove um manipulador de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="7a42c-632">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="7a42c-633">Atualmente, os tipos de eventos com `Office.EventType.ItemChanged` suporte `Office.EventType.OfficeThemeChanged`são e.</span><span class="sxs-lookup"><span data-stu-id="7a42c-633">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7a42c-634">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="7a42c-634">Parameters</span></span>

| <span data-ttu-id="7a42c-635">Nome</span><span class="sxs-lookup"><span data-stu-id="7a42c-635">Name</span></span> | <span data-ttu-id="7a42c-636">Tipo</span><span class="sxs-lookup"><span data-stu-id="7a42c-636">Type</span></span> | <span data-ttu-id="7a42c-637">Atributos</span><span class="sxs-lookup"><span data-stu-id="7a42c-637">Attributes</span></span> | <span data-ttu-id="7a42c-638">Descrição</span><span class="sxs-lookup"><span data-stu-id="7a42c-638">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="7a42c-639">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="7a42c-639">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="7a42c-640">O evento que deve revogar o manipulador.</span><span class="sxs-lookup"><span data-stu-id="7a42c-640">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="7a42c-641">Objeto</span><span class="sxs-lookup"><span data-stu-id="7a42c-641">Object</span></span> | <span data-ttu-id="7a42c-642">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-642">&lt;optional&gt;</span></span> | <span data-ttu-id="7a42c-643">Um objeto literal que contém uma ou mais das propriedades a seguir.</span><span class="sxs-lookup"><span data-stu-id="7a42c-643">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="7a42c-644">Objeto</span><span class="sxs-lookup"><span data-stu-id="7a42c-644">Object</span></span> | <span data-ttu-id="7a42c-645">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-645">&lt;optional&gt;</span></span> | <span data-ttu-id="7a42c-646">Os desenvolvedores podem fornecer qualquer objeto que desejarem acessar no método de retorno de chamada.</span><span class="sxs-lookup"><span data-stu-id="7a42c-646">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="7a42c-647">function</span><span class="sxs-lookup"><span data-stu-id="7a42c-647">function</span></span>| <span data-ttu-id="7a42c-648">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="7a42c-648">&lt;optional&gt;</span></span>|<span data-ttu-id="7a42c-649">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7a42c-649">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7a42c-650">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7a42c-650">Requirements</span></span>

|<span data-ttu-id="7a42c-651">Requisito</span><span class="sxs-lookup"><span data-stu-id="7a42c-651">Requirement</span></span>| <span data-ttu-id="7a42c-652">Valor</span><span class="sxs-lookup"><span data-stu-id="7a42c-652">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a42c-653">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7a42c-653">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7a42c-654">1,5</span><span class="sxs-lookup"><span data-stu-id="7a42c-654">1.5</span></span> |
|[<span data-ttu-id="7a42c-655">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7a42c-655">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7a42c-656">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7a42c-656">ReadItem</span></span> |
|[<span data-ttu-id="7a42c-657">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="7a42c-657">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7a42c-658">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7a42c-658">Compose or Read</span></span>|
