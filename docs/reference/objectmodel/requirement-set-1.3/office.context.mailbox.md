---
title: Office. Context. Mailbox – conjunto de requisitos 1,3
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: f1896803c38abd03f63b0a9ae689d91eeb5540de
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627017"
---
# <a name="mailbox"></a><span data-ttu-id="e0056-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="e0056-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="e0056-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="e0056-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="e0056-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="e0056-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e0056-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e0056-105">Requirements</span></span>

|<span data-ttu-id="e0056-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="e0056-106">Requirement</span></span>| <span data-ttu-id="e0056-107">Valor</span><span class="sxs-lookup"><span data-stu-id="e0056-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0056-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e0056-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e0056-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e0056-109">1.0</span></span>|
|[<span data-ttu-id="e0056-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e0056-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e0056-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="e0056-111">Restricted</span></span>|
|[<span data-ttu-id="e0056-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e0056-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e0056-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e0056-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e0056-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="e0056-114">Members and methods</span></span>

| <span data-ttu-id="e0056-115">Membro</span><span class="sxs-lookup"><span data-stu-id="e0056-115">Member</span></span> | <span data-ttu-id="e0056-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="e0056-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e0056-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="e0056-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="e0056-118">Membro</span><span class="sxs-lookup"><span data-stu-id="e0056-118">Member</span></span> |
| [<span data-ttu-id="e0056-119">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="e0056-119">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="e0056-120">Método</span><span class="sxs-lookup"><span data-stu-id="e0056-120">Method</span></span> |
| [<span data-ttu-id="e0056-121">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e0056-121">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="e0056-122">Método</span><span class="sxs-lookup"><span data-stu-id="e0056-122">Method</span></span> |
| [<span data-ttu-id="e0056-123">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="e0056-123">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="e0056-124">Método</span><span class="sxs-lookup"><span data-stu-id="e0056-124">Method</span></span> |
| [<span data-ttu-id="e0056-125">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="e0056-125">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="e0056-126">Método</span><span class="sxs-lookup"><span data-stu-id="e0056-126">Method</span></span> |
| [<span data-ttu-id="e0056-127">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e0056-127">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="e0056-128">Método</span><span class="sxs-lookup"><span data-stu-id="e0056-128">Method</span></span> |
| [<span data-ttu-id="e0056-129">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="e0056-129">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="e0056-130">Método</span><span class="sxs-lookup"><span data-stu-id="e0056-130">Method</span></span> |
| [<span data-ttu-id="e0056-131">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e0056-131">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="e0056-132">Método</span><span class="sxs-lookup"><span data-stu-id="e0056-132">Method</span></span> |
| [<span data-ttu-id="e0056-133">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e0056-133">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="e0056-134">Método</span><span class="sxs-lookup"><span data-stu-id="e0056-134">Method</span></span> |
| [<span data-ttu-id="e0056-135">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e0056-135">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="e0056-136">Método</span><span class="sxs-lookup"><span data-stu-id="e0056-136">Method</span></span> |
| [<span data-ttu-id="e0056-137">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="e0056-137">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="e0056-138">Método</span><span class="sxs-lookup"><span data-stu-id="e0056-138">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="e0056-139">Namespaces</span><span class="sxs-lookup"><span data-stu-id="e0056-139">Namespaces</span></span>

<span data-ttu-id="e0056-140">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e0056-140">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="e0056-141">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e0056-141">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="e0056-142">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e0056-142">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="e0056-143">Members</span><span class="sxs-lookup"><span data-stu-id="e0056-143">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="e0056-144">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="e0056-144">ewsUrl: String</span></span>

<span data-ttu-id="e0056-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="e0056-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e0056-147">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="e0056-147">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e0056-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="e0056-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e0056-150">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="e0056-150">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="e0056-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="e0056-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="e0056-153">Tipo</span><span class="sxs-lookup"><span data-stu-id="e0056-153">Type</span></span>

*   <span data-ttu-id="e0056-154">String</span><span class="sxs-lookup"><span data-stu-id="e0056-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e0056-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e0056-155">Requirements</span></span>

|<span data-ttu-id="e0056-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="e0056-156">Requirement</span></span>| <span data-ttu-id="e0056-157">Valor</span><span class="sxs-lookup"><span data-stu-id="e0056-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0056-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e0056-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e0056-159">1.0</span><span class="sxs-lookup"><span data-stu-id="e0056-159">1.0</span></span>|
|[<span data-ttu-id="e0056-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e0056-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e0056-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e0056-161">ReadItem</span></span>|
|[<span data-ttu-id="e0056-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e0056-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e0056-163">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e0056-163">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="e0056-164">Métodos</span><span class="sxs-lookup"><span data-stu-id="e0056-164">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="e0056-165">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e0056-165">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e0056-166">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="e0056-166">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="e0056-167">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="e0056-167">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e0056-p104">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="e0056-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e0056-170">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e0056-170">Parameters</span></span>

|<span data-ttu-id="e0056-171">Nome</span><span class="sxs-lookup"><span data-stu-id="e0056-171">Name</span></span>| <span data-ttu-id="e0056-172">Tipo</span><span class="sxs-lookup"><span data-stu-id="e0056-172">Type</span></span>| <span data-ttu-id="e0056-173">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0056-173">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e0056-174">String</span><span class="sxs-lookup"><span data-stu-id="e0056-174">String</span></span>|<span data-ttu-id="e0056-175">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="e0056-175">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="e0056-176">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e0056-176">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="e0056-177">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="e0056-177">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e0056-178">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e0056-178">Requirements</span></span>

|<span data-ttu-id="e0056-179">Requisito</span><span class="sxs-lookup"><span data-stu-id="e0056-179">Requirement</span></span>| <span data-ttu-id="e0056-180">Valor</span><span class="sxs-lookup"><span data-stu-id="e0056-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0056-181">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e0056-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e0056-182">1.3</span><span class="sxs-lookup"><span data-stu-id="e0056-182">1.3</span></span>|
|[<span data-ttu-id="e0056-183">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e0056-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e0056-184">Restrito</span><span class="sxs-lookup"><span data-stu-id="e0056-184">Restricted</span></span>|
|[<span data-ttu-id="e0056-185">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e0056-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e0056-186">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e0056-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e0056-187">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e0056-187">Returns:</span></span>

<span data-ttu-id="e0056-188">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="e0056-188">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e0056-189">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e0056-189">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-13"></a><span data-ttu-id="e0056-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="e0056-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="e0056-191">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="e0056-191">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="e0056-p105">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para as datas e horas. O Outlook em uma área de trabalho usa o fuso horário do computador cliente; o Outlook na Web usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="e0056-p105">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="e0056-p106">Se o aplicativo de email estiver sendo executado no Outlook em um cliente da área de trabalho, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador cliente. Se o aplicativo de email estiver sendo executado no Outlook na Web, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="e0056-p106">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e0056-197">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e0056-197">Parameters</span></span>

|<span data-ttu-id="e0056-198">Nome</span><span class="sxs-lookup"><span data-stu-id="e0056-198">Name</span></span>| <span data-ttu-id="e0056-199">Tipo</span><span class="sxs-lookup"><span data-stu-id="e0056-199">Type</span></span>| <span data-ttu-id="e0056-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0056-200">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="e0056-201">Date</span><span class="sxs-lookup"><span data-stu-id="e0056-201">Date</span></span>|<span data-ttu-id="e0056-202">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="e0056-202">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e0056-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e0056-203">Requirements</span></span>

|<span data-ttu-id="e0056-204">Requisito</span><span class="sxs-lookup"><span data-stu-id="e0056-204">Requirement</span></span>| <span data-ttu-id="e0056-205">Valor</span><span class="sxs-lookup"><span data-stu-id="e0056-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0056-206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e0056-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e0056-207">1.0</span><span class="sxs-lookup"><span data-stu-id="e0056-207">1.0</span></span>|
|[<span data-ttu-id="e0056-208">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e0056-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e0056-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e0056-209">ReadItem</span></span>|
|[<span data-ttu-id="e0056-210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e0056-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e0056-211">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e0056-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e0056-212">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e0056-212">Returns:</span></span>

<span data-ttu-id="e0056-213">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="e0056-213">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="e0056-214">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e0056-214">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e0056-215">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="e0056-215">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="e0056-216">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="e0056-216">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e0056-p107">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="e0056-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e0056-219">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e0056-219">Parameters</span></span>

|<span data-ttu-id="e0056-220">Nome</span><span class="sxs-lookup"><span data-stu-id="e0056-220">Name</span></span>| <span data-ttu-id="e0056-221">Tipo</span><span class="sxs-lookup"><span data-stu-id="e0056-221">Type</span></span>| <span data-ttu-id="e0056-222">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0056-222">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e0056-223">String</span><span class="sxs-lookup"><span data-stu-id="e0056-223">String</span></span>|<span data-ttu-id="e0056-224">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="e0056-224">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="e0056-225">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e0056-225">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="e0056-226">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="e0056-226">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e0056-227">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e0056-227">Requirements</span></span>

|<span data-ttu-id="e0056-228">Requisito</span><span class="sxs-lookup"><span data-stu-id="e0056-228">Requirement</span></span>| <span data-ttu-id="e0056-229">Valor</span><span class="sxs-lookup"><span data-stu-id="e0056-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0056-230">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e0056-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e0056-231">1.3</span><span class="sxs-lookup"><span data-stu-id="e0056-231">1.3</span></span>|
|[<span data-ttu-id="e0056-232">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e0056-232">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e0056-233">Restrito</span><span class="sxs-lookup"><span data-stu-id="e0056-233">Restricted</span></span>|
|[<span data-ttu-id="e0056-234">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e0056-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e0056-235">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e0056-235">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e0056-236">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e0056-236">Returns:</span></span>

<span data-ttu-id="e0056-237">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="e0056-237">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e0056-238">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e0056-238">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="e0056-239">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="e0056-239">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="e0056-240">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="e0056-240">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="e0056-241">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="e0056-241">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e0056-242">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e0056-242">Parameters</span></span>

|<span data-ttu-id="e0056-243">Nome</span><span class="sxs-lookup"><span data-stu-id="e0056-243">Name</span></span>| <span data-ttu-id="e0056-244">Tipo</span><span class="sxs-lookup"><span data-stu-id="e0056-244">Type</span></span>| <span data-ttu-id="e0056-245">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0056-245">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="e0056-246">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e0056-246">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)|<span data-ttu-id="e0056-247">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="e0056-247">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e0056-248">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e0056-248">Requirements</span></span>

|<span data-ttu-id="e0056-249">Requisito</span><span class="sxs-lookup"><span data-stu-id="e0056-249">Requirement</span></span>| <span data-ttu-id="e0056-250">Valor</span><span class="sxs-lookup"><span data-stu-id="e0056-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0056-251">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e0056-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e0056-252">1.0</span><span class="sxs-lookup"><span data-stu-id="e0056-252">1.0</span></span>|
|[<span data-ttu-id="e0056-253">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e0056-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e0056-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e0056-254">ReadItem</span></span>|
|[<span data-ttu-id="e0056-255">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e0056-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e0056-256">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e0056-256">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e0056-257">Retorna:</span><span class="sxs-lookup"><span data-stu-id="e0056-257">Returns:</span></span>

<span data-ttu-id="e0056-258">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="e0056-258">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="e0056-259">Tipo: Data</span><span class="sxs-lookup"><span data-stu-id="e0056-259">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="e0056-260">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e0056-260">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="e0056-261">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e0056-261">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="e0056-262">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="e0056-262">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e0056-263">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="e0056-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e0056-264">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="e0056-264">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e0056-p108">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir um instância da série. Isso ocorre porque no Outlook no Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="e0056-p108">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="e0056-267">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="e0056-267">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="e0056-268">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="e0056-268">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e0056-269">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e0056-269">Parameters</span></span>

|<span data-ttu-id="e0056-270">Nome</span><span class="sxs-lookup"><span data-stu-id="e0056-270">Name</span></span>| <span data-ttu-id="e0056-271">Tipo</span><span class="sxs-lookup"><span data-stu-id="e0056-271">Type</span></span>| <span data-ttu-id="e0056-272">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0056-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e0056-273">String</span><span class="sxs-lookup"><span data-stu-id="e0056-273">String</span></span>|<span data-ttu-id="e0056-274">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="e0056-274">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e0056-275">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e0056-275">Requirements</span></span>

|<span data-ttu-id="e0056-276">Requisito</span><span class="sxs-lookup"><span data-stu-id="e0056-276">Requirement</span></span>| <span data-ttu-id="e0056-277">Valor</span><span class="sxs-lookup"><span data-stu-id="e0056-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0056-278">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e0056-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e0056-279">1.0</span><span class="sxs-lookup"><span data-stu-id="e0056-279">1.0</span></span>|
|[<span data-ttu-id="e0056-280">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e0056-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e0056-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e0056-281">ReadItem</span></span>|
|[<span data-ttu-id="e0056-282">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e0056-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e0056-283">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e0056-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e0056-284">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e0056-284">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="e0056-285">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e0056-285">displayMessageForm(itemId)</span></span>

<span data-ttu-id="e0056-286">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="e0056-286">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="e0056-287">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="e0056-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e0056-288">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="e0056-288">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e0056-289">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e0056-289">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="e0056-290">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="e0056-290">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="e0056-p109">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="e0056-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e0056-293">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e0056-293">Parameters</span></span>

|<span data-ttu-id="e0056-294">Nome</span><span class="sxs-lookup"><span data-stu-id="e0056-294">Name</span></span>| <span data-ttu-id="e0056-295">Tipo</span><span class="sxs-lookup"><span data-stu-id="e0056-295">Type</span></span>| <span data-ttu-id="e0056-296">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0056-296">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e0056-297">String</span><span class="sxs-lookup"><span data-stu-id="e0056-297">String</span></span>|<span data-ttu-id="e0056-298">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="e0056-298">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e0056-299">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e0056-299">Requirements</span></span>

|<span data-ttu-id="e0056-300">Requisito</span><span class="sxs-lookup"><span data-stu-id="e0056-300">Requirement</span></span>| <span data-ttu-id="e0056-301">Valor</span><span class="sxs-lookup"><span data-stu-id="e0056-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0056-302">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e0056-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e0056-303">1.0</span><span class="sxs-lookup"><span data-stu-id="e0056-303">1.0</span></span>|
|[<span data-ttu-id="e0056-304">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e0056-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e0056-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e0056-305">ReadItem</span></span>|
|[<span data-ttu-id="e0056-306">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e0056-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e0056-307">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e0056-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e0056-308">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e0056-308">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="e0056-309">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="e0056-309">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="e0056-310">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="e0056-310">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e0056-311">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="e0056-311">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e0056-p110">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="e0056-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e0056-p111">No Outlook na Web e em dispositivos móveis, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="e0056-p111">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="e0056-p112">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="e0056-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="e0056-319">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="e0056-319">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e0056-320">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e0056-320">Parameters</span></span>

|<span data-ttu-id="e0056-321">Nome</span><span class="sxs-lookup"><span data-stu-id="e0056-321">Name</span></span>| <span data-ttu-id="e0056-322">Tipo</span><span class="sxs-lookup"><span data-stu-id="e0056-322">Type</span></span>| <span data-ttu-id="e0056-323">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0056-323">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e0056-324">Object</span><span class="sxs-lookup"><span data-stu-id="e0056-324">Object</span></span> | <span data-ttu-id="e0056-325">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="e0056-325">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="e0056-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="e0056-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="e0056-p113">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="e0056-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="e0056-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="e0056-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="e0056-p114">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="e0056-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="e0056-332">Data</span><span class="sxs-lookup"><span data-stu-id="e0056-332">Date</span></span> | <span data-ttu-id="e0056-333">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="e0056-333">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="e0056-334">Data</span><span class="sxs-lookup"><span data-stu-id="e0056-334">Date</span></span> | <span data-ttu-id="e0056-335">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="e0056-335">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="e0056-336">String</span><span class="sxs-lookup"><span data-stu-id="e0056-336">String</span></span> | <span data-ttu-id="e0056-p115">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e0056-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="e0056-339">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="e0056-339">Array.&lt;String&gt;</span></span> | <span data-ttu-id="e0056-p116">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="e0056-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e0056-342">String</span><span class="sxs-lookup"><span data-stu-id="e0056-342">String</span></span> | <span data-ttu-id="e0056-p117">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="e0056-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="e0056-345">String</span><span class="sxs-lookup"><span data-stu-id="e0056-345">String</span></span> | <span data-ttu-id="e0056-p118">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="e0056-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e0056-348">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e0056-348">Requirements</span></span>

|<span data-ttu-id="e0056-349">Requisito</span><span class="sxs-lookup"><span data-stu-id="e0056-349">Requirement</span></span>| <span data-ttu-id="e0056-350">Valor</span><span class="sxs-lookup"><span data-stu-id="e0056-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0056-351">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e0056-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e0056-352">1.0</span><span class="sxs-lookup"><span data-stu-id="e0056-352">1.0</span></span>|
|[<span data-ttu-id="e0056-353">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e0056-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e0056-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e0056-354">ReadItem</span></span>|
|[<span data-ttu-id="e0056-355">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e0056-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e0056-356">Read</span><span class="sxs-lookup"><span data-stu-id="e0056-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e0056-357">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e0056-357">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="e0056-358">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e0056-358">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e0056-359">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="e0056-359">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="e0056-p119">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="e0056-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="e0056-362">Você pode passar o token e um identificador de anexo ou identificador de item para um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="e0056-362">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="e0056-363">O sistema de terceiros usa o token como um token de autorização de portador para chamar a operação [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) dos serviços Web do Exchange (EWS) ou a operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) para retornar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="e0056-363">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="e0056-364">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="e0056-364">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e0056-365">Chamar o `getCallbackTokenAsync` método no modo de leitura requer um nível mínimo de permissão de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="e0056-365">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="e0056-366">A `getCallbackTokenAsync` chamada no modo de redação requer que você tenha salvo o item.</span><span class="sxs-lookup"><span data-stu-id="e0056-366">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="e0056-367">O [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) método requer um nível de permissão mínimo de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="e0056-367">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e0056-368">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e0056-368">Parameters</span></span>

|<span data-ttu-id="e0056-369">Nome</span><span class="sxs-lookup"><span data-stu-id="e0056-369">Name</span></span>| <span data-ttu-id="e0056-370">Tipo</span><span class="sxs-lookup"><span data-stu-id="e0056-370">Type</span></span>| <span data-ttu-id="e0056-371">Atributos</span><span class="sxs-lookup"><span data-stu-id="e0056-371">Attributes</span></span>| <span data-ttu-id="e0056-372">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0056-372">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e0056-373">function</span><span class="sxs-lookup"><span data-stu-id="e0056-373">function</span></span>||<span data-ttu-id="e0056-374">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e0056-374">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e0056-375">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e0056-375">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="e0056-376">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="e0056-376">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="e0056-377">Objeto</span><span class="sxs-lookup"><span data-stu-id="e0056-377">Object</span></span>| <span data-ttu-id="e0056-378">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e0056-378">&lt;optional&gt;</span></span>|<span data-ttu-id="e0056-379">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="e0056-379">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e0056-380">Erros</span><span class="sxs-lookup"><span data-stu-id="e0056-380">Errors</span></span>

|<span data-ttu-id="e0056-381">Código de erro</span><span class="sxs-lookup"><span data-stu-id="e0056-381">Error code</span></span>|<span data-ttu-id="e0056-382">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0056-382">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="e0056-383">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="e0056-383">The request has failed.</span></span> <span data-ttu-id="e0056-384">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="e0056-384">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="e0056-385">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="e0056-385">The Exchange server returned an error.</span></span> <span data-ttu-id="e0056-386">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="e0056-386">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="e0056-387">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="e0056-387">The user is no longer connected to the network.</span></span> <span data-ttu-id="e0056-388">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="e0056-388">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e0056-389">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e0056-389">Requirements</span></span>

|<span data-ttu-id="e0056-390">Requisito</span><span class="sxs-lookup"><span data-stu-id="e0056-390">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e0056-391">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e0056-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e0056-392">1.0</span><span class="sxs-lookup"><span data-stu-id="e0056-392">1.0</span></span> | <span data-ttu-id="e0056-393">1.3</span><span class="sxs-lookup"><span data-stu-id="e0056-393">1.3</span></span> |
|[<span data-ttu-id="e0056-394">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e0056-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e0056-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e0056-395">ReadItem</span></span> | <span data-ttu-id="e0056-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e0056-396">ReadItem</span></span> |
|[<span data-ttu-id="e0056-397">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e0056-397">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e0056-398">Read</span><span class="sxs-lookup"><span data-stu-id="e0056-398">Read</span></span> | <span data-ttu-id="e0056-399">Escrever</span><span class="sxs-lookup"><span data-stu-id="e0056-399">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="e0056-400">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e0056-400">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="e0056-401">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e0056-401">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e0056-402">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="e0056-402">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="e0056-403">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="e0056-403">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="e0056-404">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e0056-404">Parameters</span></span>

|<span data-ttu-id="e0056-405">Nome</span><span class="sxs-lookup"><span data-stu-id="e0056-405">Name</span></span>| <span data-ttu-id="e0056-406">Tipo</span><span class="sxs-lookup"><span data-stu-id="e0056-406">Type</span></span>| <span data-ttu-id="e0056-407">Atributos</span><span class="sxs-lookup"><span data-stu-id="e0056-407">Attributes</span></span>| <span data-ttu-id="e0056-408">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0056-408">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e0056-409">function</span><span class="sxs-lookup"><span data-stu-id="e0056-409">function</span></span>||<span data-ttu-id="e0056-410">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e0056-410">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e0056-411">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e0056-411">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="e0056-412">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="e0056-412">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="e0056-413">Objeto</span><span class="sxs-lookup"><span data-stu-id="e0056-413">Object</span></span>| <span data-ttu-id="e0056-414">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e0056-414">&lt;optional&gt;</span></span>|<span data-ttu-id="e0056-415">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="e0056-415">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e0056-416">Erros</span><span class="sxs-lookup"><span data-stu-id="e0056-416">Errors</span></span>

|<span data-ttu-id="e0056-417">Código de erro</span><span class="sxs-lookup"><span data-stu-id="e0056-417">Error code</span></span>|<span data-ttu-id="e0056-418">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0056-418">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="e0056-419">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="e0056-419">The request has failed.</span></span> <span data-ttu-id="e0056-420">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="e0056-420">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="e0056-421">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="e0056-421">The Exchange server returned an error.</span></span> <span data-ttu-id="e0056-422">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="e0056-422">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="e0056-423">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="e0056-423">The user is no longer connected to the network.</span></span> <span data-ttu-id="e0056-424">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="e0056-424">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e0056-425">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e0056-425">Requirements</span></span>

|<span data-ttu-id="e0056-426">Requisito</span><span class="sxs-lookup"><span data-stu-id="e0056-426">Requirement</span></span>| <span data-ttu-id="e0056-427">Valor</span><span class="sxs-lookup"><span data-stu-id="e0056-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0056-428">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e0056-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e0056-429">1.0</span><span class="sxs-lookup"><span data-stu-id="e0056-429">1.0</span></span>|
|[<span data-ttu-id="e0056-430">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e0056-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e0056-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e0056-431">ReadItem</span></span>|
|[<span data-ttu-id="e0056-432">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="e0056-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e0056-433">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e0056-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e0056-434">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e0056-434">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="e0056-435">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e0056-435">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="e0056-436">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="e0056-436">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="e0056-437">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="e0056-437">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="e0056-438">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="e0056-438">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="e0056-439">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="e0056-439">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="e0056-440">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="e0056-440">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="e0056-441">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="e0056-441">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="e0056-442">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="e0056-442">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="e0056-443">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="e0056-443">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="e0056-444">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="e0056-444">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="e0056-p129">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="e0056-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="e0056-447">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="e0056-447">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="e0056-448">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="e0056-448">Version differences</span></span>

<span data-ttu-id="e0056-449">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="e0056-449">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="e0056-p130">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="e0056-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e0056-453">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="e0056-453">Parameters</span></span>

|<span data-ttu-id="e0056-454">Nome</span><span class="sxs-lookup"><span data-stu-id="e0056-454">Name</span></span>| <span data-ttu-id="e0056-455">Tipo</span><span class="sxs-lookup"><span data-stu-id="e0056-455">Type</span></span>| <span data-ttu-id="e0056-456">Atributos</span><span class="sxs-lookup"><span data-stu-id="e0056-456">Attributes</span></span>| <span data-ttu-id="e0056-457">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0056-457">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="e0056-458">String</span><span class="sxs-lookup"><span data-stu-id="e0056-458">String</span></span>||<span data-ttu-id="e0056-459">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="e0056-459">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="e0056-460">function</span><span class="sxs-lookup"><span data-stu-id="e0056-460">function</span></span>||<span data-ttu-id="e0056-461">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e0056-461">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e0056-462">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e0056-462">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="e0056-463">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="e0056-463">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="e0056-464">Objeto</span><span class="sxs-lookup"><span data-stu-id="e0056-464">Object</span></span>| <span data-ttu-id="e0056-465">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="e0056-465">&lt;optional&gt;</span></span>|<span data-ttu-id="e0056-466">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="e0056-466">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e0056-467">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e0056-467">Requirements</span></span>

|<span data-ttu-id="e0056-468">Requisito</span><span class="sxs-lookup"><span data-stu-id="e0056-468">Requirement</span></span>| <span data-ttu-id="e0056-469">Valor</span><span class="sxs-lookup"><span data-stu-id="e0056-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="e0056-470">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e0056-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e0056-471">1.0</span><span class="sxs-lookup"><span data-stu-id="e0056-471">1.0</span></span>|
|[<span data-ttu-id="e0056-472">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e0056-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e0056-473">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="e0056-473">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="e0056-474">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e0056-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e0056-475">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e0056-475">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e0056-476">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e0056-476">Example</span></span>

<span data-ttu-id="e0056-477">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="e0056-477">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
