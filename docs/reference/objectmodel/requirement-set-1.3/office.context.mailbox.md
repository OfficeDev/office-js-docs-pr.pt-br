---
title: Office. Context. Mailbox – conjunto de requisitos 1,3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: acc304e302e3bb4d912ecbafee51cc35c88c091c
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268667"
---
# <a name="mailbox"></a><span data-ttu-id="afb52-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="afb52-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="afb52-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="afb52-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="afb52-104">Fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="afb52-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="afb52-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afb52-105">Requirements</span></span>

|<span data-ttu-id="afb52-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="afb52-106">Requirement</span></span>| <span data-ttu-id="afb52-107">Valor</span><span class="sxs-lookup"><span data-stu-id="afb52-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="afb52-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afb52-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afb52-109">1.0</span><span class="sxs-lookup"><span data-stu-id="afb52-109">1.0</span></span>|
|[<span data-ttu-id="afb52-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="afb52-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="afb52-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="afb52-111">Restricted</span></span>|
|[<span data-ttu-id="afb52-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afb52-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="afb52-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="afb52-113">Compose or Read</span></span>|

<span data-ttu-id="afb52-114">| [ewsUrl](#ewsurl-string) | Membro | | [convertToEwsId](#converttoewsiditemid-restversion--string) | Método | | [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | Método | | [convertToRestId](#converttorestiditemid-restversion--string) | Método | | [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Método | | [displayAppointmentForm](#displayappointmentformitemid) | Método | | [displayMessageForm](#displaymessageformitemid) | Método | | [displayNewAppointmentForm](#displaynewappointmentformparameters) | Método | | [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Método | | [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Método | | [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Método |</span><span class="sxs-lookup"><span data-stu-id="afb52-114">| [ewsUrl](#ewsurl-string) | Member | | [convertToEwsId](#converttoewsiditemid-restversion--string) | Method | | [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | Method | | [convertToRestId](#converttorestiditemid-restversion--string) | Method | | [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Method | | [displayAppointmentForm](#displayappointmentformitemid) | Method | | [displayMessageForm](#displaymessageformitemid) | Method | | [displayNewAppointmentForm](#displaynewappointmentformparameters) | Method | | [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Method | | [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Method | | [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Method |</span></span>

### <a name="namespaces"></a><span data-ttu-id="afb52-115">Namespaces</span><span class="sxs-lookup"><span data-stu-id="afb52-115">Namespaces</span></span>

<span data-ttu-id="afb52-116">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="afb52-116">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="afb52-117">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="afb52-117">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="afb52-118">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="afb52-118">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="afb52-119">Membros</span><span class="sxs-lookup"><span data-stu-id="afb52-119">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="afb52-120">ewsUrl: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="afb52-120">ewsUrl: String</span></span>

<span data-ttu-id="afb52-121">Obtém a URL do ponto de extremidade dos Serviços Web do Exchange (EWS) para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="afb52-121">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="afb52-122">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="afb52-122">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="afb52-123">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="afb52-123">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="afb52-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="afb52-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="afb52-126">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="afb52-126">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="afb52-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="afb52-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="afb52-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="afb52-129">Type</span></span>

*   <span data-ttu-id="afb52-130">String</span><span class="sxs-lookup"><span data-stu-id="afb52-130">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="afb52-131">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afb52-131">Requirements</span></span>

|<span data-ttu-id="afb52-132">Requisito</span><span class="sxs-lookup"><span data-stu-id="afb52-132">Requirement</span></span>| <span data-ttu-id="afb52-133">Valor</span><span class="sxs-lookup"><span data-stu-id="afb52-133">Value</span></span>|
|---|---|
|[<span data-ttu-id="afb52-134">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afb52-134">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afb52-135">1.0</span><span class="sxs-lookup"><span data-stu-id="afb52-135">1.0</span></span>|
|[<span data-ttu-id="afb52-136">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="afb52-136">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="afb52-137">ReadItem</span><span class="sxs-lookup"><span data-stu-id="afb52-137">ReadItem</span></span>|
|[<span data-ttu-id="afb52-138">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afb52-138">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="afb52-139">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="afb52-139">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="afb52-140">Métodos</span><span class="sxs-lookup"><span data-stu-id="afb52-140">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="afb52-141">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="afb52-141">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="afb52-142">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="afb52-142">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="afb52-143">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="afb52-143">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="afb52-p104">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="afb52-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="afb52-146">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="afb52-146">Parameters</span></span>

|<span data-ttu-id="afb52-147">Nome</span><span class="sxs-lookup"><span data-stu-id="afb52-147">Name</span></span>| <span data-ttu-id="afb52-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="afb52-148">Type</span></span>| <span data-ttu-id="afb52-149">Descrição</span><span class="sxs-lookup"><span data-stu-id="afb52-149">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="afb52-150">String</span><span class="sxs-lookup"><span data-stu-id="afb52-150">String</span></span>|<span data-ttu-id="afb52-151">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="afb52-151">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="afb52-152">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="afb52-152">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="afb52-153">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="afb52-153">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="afb52-154">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afb52-154">Requirements</span></span>

|<span data-ttu-id="afb52-155">Requisito</span><span class="sxs-lookup"><span data-stu-id="afb52-155">Requirement</span></span>| <span data-ttu-id="afb52-156">Valor</span><span class="sxs-lookup"><span data-stu-id="afb52-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="afb52-157">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afb52-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afb52-158">1.3</span><span class="sxs-lookup"><span data-stu-id="afb52-158">1.3</span></span>|
|[<span data-ttu-id="afb52-159">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="afb52-159">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="afb52-160">Restrito</span><span class="sxs-lookup"><span data-stu-id="afb52-160">Restricted</span></span>|
|[<span data-ttu-id="afb52-161">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afb52-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="afb52-162">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="afb52-162">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="afb52-163">Retorna:</span><span class="sxs-lookup"><span data-stu-id="afb52-163">Returns:</span></span>

<span data-ttu-id="afb52-164">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="afb52-164">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="afb52-165">Exemplo</span><span class="sxs-lookup"><span data-stu-id="afb52-165">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-13"></a><span data-ttu-id="afb52-166">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="afb52-166">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="afb52-167">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="afb52-167">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="afb52-168">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para datas e horas.</span><span class="sxs-lookup"><span data-stu-id="afb52-168">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="afb52-169">O Outlook em uma área de trabalho usa o fuso horário do computador cliente; O Outlook na Web usa o fuso horário definido no centro de administração do Exchange (Eat).</span><span class="sxs-lookup"><span data-stu-id="afb52-169">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="afb52-170">Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="afb52-170">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="afb52-171">Se o aplicativo de email estiver em execução no Outlook em um cliente desktop `convertToLocalClientTime` , o método retornará um objeto Dictionary com os valores definidos para o fuso horário do computador cliente.</span><span class="sxs-lookup"><span data-stu-id="afb52-171">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="afb52-172">Se o aplicativo de email estiver em execução no Outlook na Web, `convertToLocalClientTime` o método retornará um objeto Dictionary com os valores definidos para o fuso horário especificado no Eat.</span><span class="sxs-lookup"><span data-stu-id="afb52-172">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="afb52-173">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="afb52-173">Parameters</span></span>

|<span data-ttu-id="afb52-174">Nome</span><span class="sxs-lookup"><span data-stu-id="afb52-174">Name</span></span>| <span data-ttu-id="afb52-175">Tipo</span><span class="sxs-lookup"><span data-stu-id="afb52-175">Type</span></span>| <span data-ttu-id="afb52-176">Descrição</span><span class="sxs-lookup"><span data-stu-id="afb52-176">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="afb52-177">Date</span><span class="sxs-lookup"><span data-stu-id="afb52-177">Date</span></span>|<span data-ttu-id="afb52-178">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="afb52-178">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="afb52-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afb52-179">Requirements</span></span>

|<span data-ttu-id="afb52-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="afb52-180">Requirement</span></span>| <span data-ttu-id="afb52-181">Valor</span><span class="sxs-lookup"><span data-stu-id="afb52-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="afb52-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afb52-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afb52-183">1.0</span><span class="sxs-lookup"><span data-stu-id="afb52-183">1.0</span></span>|
|[<span data-ttu-id="afb52-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="afb52-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="afb52-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="afb52-185">ReadItem</span></span>|
|[<span data-ttu-id="afb52-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afb52-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="afb52-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="afb52-187">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="afb52-188">Retorna:</span><span class="sxs-lookup"><span data-stu-id="afb52-188">Returns:</span></span>

<span data-ttu-id="afb52-189">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="afb52-189">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="afb52-190">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="afb52-190">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="afb52-191">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="afb52-191">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="afb52-192">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="afb52-192">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="afb52-p107">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="afb52-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="afb52-195">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="afb52-195">Parameters</span></span>

|<span data-ttu-id="afb52-196">Nome</span><span class="sxs-lookup"><span data-stu-id="afb52-196">Name</span></span>| <span data-ttu-id="afb52-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="afb52-197">Type</span></span>| <span data-ttu-id="afb52-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="afb52-198">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="afb52-199">String</span><span class="sxs-lookup"><span data-stu-id="afb52-199">String</span></span>|<span data-ttu-id="afb52-200">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="afb52-200">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="afb52-201">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="afb52-201">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="afb52-202">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="afb52-202">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="afb52-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afb52-203">Requirements</span></span>

|<span data-ttu-id="afb52-204">Requisito</span><span class="sxs-lookup"><span data-stu-id="afb52-204">Requirement</span></span>| <span data-ttu-id="afb52-205">Valor</span><span class="sxs-lookup"><span data-stu-id="afb52-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="afb52-206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afb52-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afb52-207">1.3</span><span class="sxs-lookup"><span data-stu-id="afb52-207">1.3</span></span>|
|[<span data-ttu-id="afb52-208">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="afb52-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="afb52-209">Restrito</span><span class="sxs-lookup"><span data-stu-id="afb52-209">Restricted</span></span>|
|[<span data-ttu-id="afb52-210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afb52-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="afb52-211">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="afb52-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="afb52-212">Retorna:</span><span class="sxs-lookup"><span data-stu-id="afb52-212">Returns:</span></span>

<span data-ttu-id="afb52-213">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="afb52-213">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="afb52-214">Exemplo</span><span class="sxs-lookup"><span data-stu-id="afb52-214">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="afb52-215">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="afb52-215">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="afb52-216">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="afb52-216">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="afb52-217">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="afb52-217">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="afb52-218">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="afb52-218">Parameters</span></span>

|<span data-ttu-id="afb52-219">Nome</span><span class="sxs-lookup"><span data-stu-id="afb52-219">Name</span></span>| <span data-ttu-id="afb52-220">Tipo</span><span class="sxs-lookup"><span data-stu-id="afb52-220">Type</span></span>| <span data-ttu-id="afb52-221">Descrição</span><span class="sxs-lookup"><span data-stu-id="afb52-221">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="afb52-222">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="afb52-222">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)|<span data-ttu-id="afb52-223">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="afb52-223">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="afb52-224">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afb52-224">Requirements</span></span>

|<span data-ttu-id="afb52-225">Requisito</span><span class="sxs-lookup"><span data-stu-id="afb52-225">Requirement</span></span>| <span data-ttu-id="afb52-226">Valor</span><span class="sxs-lookup"><span data-stu-id="afb52-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="afb52-227">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afb52-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afb52-228">1.0</span><span class="sxs-lookup"><span data-stu-id="afb52-228">1.0</span></span>|
|[<span data-ttu-id="afb52-229">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="afb52-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="afb52-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="afb52-230">ReadItem</span></span>|
|[<span data-ttu-id="afb52-231">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="afb52-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="afb52-232">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="afb52-232">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="afb52-233">Retorna:</span><span class="sxs-lookup"><span data-stu-id="afb52-233">Returns:</span></span>

<span data-ttu-id="afb52-234">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="afb52-234">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="afb52-235">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="afb52-235">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="afb52-236">Date</span><span class="sxs-lookup"><span data-stu-id="afb52-236">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="afb52-237">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="afb52-237">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="afb52-238">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="afb52-238">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="afb52-239">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="afb52-239">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="afb52-240">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="afb52-240">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="afb52-241">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente ou o compromisso mestre de uma série recorrente, mas não é possível exibir uma instância da série.</span><span class="sxs-lookup"><span data-stu-id="afb52-241">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="afb52-242">Isso ocorre porque, no Outlook no Mac, você não pode acessar as propriedades (incluindo a ID do item) de instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="afb52-242">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="afb52-243">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres 32 KB.</span><span class="sxs-lookup"><span data-stu-id="afb52-243">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="afb52-244">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="afb52-244">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="afb52-245">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="afb52-245">Parameters</span></span>

|<span data-ttu-id="afb52-246">Nome</span><span class="sxs-lookup"><span data-stu-id="afb52-246">Name</span></span>| <span data-ttu-id="afb52-247">Tipo</span><span class="sxs-lookup"><span data-stu-id="afb52-247">Type</span></span>| <span data-ttu-id="afb52-248">Descrição</span><span class="sxs-lookup"><span data-stu-id="afb52-248">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="afb52-249">String</span><span class="sxs-lookup"><span data-stu-id="afb52-249">String</span></span>|<span data-ttu-id="afb52-250">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="afb52-250">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="afb52-251">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afb52-251">Requirements</span></span>

|<span data-ttu-id="afb52-252">Requisito</span><span class="sxs-lookup"><span data-stu-id="afb52-252">Requirement</span></span>| <span data-ttu-id="afb52-253">Valor</span><span class="sxs-lookup"><span data-stu-id="afb52-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="afb52-254">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afb52-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afb52-255">1.0</span><span class="sxs-lookup"><span data-stu-id="afb52-255">1.0</span></span>|
|[<span data-ttu-id="afb52-256">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="afb52-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="afb52-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="afb52-257">ReadItem</span></span>|
|[<span data-ttu-id="afb52-258">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="afb52-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="afb52-259">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="afb52-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="afb52-260">Exemplo</span><span class="sxs-lookup"><span data-stu-id="afb52-260">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="afb52-261">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="afb52-261">displayMessageForm(itemId)</span></span>

<span data-ttu-id="afb52-262">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="afb52-262">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="afb52-263">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="afb52-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="afb52-264">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="afb52-264">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="afb52-265">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="afb52-265">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="afb52-266">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="afb52-266">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="afb52-p109">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="afb52-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="afb52-269">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="afb52-269">Parameters</span></span>

|<span data-ttu-id="afb52-270">Nome</span><span class="sxs-lookup"><span data-stu-id="afb52-270">Name</span></span>| <span data-ttu-id="afb52-271">Tipo</span><span class="sxs-lookup"><span data-stu-id="afb52-271">Type</span></span>| <span data-ttu-id="afb52-272">Descrição</span><span class="sxs-lookup"><span data-stu-id="afb52-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="afb52-273">String</span><span class="sxs-lookup"><span data-stu-id="afb52-273">String</span></span>|<span data-ttu-id="afb52-274">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="afb52-274">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="afb52-275">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afb52-275">Requirements</span></span>

|<span data-ttu-id="afb52-276">Requisito</span><span class="sxs-lookup"><span data-stu-id="afb52-276">Requirement</span></span>| <span data-ttu-id="afb52-277">Valor</span><span class="sxs-lookup"><span data-stu-id="afb52-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="afb52-278">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afb52-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afb52-279">1.0</span><span class="sxs-lookup"><span data-stu-id="afb52-279">1.0</span></span>|
|[<span data-ttu-id="afb52-280">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="afb52-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="afb52-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="afb52-281">ReadItem</span></span>|
|[<span data-ttu-id="afb52-282">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="afb52-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="afb52-283">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="afb52-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="afb52-284">Exemplo</span><span class="sxs-lookup"><span data-stu-id="afb52-284">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="afb52-285">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="afb52-285">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="afb52-286">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="afb52-286">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="afb52-287">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="afb52-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="afb52-p110">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="afb52-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="afb52-290">No Outlook na Web e dispositivos móveis, este método sempre exibe um formulário com um campo participantes.</span><span class="sxs-lookup"><span data-stu-id="afb52-290">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="afb52-291">Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="afb52-291">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="afb52-292">Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="afb52-292">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="afb52-p112">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="afb52-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="afb52-295">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="afb52-295">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="afb52-296">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="afb52-296">Parameters</span></span>

|<span data-ttu-id="afb52-297">Nome</span><span class="sxs-lookup"><span data-stu-id="afb52-297">Name</span></span>| <span data-ttu-id="afb52-298">Tipo</span><span class="sxs-lookup"><span data-stu-id="afb52-298">Type</span></span>| <span data-ttu-id="afb52-299">Descrição</span><span class="sxs-lookup"><span data-stu-id="afb52-299">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="afb52-300">Object</span><span class="sxs-lookup"><span data-stu-id="afb52-300">Object</span></span> | <span data-ttu-id="afb52-301">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="afb52-301">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="afb52-302">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="afb52-302">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="afb52-p113">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="afb52-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="afb52-305">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="afb52-305">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="afb52-p114">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="afb52-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="afb52-308">Data</span><span class="sxs-lookup"><span data-stu-id="afb52-308">Date</span></span> | <span data-ttu-id="afb52-309">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="afb52-309">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="afb52-310">Data</span><span class="sxs-lookup"><span data-stu-id="afb52-310">Date</span></span> | <span data-ttu-id="afb52-311">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="afb52-311">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="afb52-312">String</span><span class="sxs-lookup"><span data-stu-id="afb52-312">String</span></span> | <span data-ttu-id="afb52-p115">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="afb52-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="afb52-315">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="afb52-315">Array.&lt;String&gt;</span></span> | <span data-ttu-id="afb52-p116">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="afb52-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="afb52-318">String</span><span class="sxs-lookup"><span data-stu-id="afb52-318">String</span></span> | <span data-ttu-id="afb52-p117">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="afb52-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="afb52-321">String</span><span class="sxs-lookup"><span data-stu-id="afb52-321">String</span></span> | <span data-ttu-id="afb52-p118">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="afb52-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="afb52-324">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afb52-324">Requirements</span></span>

|<span data-ttu-id="afb52-325">Requisito</span><span class="sxs-lookup"><span data-stu-id="afb52-325">Requirement</span></span>| <span data-ttu-id="afb52-326">Valor</span><span class="sxs-lookup"><span data-stu-id="afb52-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="afb52-327">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afb52-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afb52-328">1.0</span><span class="sxs-lookup"><span data-stu-id="afb52-328">1.0</span></span>|
|[<span data-ttu-id="afb52-329">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="afb52-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="afb52-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="afb52-330">ReadItem</span></span>|
|[<span data-ttu-id="afb52-331">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afb52-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="afb52-332">Read</span><span class="sxs-lookup"><span data-stu-id="afb52-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="afb52-333">Exemplo</span><span class="sxs-lookup"><span data-stu-id="afb52-333">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="afb52-334">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="afb52-334">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="afb52-335">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="afb52-335">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="afb52-p119">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="afb52-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="afb52-p120">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="afb52-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="afb52-341">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="afb52-341">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="afb52-p121">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="afb52-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="afb52-344">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="afb52-344">Parameters</span></span>

|<span data-ttu-id="afb52-345">Nome</span><span class="sxs-lookup"><span data-stu-id="afb52-345">Name</span></span>| <span data-ttu-id="afb52-346">Tipo</span><span class="sxs-lookup"><span data-stu-id="afb52-346">Type</span></span>| <span data-ttu-id="afb52-347">Atributos</span><span class="sxs-lookup"><span data-stu-id="afb52-347">Attributes</span></span>| <span data-ttu-id="afb52-348">Descrição</span><span class="sxs-lookup"><span data-stu-id="afb52-348">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="afb52-349">function</span><span class="sxs-lookup"><span data-stu-id="afb52-349">function</span></span>||<span data-ttu-id="afb52-350">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="afb52-350">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="afb52-351">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="afb52-351">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="afb52-352">Se houvesse um erro, as `asyncResult.error` propriedades `asyncResult.diagnostics` e podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="afb52-352">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="afb52-353">Objeto</span><span class="sxs-lookup"><span data-stu-id="afb52-353">Object</span></span>| <span data-ttu-id="afb52-354">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="afb52-354">&lt;optional&gt;</span></span>|<span data-ttu-id="afb52-355">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="afb52-355">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="afb52-356">Erros</span><span class="sxs-lookup"><span data-stu-id="afb52-356">Errors</span></span>

|<span data-ttu-id="afb52-357">Código de erro</span><span class="sxs-lookup"><span data-stu-id="afb52-357">Error code</span></span>|<span data-ttu-id="afb52-358">Descrição</span><span class="sxs-lookup"><span data-stu-id="afb52-358">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="afb52-359">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="afb52-359">The request has failed.</span></span> <span data-ttu-id="afb52-360">Confira o objeto Diagnostics do código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="afb52-360">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="afb52-361">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="afb52-361">The Exchange server returned an error.</span></span> <span data-ttu-id="afb52-362">Confira o objeto Diagnostics para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="afb52-362">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="afb52-363">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="afb52-363">The user is no longer connected to the network.</span></span> <span data-ttu-id="afb52-364">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="afb52-364">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="afb52-365">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afb52-365">Requirements</span></span>

|<span data-ttu-id="afb52-366">Requisito</span><span class="sxs-lookup"><span data-stu-id="afb52-366">Requirement</span></span>| <span data-ttu-id="afb52-367">Valor</span><span class="sxs-lookup"><span data-stu-id="afb52-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="afb52-368">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afb52-368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afb52-369">1.0</span><span class="sxs-lookup"><span data-stu-id="afb52-369">1.0</span></span>|
|[<span data-ttu-id="afb52-370">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="afb52-370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="afb52-371">ReadItem</span><span class="sxs-lookup"><span data-stu-id="afb52-371">ReadItem</span></span>|
|[<span data-ttu-id="afb52-372">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afb52-372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="afb52-373">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="afb52-373">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="afb52-374">Exemplo</span><span class="sxs-lookup"><span data-stu-id="afb52-374">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="afb52-375">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="afb52-375">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="afb52-376">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="afb52-376">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="afb52-377">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="afb52-377">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="afb52-378">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="afb52-378">Parameters</span></span>

|<span data-ttu-id="afb52-379">Nome</span><span class="sxs-lookup"><span data-stu-id="afb52-379">Name</span></span>| <span data-ttu-id="afb52-380">Tipo</span><span class="sxs-lookup"><span data-stu-id="afb52-380">Type</span></span>| <span data-ttu-id="afb52-381">Atributos</span><span class="sxs-lookup"><span data-stu-id="afb52-381">Attributes</span></span>| <span data-ttu-id="afb52-382">Descrição</span><span class="sxs-lookup"><span data-stu-id="afb52-382">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="afb52-383">function</span><span class="sxs-lookup"><span data-stu-id="afb52-383">function</span></span>||<span data-ttu-id="afb52-384">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="afb52-384">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="afb52-385">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="afb52-385">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="afb52-386">Se houvesse um erro, as `asyncResult.error` propriedades `asyncResult.diagnostics` e podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="afb52-386">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="afb52-387">Objeto</span><span class="sxs-lookup"><span data-stu-id="afb52-387">Object</span></span>| <span data-ttu-id="afb52-388">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="afb52-388">&lt;optional&gt;</span></span>|<span data-ttu-id="afb52-389">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="afb52-389">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="afb52-390">Erros</span><span class="sxs-lookup"><span data-stu-id="afb52-390">Errors</span></span>

|<span data-ttu-id="afb52-391">Código de erro</span><span class="sxs-lookup"><span data-stu-id="afb52-391">Error code</span></span>|<span data-ttu-id="afb52-392">Descrição</span><span class="sxs-lookup"><span data-stu-id="afb52-392">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="afb52-393">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="afb52-393">The request has failed.</span></span> <span data-ttu-id="afb52-394">Confira o objeto Diagnostics do código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="afb52-394">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="afb52-395">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="afb52-395">The Exchange server returned an error.</span></span> <span data-ttu-id="afb52-396">Confira o objeto Diagnostics para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="afb52-396">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="afb52-397">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="afb52-397">The user is no longer connected to the network.</span></span> <span data-ttu-id="afb52-398">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="afb52-398">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="afb52-399">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afb52-399">Requirements</span></span>

|<span data-ttu-id="afb52-400">Requisito</span><span class="sxs-lookup"><span data-stu-id="afb52-400">Requirement</span></span>| <span data-ttu-id="afb52-401">Valor</span><span class="sxs-lookup"><span data-stu-id="afb52-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="afb52-402">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afb52-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afb52-403">1.0</span><span class="sxs-lookup"><span data-stu-id="afb52-403">1.0</span></span>|
|[<span data-ttu-id="afb52-404">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="afb52-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="afb52-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="afb52-405">ReadItem</span></span>|
|[<span data-ttu-id="afb52-406">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="afb52-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="afb52-407">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="afb52-407">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="afb52-408">Exemplo</span><span class="sxs-lookup"><span data-stu-id="afb52-408">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="afb52-409">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="afb52-409">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="afb52-410">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="afb52-410">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="afb52-411">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="afb52-411">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="afb52-412">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="afb52-412">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="afb52-413">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="afb52-413">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="afb52-414">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="afb52-414">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="afb52-415">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="afb52-415">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="afb52-416">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="afb52-416">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="afb52-417">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="afb52-417">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="afb52-418">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="afb52-418">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="afb52-p129">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="afb52-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="afb52-421">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="afb52-421">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="afb52-422">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="afb52-422">Version differences</span></span>

<span data-ttu-id="afb52-423">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="afb52-423">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="afb52-p130">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="afb52-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="afb52-427">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="afb52-427">Parameters</span></span>

|<span data-ttu-id="afb52-428">Nome</span><span class="sxs-lookup"><span data-stu-id="afb52-428">Name</span></span>| <span data-ttu-id="afb52-429">Tipo</span><span class="sxs-lookup"><span data-stu-id="afb52-429">Type</span></span>| <span data-ttu-id="afb52-430">Atributos</span><span class="sxs-lookup"><span data-stu-id="afb52-430">Attributes</span></span>| <span data-ttu-id="afb52-431">Descrição</span><span class="sxs-lookup"><span data-stu-id="afb52-431">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="afb52-432">String</span><span class="sxs-lookup"><span data-stu-id="afb52-432">String</span></span>||<span data-ttu-id="afb52-433">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="afb52-433">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="afb52-434">function</span><span class="sxs-lookup"><span data-stu-id="afb52-434">function</span></span>||<span data-ttu-id="afb52-435">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="afb52-435">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="afb52-436">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="afb52-436">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="afb52-437">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="afb52-437">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="afb52-438">Objeto</span><span class="sxs-lookup"><span data-stu-id="afb52-438">Object</span></span>| <span data-ttu-id="afb52-439">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="afb52-439">&lt;optional&gt;</span></span>|<span data-ttu-id="afb52-440">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="afb52-440">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="afb52-441">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afb52-441">Requirements</span></span>

|<span data-ttu-id="afb52-442">Requisito</span><span class="sxs-lookup"><span data-stu-id="afb52-442">Requirement</span></span>| <span data-ttu-id="afb52-443">Valor</span><span class="sxs-lookup"><span data-stu-id="afb52-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="afb52-444">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afb52-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afb52-445">1.0</span><span class="sxs-lookup"><span data-stu-id="afb52-445">1.0</span></span>|
|[<span data-ttu-id="afb52-446">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="afb52-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="afb52-447">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="afb52-447">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="afb52-448">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afb52-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="afb52-449">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="afb52-449">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="afb52-450">Exemplo</span><span class="sxs-lookup"><span data-stu-id="afb52-450">Example</span></span>

<span data-ttu-id="afb52-451">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="afb52-451">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
