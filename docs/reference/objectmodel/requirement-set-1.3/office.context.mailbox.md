---
title: Office. Context. Mailbox – conjunto de requisitos 1,3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 18e144702689c1df2b89f53c666932a74bc795ef
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450328"
---
# <a name="mailbox"></a><span data-ttu-id="778cc-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="778cc-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="778cc-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="778cc-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="778cc-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="778cc-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="778cc-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="778cc-105">Requirements</span></span>

|<span data-ttu-id="778cc-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="778cc-106">Requirement</span></span>| <span data-ttu-id="778cc-107">Valor</span><span class="sxs-lookup"><span data-stu-id="778cc-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="778cc-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="778cc-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778cc-109">1.0</span><span class="sxs-lookup"><span data-stu-id="778cc-109">1.0</span></span>|
|[<span data-ttu-id="778cc-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="778cc-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778cc-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="778cc-111">Restricted</span></span>|
|[<span data-ttu-id="778cc-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="778cc-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778cc-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="778cc-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="778cc-114">Namespaces</span><span class="sxs-lookup"><span data-stu-id="778cc-114">Namespaces</span></span>

<span data-ttu-id="778cc-115">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="778cc-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="778cc-116">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="778cc-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="778cc-117">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="778cc-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="778cc-118">Membros</span><span class="sxs-lookup"><span data-stu-id="778cc-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="778cc-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="778cc-119">ewsUrl :String</span></span>

<span data-ttu-id="778cc-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="778cc-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="778cc-122">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="778cc-122">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="778cc-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="778cc-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="778cc-125">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o membro `ewsUrl` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="778cc-125">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="778cc-p103">No modo de composição, é preciso chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) antes de poder usar o membro `ewsUrl`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="778cc-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="778cc-128">Tipo</span><span class="sxs-lookup"><span data-stu-id="778cc-128">Type</span></span>

*   <span data-ttu-id="778cc-129">String</span><span class="sxs-lookup"><span data-stu-id="778cc-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="778cc-130">Requisitos</span><span class="sxs-lookup"><span data-stu-id="778cc-130">Requirements</span></span>

|<span data-ttu-id="778cc-131">Requisito</span><span class="sxs-lookup"><span data-stu-id="778cc-131">Requirement</span></span>| <span data-ttu-id="778cc-132">Valor</span><span class="sxs-lookup"><span data-stu-id="778cc-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="778cc-133">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="778cc-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778cc-134">1.0</span><span class="sxs-lookup"><span data-stu-id="778cc-134">1.0</span></span>|
|[<span data-ttu-id="778cc-135">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="778cc-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778cc-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778cc-136">ReadItem</span></span>|
|[<span data-ttu-id="778cc-137">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="778cc-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778cc-138">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="778cc-138">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="778cc-139">Métodos</span><span class="sxs-lookup"><span data-stu-id="778cc-139">Methods</span></span>

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="778cc-140">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="778cc-140">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="778cc-141">Converte uma ID de item formatada para REST no formato EWS.</span><span class="sxs-lookup"><span data-stu-id="778cc-141">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="778cc-142">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="778cc-142">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="778cc-p104">IDs de itens recuperadas por meio de uma API REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)) usam um formato diferente daquele usado pelos Serviços Web do Exchange (EWS). O método `convertToEwsId` converte uma ID formatada como REST para o formato adequado para EWS.</span><span class="sxs-lookup"><span data-stu-id="778cc-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778cc-145">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="778cc-145">Parameters</span></span>

|<span data-ttu-id="778cc-146">Nome</span><span class="sxs-lookup"><span data-stu-id="778cc-146">Name</span></span>| <span data-ttu-id="778cc-147">Tipo</span><span class="sxs-lookup"><span data-stu-id="778cc-147">Type</span></span>| <span data-ttu-id="778cc-148">Descrição</span><span class="sxs-lookup"><span data-stu-id="778cc-148">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="778cc-149">String</span><span class="sxs-lookup"><span data-stu-id="778cc-149">String</span></span>|<span data-ttu-id="778cc-150">Uma ID de item formatada para APIs REST do Outlook</span><span class="sxs-lookup"><span data-stu-id="778cc-150">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="778cc-151">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="778cc-151">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="778cc-152">Um valor que indica a versão da API REST do Outlook usada para recuperar a ID do item.</span><span class="sxs-lookup"><span data-stu-id="778cc-152">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778cc-153">Requisitos</span><span class="sxs-lookup"><span data-stu-id="778cc-153">Requirements</span></span>

|<span data-ttu-id="778cc-154">Requisito</span><span class="sxs-lookup"><span data-stu-id="778cc-154">Requirement</span></span>| <span data-ttu-id="778cc-155">Valor</span><span class="sxs-lookup"><span data-stu-id="778cc-155">Value</span></span>|
|---|---|
|[<span data-ttu-id="778cc-156">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="778cc-156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778cc-157">1.3</span><span class="sxs-lookup"><span data-stu-id="778cc-157">1.3</span></span>|
|[<span data-ttu-id="778cc-158">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="778cc-158">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778cc-159">Restrito</span><span class="sxs-lookup"><span data-stu-id="778cc-159">Restricted</span></span>|
|[<span data-ttu-id="778cc-160">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="778cc-160">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778cc-161">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="778cc-161">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="778cc-162">Retorna:</span><span class="sxs-lookup"><span data-stu-id="778cc-162">Returns:</span></span>

<span data-ttu-id="778cc-163">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="778cc-163">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="778cc-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="778cc-164">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime"></a><span data-ttu-id="778cc-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="778cc-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span></span>

<span data-ttu-id="778cc-166">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="778cc-166">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="778cc-p105">As datas e horas usadas por um aplicativo de email para o Outlook ou o Outlook Web App podem usar fusos horários diferentes. O Outlook usa o fuso horário do computador cliente; o Outlook Web App usa o fuso horário definido na Centro de administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="778cc-p105">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="778cc-p106">Se o aplicativo de email estiver sendo executado no Outlook, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador do cliente. Se o aplicativo de email estiver sendo executado no Outlook Web App, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="778cc-p106">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778cc-172">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="778cc-172">Parameters</span></span>

|<span data-ttu-id="778cc-173">Nome</span><span class="sxs-lookup"><span data-stu-id="778cc-173">Name</span></span>| <span data-ttu-id="778cc-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="778cc-174">Type</span></span>| <span data-ttu-id="778cc-175">Descrição</span><span class="sxs-lookup"><span data-stu-id="778cc-175">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="778cc-176">Date</span><span class="sxs-lookup"><span data-stu-id="778cc-176">Date</span></span>|<span data-ttu-id="778cc-177">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="778cc-177">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778cc-178">Requisitos</span><span class="sxs-lookup"><span data-stu-id="778cc-178">Requirements</span></span>

|<span data-ttu-id="778cc-179">Requisito</span><span class="sxs-lookup"><span data-stu-id="778cc-179">Requirement</span></span>| <span data-ttu-id="778cc-180">Valor</span><span class="sxs-lookup"><span data-stu-id="778cc-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="778cc-181">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="778cc-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778cc-182">1.0</span><span class="sxs-lookup"><span data-stu-id="778cc-182">1.0</span></span>|
|[<span data-ttu-id="778cc-183">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="778cc-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778cc-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778cc-184">ReadItem</span></span>|
|[<span data-ttu-id="778cc-185">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="778cc-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778cc-186">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="778cc-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="778cc-187">Retorna:</span><span class="sxs-lookup"><span data-stu-id="778cc-187">Returns:</span></span>

<span data-ttu-id="778cc-188">Tipo: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="778cc-188">Type: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="778cc-189">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="778cc-189">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="778cc-190">Converte uma ID de item formatada para EWS no formato REST.</span><span class="sxs-lookup"><span data-stu-id="778cc-190">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="778cc-191">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="778cc-191">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="778cc-p107">IDs de itens recuperadas por EWS ou pela propriedade `itemId` usam um formato diferente daquele usado por APIs REST (como a [API do Email do Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou o [Microsoft Graph](https://graph.microsoft.io/)). O método `convertToRestId` converte uma ID formatada como EWS para o formato adequado para REST.</span><span class="sxs-lookup"><span data-stu-id="778cc-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778cc-194">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="778cc-194">Parameters</span></span>

|<span data-ttu-id="778cc-195">Nome</span><span class="sxs-lookup"><span data-stu-id="778cc-195">Name</span></span>| <span data-ttu-id="778cc-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="778cc-196">Type</span></span>| <span data-ttu-id="778cc-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="778cc-197">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="778cc-198">String</span><span class="sxs-lookup"><span data-stu-id="778cc-198">String</span></span>|<span data-ttu-id="778cc-199">Uma ID de item formatada para os Serviços Web do Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="778cc-199">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="778cc-200">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="778cc-200">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="778cc-201">Um valor que indica a versão da API REST do Outlook com a qual a ID convertida será usada.</span><span class="sxs-lookup"><span data-stu-id="778cc-201">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778cc-202">Requisitos</span><span class="sxs-lookup"><span data-stu-id="778cc-202">Requirements</span></span>

|<span data-ttu-id="778cc-203">Requisito</span><span class="sxs-lookup"><span data-stu-id="778cc-203">Requirement</span></span>| <span data-ttu-id="778cc-204">Valor</span><span class="sxs-lookup"><span data-stu-id="778cc-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="778cc-205">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="778cc-205">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778cc-206">1.3</span><span class="sxs-lookup"><span data-stu-id="778cc-206">1.3</span></span>|
|[<span data-ttu-id="778cc-207">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="778cc-207">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778cc-208">Restrito</span><span class="sxs-lookup"><span data-stu-id="778cc-208">Restricted</span></span>|
|[<span data-ttu-id="778cc-209">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="778cc-209">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778cc-210">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="778cc-210">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="778cc-211">Retorna:</span><span class="sxs-lookup"><span data-stu-id="778cc-211">Returns:</span></span>

<span data-ttu-id="778cc-212">Tipo: String</span><span class="sxs-lookup"><span data-stu-id="778cc-212">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="778cc-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="778cc-213">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="778cc-214">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="778cc-214">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="778cc-215">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="778cc-215">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="778cc-216">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="778cc-216">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778cc-217">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="778cc-217">Parameters</span></span>

|<span data-ttu-id="778cc-218">Nome</span><span class="sxs-lookup"><span data-stu-id="778cc-218">Name</span></span>| <span data-ttu-id="778cc-219">Tipo</span><span class="sxs-lookup"><span data-stu-id="778cc-219">Type</span></span>| <span data-ttu-id="778cc-220">Descrição</span><span class="sxs-lookup"><span data-stu-id="778cc-220">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="778cc-221">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="778cc-221">LocalClientTime</span></span>](/javascript/api/outlook_1_3/office.LocalClientTime)|<span data-ttu-id="778cc-222">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="778cc-222">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778cc-223">Requisitos</span><span class="sxs-lookup"><span data-stu-id="778cc-223">Requirements</span></span>

|<span data-ttu-id="778cc-224">Requisito</span><span class="sxs-lookup"><span data-stu-id="778cc-224">Requirement</span></span>| <span data-ttu-id="778cc-225">Valor</span><span class="sxs-lookup"><span data-stu-id="778cc-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="778cc-226">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="778cc-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778cc-227">1.0</span><span class="sxs-lookup"><span data-stu-id="778cc-227">1.0</span></span>|
|[<span data-ttu-id="778cc-228">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="778cc-228">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778cc-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778cc-229">ReadItem</span></span>|
|[<span data-ttu-id="778cc-230">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="778cc-230">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778cc-231">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="778cc-231">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="778cc-232">Retorna:</span><span class="sxs-lookup"><span data-stu-id="778cc-232">Returns:</span></span>

<span data-ttu-id="778cc-233">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="778cc-233">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="778cc-234">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="778cc-234">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="778cc-235">Date</span><span class="sxs-lookup"><span data-stu-id="778cc-235">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="778cc-236">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="778cc-236">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="778cc-237">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="778cc-237">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="778cc-238">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="778cc-238">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="778cc-239">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="778cc-239">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="778cc-p108">No Outlook para Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir uma instância da série. Isso ocorre porque no Outlook para Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="778cc-p108">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="778cc-242">No Outlook Web App, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="778cc-242">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="778cc-243">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="778cc-243">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778cc-244">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="778cc-244">Parameters</span></span>

|<span data-ttu-id="778cc-245">Nome</span><span class="sxs-lookup"><span data-stu-id="778cc-245">Name</span></span>| <span data-ttu-id="778cc-246">Tipo</span><span class="sxs-lookup"><span data-stu-id="778cc-246">Type</span></span>| <span data-ttu-id="778cc-247">Descrição</span><span class="sxs-lookup"><span data-stu-id="778cc-247">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="778cc-248">String</span><span class="sxs-lookup"><span data-stu-id="778cc-248">String</span></span>|<span data-ttu-id="778cc-249">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="778cc-249">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778cc-250">Requisitos</span><span class="sxs-lookup"><span data-stu-id="778cc-250">Requirements</span></span>

|<span data-ttu-id="778cc-251">Requisito</span><span class="sxs-lookup"><span data-stu-id="778cc-251">Requirement</span></span>| <span data-ttu-id="778cc-252">Valor</span><span class="sxs-lookup"><span data-stu-id="778cc-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="778cc-253">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="778cc-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778cc-254">1.0</span><span class="sxs-lookup"><span data-stu-id="778cc-254">1.0</span></span>|
|[<span data-ttu-id="778cc-255">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="778cc-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778cc-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778cc-256">ReadItem</span></span>|
|[<span data-ttu-id="778cc-257">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="778cc-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778cc-258">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="778cc-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778cc-259">Exemplo</span><span class="sxs-lookup"><span data-stu-id="778cc-259">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="778cc-260">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="778cc-260">displayMessageForm(itemId)</span></span>

<span data-ttu-id="778cc-261">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="778cc-261">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="778cc-262">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="778cc-262">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="778cc-263">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="778cc-263">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="778cc-264">No Outlook Web App, este método abre o formulário especificado somente se o corpo do formulário for menor que ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="778cc-264">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="778cc-265">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="778cc-265">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="778cc-p109">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="778cc-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778cc-268">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="778cc-268">Parameters</span></span>

|<span data-ttu-id="778cc-269">Nome</span><span class="sxs-lookup"><span data-stu-id="778cc-269">Name</span></span>| <span data-ttu-id="778cc-270">Tipo</span><span class="sxs-lookup"><span data-stu-id="778cc-270">Type</span></span>| <span data-ttu-id="778cc-271">Descrição</span><span class="sxs-lookup"><span data-stu-id="778cc-271">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="778cc-272">String</span><span class="sxs-lookup"><span data-stu-id="778cc-272">String</span></span>|<span data-ttu-id="778cc-273">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="778cc-273">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778cc-274">Requisitos</span><span class="sxs-lookup"><span data-stu-id="778cc-274">Requirements</span></span>

|<span data-ttu-id="778cc-275">Requisito</span><span class="sxs-lookup"><span data-stu-id="778cc-275">Requirement</span></span>| <span data-ttu-id="778cc-276">Valor</span><span class="sxs-lookup"><span data-stu-id="778cc-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="778cc-277">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="778cc-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778cc-278">1.0</span><span class="sxs-lookup"><span data-stu-id="778cc-278">1.0</span></span>|
|[<span data-ttu-id="778cc-279">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="778cc-279">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778cc-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778cc-280">ReadItem</span></span>|
|[<span data-ttu-id="778cc-281">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="778cc-281">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778cc-282">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="778cc-282">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778cc-283">Exemplo</span><span class="sxs-lookup"><span data-stu-id="778cc-283">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="778cc-284">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="778cc-284">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="778cc-285">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="778cc-285">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="778cc-286">Não há suporte para esse método no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="778cc-286">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="778cc-p110">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="778cc-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="778cc-p111">No Outlook Web App e no OWA para Dispositivos, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="778cc-p111">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="778cc-p112">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="778cc-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="778cc-294">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="778cc-294">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778cc-295">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="778cc-295">Parameters</span></span>

|<span data-ttu-id="778cc-296">Nome</span><span class="sxs-lookup"><span data-stu-id="778cc-296">Name</span></span>| <span data-ttu-id="778cc-297">Tipo</span><span class="sxs-lookup"><span data-stu-id="778cc-297">Type</span></span>| <span data-ttu-id="778cc-298">Descrição</span><span class="sxs-lookup"><span data-stu-id="778cc-298">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="778cc-299">Object</span><span class="sxs-lookup"><span data-stu-id="778cc-299">Object</span></span> | <span data-ttu-id="778cc-300">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="778cc-300">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="778cc-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="778cc-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="778cc-p113">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="778cc-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="778cc-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="778cc-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="778cc-p114">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="778cc-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="778cc-307">Data</span><span class="sxs-lookup"><span data-stu-id="778cc-307">Date</span></span> | <span data-ttu-id="778cc-308">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="778cc-308">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="778cc-309">Data</span><span class="sxs-lookup"><span data-stu-id="778cc-309">Date</span></span> | <span data-ttu-id="778cc-310">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="778cc-310">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="778cc-311">String</span><span class="sxs-lookup"><span data-stu-id="778cc-311">String</span></span> | <span data-ttu-id="778cc-p115">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="778cc-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="778cc-314">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="778cc-314">Array.&lt;String&gt;</span></span> | <span data-ttu-id="778cc-p116">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="778cc-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="778cc-317">String</span><span class="sxs-lookup"><span data-stu-id="778cc-317">String</span></span> | <span data-ttu-id="778cc-p117">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="778cc-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="778cc-320">String</span><span class="sxs-lookup"><span data-stu-id="778cc-320">String</span></span> | <span data-ttu-id="778cc-p118">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="778cc-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="778cc-323">Requisitos</span><span class="sxs-lookup"><span data-stu-id="778cc-323">Requirements</span></span>

|<span data-ttu-id="778cc-324">Requisito</span><span class="sxs-lookup"><span data-stu-id="778cc-324">Requirement</span></span>| <span data-ttu-id="778cc-325">Valor</span><span class="sxs-lookup"><span data-stu-id="778cc-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="778cc-326">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="778cc-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778cc-327">1.0</span><span class="sxs-lookup"><span data-stu-id="778cc-327">1.0</span></span>|
|[<span data-ttu-id="778cc-328">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="778cc-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778cc-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778cc-329">ReadItem</span></span>|
|[<span data-ttu-id="778cc-330">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="778cc-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778cc-331">Read</span><span class="sxs-lookup"><span data-stu-id="778cc-331">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778cc-332">Exemplo</span><span class="sxs-lookup"><span data-stu-id="778cc-332">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="778cc-333">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="778cc-333">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="778cc-334">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="778cc-334">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="778cc-p119">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="778cc-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="778cc-p120">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="778cc-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="778cc-340">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync` em modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="778cc-340">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="778cc-p121">No modo de composição, você deve chamar o método [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) para obter um identificador de item para passar ao método `getCallbackTokenAsync`. Seu aplicativo deve ter permissões **ReadWriteItem** para chamar o método `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="778cc-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778cc-343">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="778cc-343">Parameters</span></span>

|<span data-ttu-id="778cc-344">Nome</span><span class="sxs-lookup"><span data-stu-id="778cc-344">Name</span></span>| <span data-ttu-id="778cc-345">Tipo</span><span class="sxs-lookup"><span data-stu-id="778cc-345">Type</span></span>| <span data-ttu-id="778cc-346">Atributos</span><span class="sxs-lookup"><span data-stu-id="778cc-346">Attributes</span></span>| <span data-ttu-id="778cc-347">Descrição</span><span class="sxs-lookup"><span data-stu-id="778cc-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="778cc-348">function</span><span class="sxs-lookup"><span data-stu-id="778cc-348">function</span></span>||<span data-ttu-id="778cc-p122">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult). O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="778cc-p122">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="778cc-351">Objeto</span><span class="sxs-lookup"><span data-stu-id="778cc-351">Object</span></span>| <span data-ttu-id="778cc-352">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="778cc-352">&lt;optional&gt;</span></span>|<span data-ttu-id="778cc-353">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="778cc-353">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778cc-354">Requisitos</span><span class="sxs-lookup"><span data-stu-id="778cc-354">Requirements</span></span>

|<span data-ttu-id="778cc-355">Requisito</span><span class="sxs-lookup"><span data-stu-id="778cc-355">Requirement</span></span>| <span data-ttu-id="778cc-356">Valor</span><span class="sxs-lookup"><span data-stu-id="778cc-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="778cc-357">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="778cc-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778cc-358">1.3</span><span class="sxs-lookup"><span data-stu-id="778cc-358">1.3</span></span>|
|[<span data-ttu-id="778cc-359">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="778cc-359">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778cc-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778cc-360">ReadItem</span></span>|
|[<span data-ttu-id="778cc-361">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="778cc-361">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778cc-362">Redação e leitura</span><span class="sxs-lookup"><span data-stu-id="778cc-362">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="778cc-363">Exemplo</span><span class="sxs-lookup"><span data-stu-id="778cc-363">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="778cc-364">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="778cc-364">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="778cc-365">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="778cc-365">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="778cc-366">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="778cc-366">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="778cc-367">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="778cc-367">Parameters</span></span>

|<span data-ttu-id="778cc-368">Nome</span><span class="sxs-lookup"><span data-stu-id="778cc-368">Name</span></span>| <span data-ttu-id="778cc-369">Tipo</span><span class="sxs-lookup"><span data-stu-id="778cc-369">Type</span></span>| <span data-ttu-id="778cc-370">Atributos</span><span class="sxs-lookup"><span data-stu-id="778cc-370">Attributes</span></span>| <span data-ttu-id="778cc-371">Descrição</span><span class="sxs-lookup"><span data-stu-id="778cc-371">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="778cc-372">function</span><span class="sxs-lookup"><span data-stu-id="778cc-372">function</span></span>||<span data-ttu-id="778cc-373">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="778cc-373">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="778cc-374">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="778cc-374">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="778cc-375">Object</span><span class="sxs-lookup"><span data-stu-id="778cc-375">Object</span></span>| <span data-ttu-id="778cc-376">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="778cc-376">&lt;optional&gt;</span></span>|<span data-ttu-id="778cc-377">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="778cc-377">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778cc-378">Requisitos</span><span class="sxs-lookup"><span data-stu-id="778cc-378">Requirements</span></span>

|<span data-ttu-id="778cc-379">Requisito</span><span class="sxs-lookup"><span data-stu-id="778cc-379">Requirement</span></span>| <span data-ttu-id="778cc-380">Valor</span><span class="sxs-lookup"><span data-stu-id="778cc-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="778cc-381">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="778cc-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778cc-382">1.0</span><span class="sxs-lookup"><span data-stu-id="778cc-382">1.0</span></span>|
|[<span data-ttu-id="778cc-383">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="778cc-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778cc-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="778cc-384">ReadItem</span></span>|
|[<span data-ttu-id="778cc-385">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="778cc-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778cc-386">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="778cc-386">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778cc-387">Exemplo</span><span class="sxs-lookup"><span data-stu-id="778cc-387">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="778cc-388">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="778cc-388">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="778cc-389">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="778cc-389">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="778cc-390">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="778cc-390">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="778cc-391">No Outlook para iOS ou no Outlook para Android</span><span class="sxs-lookup"><span data-stu-id="778cc-391">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="778cc-392">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="778cc-392">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="778cc-393">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="778cc-393">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="778cc-394">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="778cc-394">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="778cc-395">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="778cc-395">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="778cc-396">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="778cc-396">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="778cc-397">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="778cc-397">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="778cc-p124">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="778cc-p124">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="778cc-400">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="778cc-400">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="778cc-401">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="778cc-401">Version differences</span></span>

<span data-ttu-id="778cc-402">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="778cc-402">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="778cc-p125">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="778cc-p125">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="778cc-406">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="778cc-406">Parameters</span></span>

|<span data-ttu-id="778cc-407">Nome</span><span class="sxs-lookup"><span data-stu-id="778cc-407">Name</span></span>| <span data-ttu-id="778cc-408">Tipo</span><span class="sxs-lookup"><span data-stu-id="778cc-408">Type</span></span>| <span data-ttu-id="778cc-409">Atributos</span><span class="sxs-lookup"><span data-stu-id="778cc-409">Attributes</span></span>| <span data-ttu-id="778cc-410">Descrição</span><span class="sxs-lookup"><span data-stu-id="778cc-410">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="778cc-411">String</span><span class="sxs-lookup"><span data-stu-id="778cc-411">String</span></span>||<span data-ttu-id="778cc-412">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="778cc-412">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="778cc-413">function</span><span class="sxs-lookup"><span data-stu-id="778cc-413">function</span></span>||<span data-ttu-id="778cc-414">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="778cc-414">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="778cc-415">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="778cc-415">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="778cc-416">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="778cc-416">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="778cc-417">Objeto</span><span class="sxs-lookup"><span data-stu-id="778cc-417">Object</span></span>| <span data-ttu-id="778cc-418">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="778cc-418">&lt;optional&gt;</span></span>|<span data-ttu-id="778cc-419">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="778cc-419">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778cc-420">Requisitos</span><span class="sxs-lookup"><span data-stu-id="778cc-420">Requirements</span></span>

|<span data-ttu-id="778cc-421">Requisito</span><span class="sxs-lookup"><span data-stu-id="778cc-421">Requirement</span></span>| <span data-ttu-id="778cc-422">Valor</span><span class="sxs-lookup"><span data-stu-id="778cc-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="778cc-423">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="778cc-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778cc-424">1.0</span><span class="sxs-lookup"><span data-stu-id="778cc-424">1.0</span></span>|
|[<span data-ttu-id="778cc-425">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="778cc-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778cc-426">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="778cc-426">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="778cc-427">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="778cc-427">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778cc-428">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="778cc-428">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778cc-429">Exemplo</span><span class="sxs-lookup"><span data-stu-id="778cc-429">Example</span></span>

<span data-ttu-id="778cc-430">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="778cc-430">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
