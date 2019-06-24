---
title: Office. Context. Mailbox – conjunto de requisitos 1,1
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 256bd2b992531fa52953098893025e4a006caf08
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127489"
---
# <a name="mailbox"></a><span data-ttu-id="2bda8-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="2bda8-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="2bda8-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="2bda8-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="2bda8-104">Fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="2bda8-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2bda8-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2bda8-105">Requirements</span></span>

|<span data-ttu-id="2bda8-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="2bda8-106">Requirement</span></span>| <span data-ttu-id="2bda8-107">Valor</span><span class="sxs-lookup"><span data-stu-id="2bda8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2bda8-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2bda8-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2bda8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="2bda8-109">1.0</span></span>|
|[<span data-ttu-id="2bda8-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2bda8-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2bda8-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="2bda8-111">Restricted</span></span>|
|[<span data-ttu-id="2bda8-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2bda8-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2bda8-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2bda8-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="2bda8-114">Namespaces</span><span class="sxs-lookup"><span data-stu-id="2bda8-114">Namespaces</span></span>

<span data-ttu-id="2bda8-115">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="2bda8-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="2bda8-116">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="2bda8-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="2bda8-117">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="2bda8-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="2bda8-118">Membros</span><span class="sxs-lookup"><span data-stu-id="2bda8-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="2bda8-119">ewsUrl: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2bda8-119">ewsUrl: String</span></span>

<span data-ttu-id="2bda8-120">Obtém a URL do ponto de extremidade dos Serviços Web do Exchange (EWS) para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="2bda8-120">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="2bda8-121">Somente modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="2bda8-121">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2bda8-122">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="2bda8-122">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2bda8-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="2bda8-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="2bda8-125">Tipo</span><span class="sxs-lookup"><span data-stu-id="2bda8-125">Type</span></span>

*   <span data-ttu-id="2bda8-126">String</span><span class="sxs-lookup"><span data-stu-id="2bda8-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2bda8-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2bda8-127">Requirements</span></span>

|<span data-ttu-id="2bda8-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="2bda8-128">Requirement</span></span>| <span data-ttu-id="2bda8-129">Valor</span><span class="sxs-lookup"><span data-stu-id="2bda8-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="2bda8-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2bda8-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2bda8-131">1.0</span><span class="sxs-lookup"><span data-stu-id="2bda8-131">1.0</span></span>|
|[<span data-ttu-id="2bda8-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2bda8-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2bda8-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2bda8-133">ReadItem</span></span>|
|[<span data-ttu-id="2bda8-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2bda8-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2bda8-135">Read</span><span class="sxs-lookup"><span data-stu-id="2bda8-135">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="2bda8-136">Métodos</span><span class="sxs-lookup"><span data-stu-id="2bda8-136">Methods</span></span>

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime"></a><span data-ttu-id="2bda8-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="2bda8-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)}</span></span>

<span data-ttu-id="2bda8-138">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="2bda8-138">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="2bda8-139">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para datas e horas.</span><span class="sxs-lookup"><span data-stu-id="2bda8-139">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="2bda8-140">O Outlook em uma área de trabalho usa o fuso horário do computador cliente; O Outlook na Web usa o fuso horário definido no centro de administração do Exchange (Eat).</span><span class="sxs-lookup"><span data-stu-id="2bda8-140">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="2bda8-141">Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="2bda8-141">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="2bda8-142">Se o aplicativo de email estiver em execução no Outlook em um cliente desktop `convertToLocalClientTime` , o método retornará um objeto Dictionary com os valores definidos para o fuso horário do computador cliente.</span><span class="sxs-lookup"><span data-stu-id="2bda8-142">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="2bda8-143">Se o aplicativo de email estiver em execução no Outlook na Web, `convertToLocalClientTime` o método retornará um objeto Dictionary com os valores definidos para o fuso horário especificado no Eat.</span><span class="sxs-lookup"><span data-stu-id="2bda8-143">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2bda8-144">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2bda8-144">Parameters</span></span>

|<span data-ttu-id="2bda8-145">Nome</span><span class="sxs-lookup"><span data-stu-id="2bda8-145">Name</span></span>| <span data-ttu-id="2bda8-146">Tipo</span><span class="sxs-lookup"><span data-stu-id="2bda8-146">Type</span></span>| <span data-ttu-id="2bda8-147">Descrição</span><span class="sxs-lookup"><span data-stu-id="2bda8-147">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="2bda8-148">Date</span><span class="sxs-lookup"><span data-stu-id="2bda8-148">Date</span></span>|<span data-ttu-id="2bda8-149">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="2bda8-149">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2bda8-150">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2bda8-150">Requirements</span></span>

|<span data-ttu-id="2bda8-151">Requisito</span><span class="sxs-lookup"><span data-stu-id="2bda8-151">Requirement</span></span>| <span data-ttu-id="2bda8-152">Valor</span><span class="sxs-lookup"><span data-stu-id="2bda8-152">Value</span></span>|
|---|---|
|[<span data-ttu-id="2bda8-153">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2bda8-153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2bda8-154">1.0</span><span class="sxs-lookup"><span data-stu-id="2bda8-154">1.0</span></span>|
|[<span data-ttu-id="2bda8-155">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2bda8-155">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2bda8-156">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2bda8-156">ReadItem</span></span>|
|[<span data-ttu-id="2bda8-157">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2bda8-157">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2bda8-158">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2bda8-158">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2bda8-159">Retorna:</span><span class="sxs-lookup"><span data-stu-id="2bda8-159">Returns:</span></span>

<span data-ttu-id="2bda8-160">Tipo: [LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="2bda8-160">Type: [LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)</span></span>

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="2bda8-161">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="2bda8-161">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="2bda8-162">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="2bda8-162">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="2bda8-163">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="2bda8-163">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2bda8-164">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2bda8-164">Parameters</span></span>

|<span data-ttu-id="2bda8-165">Nome</span><span class="sxs-lookup"><span data-stu-id="2bda8-165">Name</span></span>| <span data-ttu-id="2bda8-166">Tipo</span><span class="sxs-lookup"><span data-stu-id="2bda8-166">Type</span></span>| <span data-ttu-id="2bda8-167">Descrição</span><span class="sxs-lookup"><span data-stu-id="2bda8-167">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="2bda8-168">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="2bda8-168">LocalClientTime</span></span>](/javascript/api/outlook_1_1/office.LocalClientTime)|<span data-ttu-id="2bda8-169">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="2bda8-169">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2bda8-170">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2bda8-170">Requirements</span></span>

|<span data-ttu-id="2bda8-171">Requisito</span><span class="sxs-lookup"><span data-stu-id="2bda8-171">Requirement</span></span>| <span data-ttu-id="2bda8-172">Valor</span><span class="sxs-lookup"><span data-stu-id="2bda8-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="2bda8-173">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2bda8-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2bda8-174">1.0</span><span class="sxs-lookup"><span data-stu-id="2bda8-174">1.0</span></span>|
|[<span data-ttu-id="2bda8-175">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2bda8-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2bda8-176">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2bda8-176">ReadItem</span></span>|
|[<span data-ttu-id="2bda8-177">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2bda8-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2bda8-178">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2bda8-178">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2bda8-179">Retorna:</span><span class="sxs-lookup"><span data-stu-id="2bda8-179">Returns:</span></span>

<span data-ttu-id="2bda8-180">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="2bda8-180">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="2bda8-181">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="2bda8-181">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2bda8-182">Date</span><span class="sxs-lookup"><span data-stu-id="2bda8-182">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="2bda8-183">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="2bda8-183">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="2bda8-184">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="2bda8-184">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2bda8-185">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="2bda8-185">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2bda8-186">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="2bda8-186">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="2bda8-187">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente ou o compromisso mestre de uma série recorrente, mas não é possível exibir uma instância da série.</span><span class="sxs-lookup"><span data-stu-id="2bda8-187">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="2bda8-188">Isso ocorre porque, no Outlook no Mac, você não pode acessar as propriedades (incluindo a ID do item) de instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="2bda8-188">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="2bda8-189">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres 32 KB.</span><span class="sxs-lookup"><span data-stu-id="2bda8-189">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="2bda8-190">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="2bda8-190">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2bda8-191">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2bda8-191">Parameters</span></span>

|<span data-ttu-id="2bda8-192">Nome</span><span class="sxs-lookup"><span data-stu-id="2bda8-192">Name</span></span>| <span data-ttu-id="2bda8-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="2bda8-193">Type</span></span>| <span data-ttu-id="2bda8-194">Descrição</span><span class="sxs-lookup"><span data-stu-id="2bda8-194">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2bda8-195">String</span><span class="sxs-lookup"><span data-stu-id="2bda8-195">String</span></span>|<span data-ttu-id="2bda8-196">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="2bda8-196">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2bda8-197">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2bda8-197">Requirements</span></span>

|<span data-ttu-id="2bda8-198">Requisito</span><span class="sxs-lookup"><span data-stu-id="2bda8-198">Requirement</span></span>| <span data-ttu-id="2bda8-199">Valor</span><span class="sxs-lookup"><span data-stu-id="2bda8-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="2bda8-200">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2bda8-200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2bda8-201">1.0</span><span class="sxs-lookup"><span data-stu-id="2bda8-201">1.0</span></span>|
|[<span data-ttu-id="2bda8-202">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2bda8-202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2bda8-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2bda8-203">ReadItem</span></span>|
|[<span data-ttu-id="2bda8-204">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2bda8-204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2bda8-205">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2bda8-205">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2bda8-206">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2bda8-206">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="2bda8-207">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="2bda8-207">displayMessageForm(itemId)</span></span>

<span data-ttu-id="2bda8-208">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="2bda8-208">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="2bda8-209">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="2bda8-209">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2bda8-210">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="2bda8-210">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="2bda8-211">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="2bda8-211">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="2bda8-212">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="2bda8-212">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="2bda8-p106">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="2bda8-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2bda8-215">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2bda8-215">Parameters</span></span>

|<span data-ttu-id="2bda8-216">Nome</span><span class="sxs-lookup"><span data-stu-id="2bda8-216">Name</span></span>| <span data-ttu-id="2bda8-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="2bda8-217">Type</span></span>| <span data-ttu-id="2bda8-218">Descrição</span><span class="sxs-lookup"><span data-stu-id="2bda8-218">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2bda8-219">String</span><span class="sxs-lookup"><span data-stu-id="2bda8-219">String</span></span>|<span data-ttu-id="2bda8-220">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="2bda8-220">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2bda8-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2bda8-221">Requirements</span></span>

|<span data-ttu-id="2bda8-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="2bda8-222">Requirement</span></span>| <span data-ttu-id="2bda8-223">Valor</span><span class="sxs-lookup"><span data-stu-id="2bda8-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="2bda8-224">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2bda8-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2bda8-225">1.0</span><span class="sxs-lookup"><span data-stu-id="2bda8-225">1.0</span></span>|
|[<span data-ttu-id="2bda8-226">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2bda8-226">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2bda8-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2bda8-227">ReadItem</span></span>|
|[<span data-ttu-id="2bda8-228">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2bda8-228">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2bda8-229">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2bda8-229">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2bda8-230">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2bda8-230">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="2bda8-231">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="2bda8-231">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="2bda8-232">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="2bda8-232">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2bda8-233">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="2bda8-233">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2bda8-p107">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="2bda8-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="2bda8-236">No Outlook na Web e dispositivos móveis, este método sempre exibe um formulário com um campo participantes.</span><span class="sxs-lookup"><span data-stu-id="2bda8-236">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="2bda8-237">Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="2bda8-237">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="2bda8-238">Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="2bda8-238">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="2bda8-p109">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="2bda8-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="2bda8-241">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="2bda8-241">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2bda8-242">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2bda8-242">Parameters</span></span>

|<span data-ttu-id="2bda8-243">Nome</span><span class="sxs-lookup"><span data-stu-id="2bda8-243">Name</span></span>| <span data-ttu-id="2bda8-244">Tipo</span><span class="sxs-lookup"><span data-stu-id="2bda8-244">Type</span></span>| <span data-ttu-id="2bda8-245">Descrição</span><span class="sxs-lookup"><span data-stu-id="2bda8-245">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="2bda8-246">Objeto</span><span class="sxs-lookup"><span data-stu-id="2bda8-246">Object</span></span> | <span data-ttu-id="2bda8-247">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="2bda8-247">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="2bda8-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2bda8-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2bda8-p110">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="2bda8-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="2bda8-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="2bda8-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="2bda8-p111">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="2bda8-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="2bda8-254">Data</span><span class="sxs-lookup"><span data-stu-id="2bda8-254">Date</span></span> | <span data-ttu-id="2bda8-255">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="2bda8-255">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="2bda8-256">Data</span><span class="sxs-lookup"><span data-stu-id="2bda8-256">Date</span></span> | <span data-ttu-id="2bda8-257">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="2bda8-257">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="2bda8-258">String</span><span class="sxs-lookup"><span data-stu-id="2bda8-258">String</span></span> | <span data-ttu-id="2bda8-p112">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2bda8-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="2bda8-261">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="2bda8-261">Array.&lt;String&gt;</span></span> | <span data-ttu-id="2bda8-p113">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="2bda8-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="2bda8-264">String</span><span class="sxs-lookup"><span data-stu-id="2bda8-264">String</span></span> | <span data-ttu-id="2bda8-p114">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="2bda8-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="2bda8-267">String</span><span class="sxs-lookup"><span data-stu-id="2bda8-267">String</span></span> | <span data-ttu-id="2bda8-p115">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="2bda8-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2bda8-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2bda8-270">Requirements</span></span>

|<span data-ttu-id="2bda8-271">Requisito</span><span class="sxs-lookup"><span data-stu-id="2bda8-271">Requirement</span></span>| <span data-ttu-id="2bda8-272">Valor</span><span class="sxs-lookup"><span data-stu-id="2bda8-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="2bda8-273">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2bda8-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2bda8-274">1.0</span><span class="sxs-lookup"><span data-stu-id="2bda8-274">1.0</span></span>|
|[<span data-ttu-id="2bda8-275">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2bda8-275">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2bda8-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2bda8-276">ReadItem</span></span>|
|[<span data-ttu-id="2bda8-277">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2bda8-277">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2bda8-278">Read</span><span class="sxs-lookup"><span data-stu-id="2bda8-278">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2bda8-279">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2bda8-279">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="2bda8-280">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2bda8-280">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="2bda8-281">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="2bda8-281">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="2bda8-p116">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="2bda8-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="2bda8-p117">Você pode passar o token e um identificador de anexo ou um identificador de item a um sistema de terceiros. O sistema de terceiros usa o token como portador da autorização para chamar as operações [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) dos Serviços Web do Exchange (EWS) para retornar um anexo ou item. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="2bda8-p117">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="2bda8-287">Seu aplicativo deve ter a permissão **ReadItem** especificada em seu manifesto para chamar o método `getCallbackTokenAsync`.</span><span class="sxs-lookup"><span data-stu-id="2bda8-287">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2bda8-288">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2bda8-288">Parameters</span></span>

|<span data-ttu-id="2bda8-289">Nome</span><span class="sxs-lookup"><span data-stu-id="2bda8-289">Name</span></span>| <span data-ttu-id="2bda8-290">Tipo</span><span class="sxs-lookup"><span data-stu-id="2bda8-290">Type</span></span>| <span data-ttu-id="2bda8-291">Atributos</span><span class="sxs-lookup"><span data-stu-id="2bda8-291">Attributes</span></span>| <span data-ttu-id="2bda8-292">Descrição</span><span class="sxs-lookup"><span data-stu-id="2bda8-292">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2bda8-293">function</span><span class="sxs-lookup"><span data-stu-id="2bda8-293">function</span></span>||<span data-ttu-id="2bda8-294">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2bda8-294">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2bda8-295">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2bda8-295">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="2bda8-296">Objeto</span><span class="sxs-lookup"><span data-stu-id="2bda8-296">Object</span></span>| <span data-ttu-id="2bda8-297">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2bda8-297">&lt;optional&gt;</span></span>|<span data-ttu-id="2bda8-298">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="2bda8-298">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2bda8-299">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2bda8-299">Requirements</span></span>

|<span data-ttu-id="2bda8-300">Requisito</span><span class="sxs-lookup"><span data-stu-id="2bda8-300">Requirement</span></span>| <span data-ttu-id="2bda8-301">Valor</span><span class="sxs-lookup"><span data-stu-id="2bda8-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="2bda8-302">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2bda8-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2bda8-303">1.0</span><span class="sxs-lookup"><span data-stu-id="2bda8-303">1.0</span></span>|
|[<span data-ttu-id="2bda8-304">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2bda8-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2bda8-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2bda8-305">ReadItem</span></span>|
|[<span data-ttu-id="2bda8-306">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2bda8-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2bda8-307">Read</span><span class="sxs-lookup"><span data-stu-id="2bda8-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2bda8-308">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2bda8-308">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="2bda8-309">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2bda8-309">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="2bda8-310">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="2bda8-310">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="2bda8-311">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="2bda8-311">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="2bda8-312">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2bda8-312">Parameters</span></span>

|<span data-ttu-id="2bda8-313">Nome</span><span class="sxs-lookup"><span data-stu-id="2bda8-313">Name</span></span>| <span data-ttu-id="2bda8-314">Tipo</span><span class="sxs-lookup"><span data-stu-id="2bda8-314">Type</span></span>| <span data-ttu-id="2bda8-315">Atributos</span><span class="sxs-lookup"><span data-stu-id="2bda8-315">Attributes</span></span>| <span data-ttu-id="2bda8-316">Descrição</span><span class="sxs-lookup"><span data-stu-id="2bda8-316">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2bda8-317">function</span><span class="sxs-lookup"><span data-stu-id="2bda8-317">function</span></span>||<span data-ttu-id="2bda8-318">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2bda8-318">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2bda8-319">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2bda8-319">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="2bda8-320">Objeto</span><span class="sxs-lookup"><span data-stu-id="2bda8-320">Object</span></span>| <span data-ttu-id="2bda8-321">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2bda8-321">&lt;optional&gt;</span></span>|<span data-ttu-id="2bda8-322">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="2bda8-322">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2bda8-323">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2bda8-323">Requirements</span></span>

|<span data-ttu-id="2bda8-324">Requisito</span><span class="sxs-lookup"><span data-stu-id="2bda8-324">Requirement</span></span>| <span data-ttu-id="2bda8-325">Valor</span><span class="sxs-lookup"><span data-stu-id="2bda8-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="2bda8-326">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2bda8-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2bda8-327">1.0</span><span class="sxs-lookup"><span data-stu-id="2bda8-327">1.0</span></span>|
|[<span data-ttu-id="2bda8-328">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2bda8-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2bda8-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2bda8-329">ReadItem</span></span>|
|[<span data-ttu-id="2bda8-330">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="2bda8-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2bda8-331">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2bda8-331">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2bda8-332">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2bda8-332">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="2bda8-333">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2bda8-333">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="2bda8-334">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="2bda8-334">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="2bda8-335">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="2bda8-335">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="2bda8-336">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="2bda8-336">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="2bda8-337">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="2bda8-337">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="2bda8-338">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="2bda8-338">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="2bda8-339">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="2bda8-339">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="2bda8-340">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="2bda8-340">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="2bda8-341">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="2bda8-341">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="2bda8-342">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="2bda8-342">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="2bda8-p119">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="2bda8-p119">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="2bda8-345">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="2bda8-345">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="2bda8-346">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="2bda8-346">Version differences</span></span>

<span data-ttu-id="2bda8-347">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="2bda8-347">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="2bda8-p120">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="2bda8-p120">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2bda8-351">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="2bda8-351">Parameters</span></span>

|<span data-ttu-id="2bda8-352">Nome</span><span class="sxs-lookup"><span data-stu-id="2bda8-352">Name</span></span>| <span data-ttu-id="2bda8-353">Tipo</span><span class="sxs-lookup"><span data-stu-id="2bda8-353">Type</span></span>| <span data-ttu-id="2bda8-354">Atributos</span><span class="sxs-lookup"><span data-stu-id="2bda8-354">Attributes</span></span>| <span data-ttu-id="2bda8-355">Descrição</span><span class="sxs-lookup"><span data-stu-id="2bda8-355">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="2bda8-356">String</span><span class="sxs-lookup"><span data-stu-id="2bda8-356">String</span></span>||<span data-ttu-id="2bda8-357">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="2bda8-357">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="2bda8-358">function</span><span class="sxs-lookup"><span data-stu-id="2bda8-358">function</span></span>||<span data-ttu-id="2bda8-359">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2bda8-359">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2bda8-360">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2bda8-360">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="2bda8-361">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="2bda8-361">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="2bda8-362">Objeto</span><span class="sxs-lookup"><span data-stu-id="2bda8-362">Object</span></span>| <span data-ttu-id="2bda8-363">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="2bda8-363">&lt;optional&gt;</span></span>|<span data-ttu-id="2bda8-364">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="2bda8-364">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2bda8-365">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2bda8-365">Requirements</span></span>

|<span data-ttu-id="2bda8-366">Requisito</span><span class="sxs-lookup"><span data-stu-id="2bda8-366">Requirement</span></span>| <span data-ttu-id="2bda8-367">Valor</span><span class="sxs-lookup"><span data-stu-id="2bda8-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="2bda8-368">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2bda8-368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2bda8-369">1.0</span><span class="sxs-lookup"><span data-stu-id="2bda8-369">1.0</span></span>|
|[<span data-ttu-id="2bda8-370">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2bda8-370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2bda8-371">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="2bda8-371">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="2bda8-372">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2bda8-372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2bda8-373">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2bda8-373">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2bda8-374">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2bda8-374">Example</span></span>

<span data-ttu-id="2bda8-375">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="2bda8-375">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
