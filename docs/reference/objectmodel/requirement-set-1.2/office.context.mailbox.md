---
title: Office. Context. Mailbox – conjunto de requisitos 1,2
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 542e8c9899c2d4a3c5b4546c3d5a73ba0d3c3a7e
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626996"
---
# <a name="mailbox"></a><span data-ttu-id="f90d7-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="f90d7-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="f90d7-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="f90d7-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="f90d7-104">Fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="f90d7-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f90d7-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f90d7-105">Requirements</span></span>

|<span data-ttu-id="f90d7-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="f90d7-106">Requirement</span></span>| <span data-ttu-id="f90d7-107">Valor</span><span class="sxs-lookup"><span data-stu-id="f90d7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f90d7-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f90d7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f90d7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f90d7-109">1.0</span></span>|
|[<span data-ttu-id="f90d7-110">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f90d7-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f90d7-111">Restrito</span><span class="sxs-lookup"><span data-stu-id="f90d7-111">Restricted</span></span>|
|[<span data-ttu-id="f90d7-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f90d7-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f90d7-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f90d7-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f90d7-114">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="f90d7-114">Members and methods</span></span>

| <span data-ttu-id="f90d7-115">Membro</span><span class="sxs-lookup"><span data-stu-id="f90d7-115">Member</span></span> | <span data-ttu-id="f90d7-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="f90d7-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f90d7-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="f90d7-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="f90d7-118">Membro</span><span class="sxs-lookup"><span data-stu-id="f90d7-118">Member</span></span> |
| [<span data-ttu-id="f90d7-119">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="f90d7-119">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="f90d7-120">Método</span><span class="sxs-lookup"><span data-stu-id="f90d7-120">Method</span></span> |
| [<span data-ttu-id="f90d7-121">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="f90d7-121">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="f90d7-122">Método</span><span class="sxs-lookup"><span data-stu-id="f90d7-122">Method</span></span> |
| [<span data-ttu-id="f90d7-123">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="f90d7-123">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="f90d7-124">Método</span><span class="sxs-lookup"><span data-stu-id="f90d7-124">Method</span></span> |
| [<span data-ttu-id="f90d7-125">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="f90d7-125">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="f90d7-126">Método</span><span class="sxs-lookup"><span data-stu-id="f90d7-126">Method</span></span> |
| [<span data-ttu-id="f90d7-127">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="f90d7-127">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="f90d7-128">Método</span><span class="sxs-lookup"><span data-stu-id="f90d7-128">Method</span></span> |
| [<span data-ttu-id="f90d7-129">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="f90d7-129">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="f90d7-130">Método</span><span class="sxs-lookup"><span data-stu-id="f90d7-130">Method</span></span> |
| [<span data-ttu-id="f90d7-131">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="f90d7-131">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="f90d7-132">Método</span><span class="sxs-lookup"><span data-stu-id="f90d7-132">Method</span></span> |
| [<span data-ttu-id="f90d7-133">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="f90d7-133">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="f90d7-134">Método</span><span class="sxs-lookup"><span data-stu-id="f90d7-134">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f90d7-135">Namespaces</span><span class="sxs-lookup"><span data-stu-id="f90d7-135">Namespaces</span></span>

<span data-ttu-id="f90d7-136">[diagnostics](Office.context.mailbox.diagnostics.md): Fornece informações de diagnóstico para um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="f90d7-136">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="f90d7-137">[item](Office.context.mailbox.item.md): Fornece propriedades e métodos para acessar uma mensagem ou um compromisso em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="f90d7-137">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="f90d7-138">[userProfile](Office.context.mailbox.userProfile.md): Fornece informações sobre o usuário em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="f90d7-138">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="f90d7-139">Members</span><span class="sxs-lookup"><span data-stu-id="f90d7-139">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="f90d7-140">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="f90d7-140">ewsUrl: String</span></span>

<span data-ttu-id="f90d7-p101">Obtém a URL do ponto de extremidade dos EWS (Serviços Web do Exchange) para esta conta de email. Somente modo de Leitura.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f90d7-143">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="f90d7-143">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f90d7-p102">O valor `ewsUrl` pode ser usado por um serviço remoto para fazer chamadas do EWS à caixa de correio do usuário. Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="f90d7-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="f90d7-146">Tipo</span><span class="sxs-lookup"><span data-stu-id="f90d7-146">Type</span></span>

*   <span data-ttu-id="f90d7-147">String</span><span class="sxs-lookup"><span data-stu-id="f90d7-147">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f90d7-148">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f90d7-148">Requirements</span></span>

|<span data-ttu-id="f90d7-149">Requisito</span><span class="sxs-lookup"><span data-stu-id="f90d7-149">Requirement</span></span>| <span data-ttu-id="f90d7-150">Valor</span><span class="sxs-lookup"><span data-stu-id="f90d7-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="f90d7-151">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f90d7-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f90d7-152">1.0</span><span class="sxs-lookup"><span data-stu-id="f90d7-152">1.0</span></span>|
|[<span data-ttu-id="f90d7-153">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f90d7-153">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f90d7-154">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f90d7-154">ReadItem</span></span>|
|[<span data-ttu-id="f90d7-155">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f90d7-155">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f90d7-156">Read</span><span class="sxs-lookup"><span data-stu-id="f90d7-156">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="f90d7-157">Métodos</span><span class="sxs-lookup"><span data-stu-id="f90d7-157">Methods</span></span>

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-12"></a><span data-ttu-id="f90d7-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="f90d7-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="f90d7-159">Obtém um dicionário contendo informações de hora em tempo local do cliente.</span><span class="sxs-lookup"><span data-stu-id="f90d7-159">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="f90d7-p103">Um aplicativo de email para o Outlook em uma área de trabalho ou na Web pode usar fusos horários diferentes para as datas e horas. O Outlook em uma área de trabalho usa o fuso horário do computador cliente; o Outlook na Web usa o fuso horário definido no Centro de Administração do Exchange (EAC). Você deve lidar com valores de data e hora para que os valores exibidos na interface do usuário sejam sempre consistentes com o fuso horário que o usuário espera.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p103">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="f90d7-p104">Se o aplicativo de email estiver sendo executado no Outlook em um cliente da área de trabalho, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário do computador cliente. Se o aplicativo de email estiver sendo executado no Outlook na Web, o método `convertToLocalClientTime` retornará um objeto de dicionário com os valores definidos para o fuso horário especificado no EAC.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p104">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f90d7-165">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="f90d7-165">Parameters</span></span>

|<span data-ttu-id="f90d7-166">Nome</span><span class="sxs-lookup"><span data-stu-id="f90d7-166">Name</span></span>| <span data-ttu-id="f90d7-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="f90d7-167">Type</span></span>| <span data-ttu-id="f90d7-168">Descrição</span><span class="sxs-lookup"><span data-stu-id="f90d7-168">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="f90d7-169">Date</span><span class="sxs-lookup"><span data-stu-id="f90d7-169">Date</span></span>|<span data-ttu-id="f90d7-170">Um objeto Date</span><span class="sxs-lookup"><span data-stu-id="f90d7-170">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f90d7-171">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f90d7-171">Requirements</span></span>

|<span data-ttu-id="f90d7-172">Requisito</span><span class="sxs-lookup"><span data-stu-id="f90d7-172">Requirement</span></span>| <span data-ttu-id="f90d7-173">Valor</span><span class="sxs-lookup"><span data-stu-id="f90d7-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="f90d7-174">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f90d7-174">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f90d7-175">1.0</span><span class="sxs-lookup"><span data-stu-id="f90d7-175">1.0</span></span>|
|[<span data-ttu-id="f90d7-176">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f90d7-176">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f90d7-177">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f90d7-177">ReadItem</span></span>|
|[<span data-ttu-id="f90d7-178">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f90d7-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f90d7-179">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f90d7-179">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f90d7-180">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f90d7-180">Returns:</span></span>

<span data-ttu-id="f90d7-181">Tipo: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="f90d7-181">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span></span>

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="f90d7-182">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="f90d7-182">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="f90d7-183">Obtém um objeto Date de um dicionário contendo as informações de hora.</span><span class="sxs-lookup"><span data-stu-id="f90d7-183">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="f90d7-184">O método `convertToUtcClientTime` converte um dicionário que contém uma data e hora locais para um objeto Date com os valores corretos para a data e hora locais.</span><span class="sxs-lookup"><span data-stu-id="f90d7-184">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f90d7-185">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="f90d7-185">Parameters</span></span>

|<span data-ttu-id="f90d7-186">Nome</span><span class="sxs-lookup"><span data-stu-id="f90d7-186">Name</span></span>| <span data-ttu-id="f90d7-187">Tipo</span><span class="sxs-lookup"><span data-stu-id="f90d7-187">Type</span></span>| <span data-ttu-id="f90d7-188">Descrição</span><span class="sxs-lookup"><span data-stu-id="f90d7-188">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="f90d7-189">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="f90d7-189">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)|<span data-ttu-id="f90d7-190">O valor de hora local a converter.</span><span class="sxs-lookup"><span data-stu-id="f90d7-190">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f90d7-191">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f90d7-191">Requirements</span></span>

|<span data-ttu-id="f90d7-192">Requisito</span><span class="sxs-lookup"><span data-stu-id="f90d7-192">Requirement</span></span>| <span data-ttu-id="f90d7-193">Valor</span><span class="sxs-lookup"><span data-stu-id="f90d7-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="f90d7-194">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f90d7-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f90d7-195">1.0</span><span class="sxs-lookup"><span data-stu-id="f90d7-195">1.0</span></span>|
|[<span data-ttu-id="f90d7-196">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f90d7-196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f90d7-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f90d7-197">ReadItem</span></span>|
|[<span data-ttu-id="f90d7-198">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f90d7-198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f90d7-199">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f90d7-199">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f90d7-200">Retorna:</span><span class="sxs-lookup"><span data-stu-id="f90d7-200">Returns:</span></span>

<span data-ttu-id="f90d7-201">Um objeto Date com a hora expressa em UTC.</span><span class="sxs-lookup"><span data-stu-id="f90d7-201">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="f90d7-202">Tipo: Data</span><span class="sxs-lookup"><span data-stu-id="f90d7-202">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="f90d7-203">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f90d7-203">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="f90d7-204">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="f90d7-204">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="f90d7-205">Exibe um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="f90d7-205">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f90d7-206">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="f90d7-206">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f90d7-207">O método `displayAppointmentForm` abre um compromisso de calendário existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="f90d7-207">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="f90d7-p105">No Outlook no Mac, você pode usar esse método para exibir um único compromisso que não faz parte de uma série recorrente, ou o compromisso mestre de uma série recorrente, mas não pode exibir um instância da série. Isso ocorre porque no Outlook no Mac você não pode acessar as propriedades (incluindo a ID do item) das instâncias de uma série recorrente.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p105">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="f90d7-210">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32KB.</span><span class="sxs-lookup"><span data-stu-id="f90d7-210">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="f90d7-211">Se o identificador do item especificado não identificar um compromisso existente, um painel em branco abre no dispositivo ou no computador cliente e nenhuma mensagem de erro será exibida.</span><span class="sxs-lookup"><span data-stu-id="f90d7-211">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f90d7-212">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="f90d7-212">Parameters</span></span>

|<span data-ttu-id="f90d7-213">Nome</span><span class="sxs-lookup"><span data-stu-id="f90d7-213">Name</span></span>| <span data-ttu-id="f90d7-214">Tipo</span><span class="sxs-lookup"><span data-stu-id="f90d7-214">Type</span></span>| <span data-ttu-id="f90d7-215">Descrição</span><span class="sxs-lookup"><span data-stu-id="f90d7-215">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f90d7-216">String</span><span class="sxs-lookup"><span data-stu-id="f90d7-216">String</span></span>|<span data-ttu-id="f90d7-217">O identificador dos Serviços Web do Exchange (EWS) para um compromisso de calendário existente.</span><span class="sxs-lookup"><span data-stu-id="f90d7-217">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f90d7-218">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f90d7-218">Requirements</span></span>

|<span data-ttu-id="f90d7-219">Requisito</span><span class="sxs-lookup"><span data-stu-id="f90d7-219">Requirement</span></span>| <span data-ttu-id="f90d7-220">Valor</span><span class="sxs-lookup"><span data-stu-id="f90d7-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="f90d7-221">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f90d7-221">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f90d7-222">1.0</span><span class="sxs-lookup"><span data-stu-id="f90d7-222">1.0</span></span>|
|[<span data-ttu-id="f90d7-223">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f90d7-223">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f90d7-224">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f90d7-224">ReadItem</span></span>|
|[<span data-ttu-id="f90d7-225">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f90d7-225">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f90d7-226">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f90d7-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f90d7-227">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f90d7-227">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="f90d7-228">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="f90d7-228">displayMessageForm(itemId)</span></span>

<span data-ttu-id="f90d7-229">Exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="f90d7-229">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="f90d7-230">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="f90d7-230">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f90d7-231">O método `displayMessageForm` abre uma mensagem existente em uma nova janela na área de trabalho ou em uma caixa de diálogo em dispositivos móveis.</span><span class="sxs-lookup"><span data-stu-id="f90d7-231">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="f90d7-232">No Outlook na Web, este método abre o formulário especificado somente se o corpo do formulário for menor ou igual ao número de caracteres de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f90d7-232">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="f90d7-233">Se o identificador do item especificado não identificar uma mensagem existente, não será exibida mensagem no computador cliente e nenhuma mensagem de erro será retornada.</span><span class="sxs-lookup"><span data-stu-id="f90d7-233">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="f90d7-p106">Não use o método `displayMessageForm` com um `itemId` que representa um compromisso. Use o método `displayAppointmentForm` para exibir um compromisso existente e `displayNewAppointmentForm` para exibir um formulário e criar um novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f90d7-236">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="f90d7-236">Parameters</span></span>

|<span data-ttu-id="f90d7-237">Nome</span><span class="sxs-lookup"><span data-stu-id="f90d7-237">Name</span></span>| <span data-ttu-id="f90d7-238">Tipo</span><span class="sxs-lookup"><span data-stu-id="f90d7-238">Type</span></span>| <span data-ttu-id="f90d7-239">Descrição</span><span class="sxs-lookup"><span data-stu-id="f90d7-239">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f90d7-240">String</span><span class="sxs-lookup"><span data-stu-id="f90d7-240">String</span></span>|<span data-ttu-id="f90d7-241">O identificador dos Serviços Web do Exchange (EWS) para uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="f90d7-241">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f90d7-242">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f90d7-242">Requirements</span></span>

|<span data-ttu-id="f90d7-243">Requisito</span><span class="sxs-lookup"><span data-stu-id="f90d7-243">Requirement</span></span>| <span data-ttu-id="f90d7-244">Valor</span><span class="sxs-lookup"><span data-stu-id="f90d7-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="f90d7-245">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f90d7-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f90d7-246">1.0</span><span class="sxs-lookup"><span data-stu-id="f90d7-246">1.0</span></span>|
|[<span data-ttu-id="f90d7-247">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f90d7-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f90d7-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f90d7-248">ReadItem</span></span>|
|[<span data-ttu-id="f90d7-249">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f90d7-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f90d7-250">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f90d7-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f90d7-251">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f90d7-251">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="f90d7-252">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="f90d7-252">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="f90d7-253">Exibe um formulário para criar um compromisso no calendário.</span><span class="sxs-lookup"><span data-stu-id="f90d7-253">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f90d7-254">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="f90d7-254">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f90d7-p107">O método `displayNewAppointmentForm` abre um formulário que permite ao usuário criar um novo compromisso ou reunião. Se os parâmetros forem especificados, os campos de formulário do compromisso serão preenchidos automaticamente com o conteúdo dos parâmetros.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="f90d7-p108">No Outlook na Web e em dispositivos móveis, este método sempre exibe um formulário com um campo de participantes. Se você não especificar quaisquer participantes como argumentos de entrada, o método exibe um formulário com um botão **Salvar**. Se você especificar participantes, o formulário inclui os participantes e um botão **Enviar**.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p108">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="f90d7-p109">No cliente avançado do Outlook e no Outlook RT, se você especificar quaisquer participantes ou recursos nos parâmetros `requiredAttendees`, `optionalAttendees`ou `resources`, este método exibirá um formulário de reunião com um botão **Enviar**. Se você não especificar destinatários, este método exibirá um formulário de compromisso com um botão **Salvar e Fechar**.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="f90d7-262">Se qualquer dos parâmetros exceder os limites de tamanho especificados, ou se um nome de parâmetro desconhecido for especificado, ocorre uma exceção.</span><span class="sxs-lookup"><span data-stu-id="f90d7-262">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f90d7-263">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="f90d7-263">Parameters</span></span>

|<span data-ttu-id="f90d7-264">Nome</span><span class="sxs-lookup"><span data-stu-id="f90d7-264">Name</span></span>| <span data-ttu-id="f90d7-265">Tipo</span><span class="sxs-lookup"><span data-stu-id="f90d7-265">Type</span></span>| <span data-ttu-id="f90d7-266">Descrição</span><span class="sxs-lookup"><span data-stu-id="f90d7-266">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="f90d7-267">Object</span><span class="sxs-lookup"><span data-stu-id="f90d7-267">Object</span></span> | <span data-ttu-id="f90d7-268">Um dicionário de parâmetros que descreve o novo compromisso.</span><span class="sxs-lookup"><span data-stu-id="f90d7-268">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="f90d7-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="f90d7-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="f90d7-p110">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes obrigatórios do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="f90d7-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="f90d7-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="f90d7-p111">Uma matriz de cadeias de caracteres que contém os endereços de email ou uma matriz contendo um objeto `EmailAddressDetails` para cada um dos participantes opcionais do compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="f90d7-275">Data</span><span class="sxs-lookup"><span data-stu-id="f90d7-275">Date</span></span> | <span data-ttu-id="f90d7-276">Um objeto `Date` que especifica a data e a hora de início do compromisso.</span><span class="sxs-lookup"><span data-stu-id="f90d7-276">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="f90d7-277">Data</span><span class="sxs-lookup"><span data-stu-id="f90d7-277">Date</span></span> | <span data-ttu-id="f90d7-278">Um objeto `Date` que especifica a data e a hora de término do compromisso.</span><span class="sxs-lookup"><span data-stu-id="f90d7-278">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="f90d7-279">String</span><span class="sxs-lookup"><span data-stu-id="f90d7-279">String</span></span> | <span data-ttu-id="f90d7-p112">Uma cadeia de caracteres que contém o local do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="f90d7-282">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="f90d7-282">Array.&lt;String&gt;</span></span> | <span data-ttu-id="f90d7-p113">Uma matriz de cadeias de caracteres que contém os recursos necessários para o compromisso. A matriz está limitada a um máximo de 100 entradas.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="f90d7-285">String</span><span class="sxs-lookup"><span data-stu-id="f90d7-285">String</span></span> | <span data-ttu-id="f90d7-p114">Uma cadeia de caracteres que contém o assunto do compromisso. A cadeia de caracteres está limitada a um máximo de 255 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="f90d7-288">String</span><span class="sxs-lookup"><span data-stu-id="f90d7-288">String</span></span> | <span data-ttu-id="f90d7-p115">O corpo do compromisso. O conteúdo do corpo está limitado a um tamanho máximo de 32 KB.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f90d7-291">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f90d7-291">Requirements</span></span>

|<span data-ttu-id="f90d7-292">Requisito</span><span class="sxs-lookup"><span data-stu-id="f90d7-292">Requirement</span></span>| <span data-ttu-id="f90d7-293">Valor</span><span class="sxs-lookup"><span data-stu-id="f90d7-293">Value</span></span>|
|---|---|
|[<span data-ttu-id="f90d7-294">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f90d7-294">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f90d7-295">1.0</span><span class="sxs-lookup"><span data-stu-id="f90d7-295">1.0</span></span>|
|[<span data-ttu-id="f90d7-296">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f90d7-296">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f90d7-297">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f90d7-297">ReadItem</span></span>|
|[<span data-ttu-id="f90d7-298">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f90d7-298">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f90d7-299">Read</span><span class="sxs-lookup"><span data-stu-id="f90d7-299">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f90d7-300">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f90d7-300">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="f90d7-301">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f90d7-301">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="f90d7-302">Obtém uma cadeia de caracteres que contém um token usado para obter um anexo ou um item de um Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="f90d7-302">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="f90d7-p116">O método `getCallbackTokenAsync` faz uma chamada assíncrona para obter um token opaco do Exchange Server que hospeda a caixa de correio do usuário. A vida útil do token de retorno de chamada é de 5 minutos.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="f90d7-305">Você pode passar o token e um identificador de anexo ou identificador de item para um sistema de terceiros.</span><span class="sxs-lookup"><span data-stu-id="f90d7-305">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="f90d7-306">O sistema de terceiros usa o token como um token de autorização de portador para chamar a operação [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) dos serviços Web do Exchange (EWS) ou a operação [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) para retornar um anexo ou item.</span><span class="sxs-lookup"><span data-stu-id="f90d7-306">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="f90d7-307">Por exemplo, você pode criar um serviço remoto para [obter anexos do item selecionado](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="f90d7-307">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="f90d7-308">Chamar o `getCallbackTokenAsync` método requer um nível de permissão mínimo de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="f90d7-308">Calling the `getCallbackTokenAsync` method requires a minimum permission level of **ReadItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f90d7-309">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="f90d7-309">Parameters</span></span>

|<span data-ttu-id="f90d7-310">Nome</span><span class="sxs-lookup"><span data-stu-id="f90d7-310">Name</span></span>| <span data-ttu-id="f90d7-311">Tipo</span><span class="sxs-lookup"><span data-stu-id="f90d7-311">Type</span></span>| <span data-ttu-id="f90d7-312">Atributos</span><span class="sxs-lookup"><span data-stu-id="f90d7-312">Attributes</span></span>| <span data-ttu-id="f90d7-313">Descrição</span><span class="sxs-lookup"><span data-stu-id="f90d7-313">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f90d7-314">function</span><span class="sxs-lookup"><span data-stu-id="f90d7-314">function</span></span>||<span data-ttu-id="f90d7-315">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f90d7-315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f90d7-316">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f90d7-316">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="f90d7-317">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="f90d7-317">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="f90d7-318">Objeto</span><span class="sxs-lookup"><span data-stu-id="f90d7-318">Object</span></span>| <span data-ttu-id="f90d7-319">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f90d7-319">&lt;optional&gt;</span></span>|<span data-ttu-id="f90d7-320">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="f90d7-320">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f90d7-321">Erros</span><span class="sxs-lookup"><span data-stu-id="f90d7-321">Errors</span></span>

|<span data-ttu-id="f90d7-322">Código de erro</span><span class="sxs-lookup"><span data-stu-id="f90d7-322">Error code</span></span>|<span data-ttu-id="f90d7-323">Descrição</span><span class="sxs-lookup"><span data-stu-id="f90d7-323">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="f90d7-324">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="f90d7-324">The request has failed.</span></span> <span data-ttu-id="f90d7-325">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="f90d7-325">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="f90d7-326">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="f90d7-326">The Exchange server returned an error.</span></span> <span data-ttu-id="f90d7-327">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="f90d7-327">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="f90d7-328">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="f90d7-328">The user is no longer connected to the network.</span></span> <span data-ttu-id="f90d7-329">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="f90d7-329">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f90d7-330">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f90d7-330">Requirements</span></span>

|<span data-ttu-id="f90d7-331">Requisito</span><span class="sxs-lookup"><span data-stu-id="f90d7-331">Requirement</span></span>| <span data-ttu-id="f90d7-332">Valor</span><span class="sxs-lookup"><span data-stu-id="f90d7-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="f90d7-333">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f90d7-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f90d7-334">1.0</span><span class="sxs-lookup"><span data-stu-id="f90d7-334">1.0</span></span>|
|[<span data-ttu-id="f90d7-335">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f90d7-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f90d7-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f90d7-336">ReadItem</span></span>|
|[<span data-ttu-id="f90d7-337">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f90d7-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f90d7-338">Read</span><span class="sxs-lookup"><span data-stu-id="f90d7-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f90d7-339">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f90d7-339">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="f90d7-340">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f90d7-340">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="f90d7-341">Obtém um símbolo que identifica o usuário e o suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="f90d7-341">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="f90d7-342">O método `getUserIdentityTokenAsync` retorna um token que pode ser utilizado para identificar e [autenticar o suplemento e o usuário com um sistema de terceiros](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="f90d7-342">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="f90d7-343">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="f90d7-343">Parameters</span></span>

|<span data-ttu-id="f90d7-344">Nome</span><span class="sxs-lookup"><span data-stu-id="f90d7-344">Name</span></span>| <span data-ttu-id="f90d7-345">Tipo</span><span class="sxs-lookup"><span data-stu-id="f90d7-345">Type</span></span>| <span data-ttu-id="f90d7-346">Atributos</span><span class="sxs-lookup"><span data-stu-id="f90d7-346">Attributes</span></span>| <span data-ttu-id="f90d7-347">Descrição</span><span class="sxs-lookup"><span data-stu-id="f90d7-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f90d7-348">function</span><span class="sxs-lookup"><span data-stu-id="f90d7-348">function</span></span>||<span data-ttu-id="f90d7-349">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f90d7-349">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f90d7-350">O token é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f90d7-350">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="f90d7-351">Se ocorreu um erro, as propriedades`asyncResult.error` e `asyncResult.diagnostics` podem fornecer informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="f90d7-351">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="f90d7-352">Objeto</span><span class="sxs-lookup"><span data-stu-id="f90d7-352">Object</span></span>| <span data-ttu-id="f90d7-353">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f90d7-353">&lt;optional&gt;</span></span>|<span data-ttu-id="f90d7-354">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="f90d7-354">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f90d7-355">Erros</span><span class="sxs-lookup"><span data-stu-id="f90d7-355">Errors</span></span>

|<span data-ttu-id="f90d7-356">Código de erro</span><span class="sxs-lookup"><span data-stu-id="f90d7-356">Error code</span></span>|<span data-ttu-id="f90d7-357">Descrição</span><span class="sxs-lookup"><span data-stu-id="f90d7-357">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="f90d7-358">A solicitação falhou.</span><span class="sxs-lookup"><span data-stu-id="f90d7-358">The request has failed.</span></span> <span data-ttu-id="f90d7-359">Examine o objeto de diagnóstico para o código de erro HTTP.</span><span class="sxs-lookup"><span data-stu-id="f90d7-359">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="f90d7-360">O servidor do Exchange retornou um erro.</span><span class="sxs-lookup"><span data-stu-id="f90d7-360">The Exchange server returned an error.</span></span> <span data-ttu-id="f90d7-361">Para saber mais, confira o objeto de diagnóstico.</span><span class="sxs-lookup"><span data-stu-id="f90d7-361">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="f90d7-362">O usuário não está mais conectado à rede.</span><span class="sxs-lookup"><span data-stu-id="f90d7-362">The user is no longer connected to the network.</span></span> <span data-ttu-id="f90d7-363">Verifique sua conexão de rede e tente novamente.</span><span class="sxs-lookup"><span data-stu-id="f90d7-363">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f90d7-364">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f90d7-364">Requirements</span></span>

|<span data-ttu-id="f90d7-365">Requisito</span><span class="sxs-lookup"><span data-stu-id="f90d7-365">Requirement</span></span>| <span data-ttu-id="f90d7-366">Valor</span><span class="sxs-lookup"><span data-stu-id="f90d7-366">Value</span></span>|
|---|---|
|[<span data-ttu-id="f90d7-367">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f90d7-367">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f90d7-368">1.0</span><span class="sxs-lookup"><span data-stu-id="f90d7-368">1.0</span></span>|
|[<span data-ttu-id="f90d7-369">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f90d7-369">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f90d7-370">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f90d7-370">ReadItem</span></span>|
|[<span data-ttu-id="f90d7-371">Modo Aplicável do Outlook</span><span class="sxs-lookup"><span data-stu-id="f90d7-371">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f90d7-372">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f90d7-372">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f90d7-373">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f90d7-373">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="f90d7-374">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f90d7-374">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="f90d7-375">Faz uma solicitação assíncrona em um serviço dos EWS (Serviços Web do Exchange) no servidor Exchange que hospeda a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="f90d7-375">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="f90d7-376">Esse método não tem suporte nas seguintes situações.</span><span class="sxs-lookup"><span data-stu-id="f90d7-376">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="f90d7-377">No Outlook no iOS ou no Android</span><span class="sxs-lookup"><span data-stu-id="f90d7-377">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="f90d7-378">Quando o suplemento é carregado em uma caixa de correio do Gmail</span><span class="sxs-lookup"><span data-stu-id="f90d7-378">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="f90d7-379">Nesses casos, os suplementos devem [usar as APIs REST](/outlook/add-ins/use-rest-api) para acessar a caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="f90d7-379">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="f90d7-380">O método `makeEwsRequestAsync` envia uma solicitação do EWS em nome do suplemento ao Exchange.</span><span class="sxs-lookup"><span data-stu-id="f90d7-380">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="f90d7-381">Consulte [Chamar serviços Web de um suplemento do Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) para obter uma lista das operações de EWS compatíveis.</span><span class="sxs-lookup"><span data-stu-id="f90d7-381">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="f90d7-382">Não é possível solicitar os itens associados da pasta com o método `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="f90d7-382">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="f90d7-383">A solicitação XML deve especificar a codificação UTF-8.</span><span class="sxs-lookup"><span data-stu-id="f90d7-383">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="f90d7-p125">O suplemento deve ter a permissão **ReadWriteMailbox** para usar o método `makeEwsRequestAsync`. Para saber mais sobre como usar a permissão **ReadWriteMailbox** e as operações do EWS que você pode chamar com o método `makeEwsRequestAsync`, confira [Especificar permissões para acesso de suplemento de email na caixa de correio do usuário](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="f90d7-p125">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="f90d7-386">O administrador do servidor deve definir `OAuthAuthentication` como true no diretório do EWS para o Servidor de Acesso para Cliente a fim de habilitar o método `makeEwsRequestAsync` a realizar solicitações do EWS.</span><span class="sxs-lookup"><span data-stu-id="f90d7-386">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="f90d7-387">Diferenças de versão</span><span class="sxs-lookup"><span data-stu-id="f90d7-387">Version differences</span></span>

<span data-ttu-id="f90d7-388">Ao usar o método `makeEwsRequestAsync` nos aplicativos de email em execução em versões do Outlook anteriores à 15.0.4535.1004, é preciso definir o valor de codificação como `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="f90d7-388">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="f90d7-p126">Não é necessário definir o valor de codificação quando o aplicativo de email estiver em execução no Outlook na Web. Você pode determinar se o aplicativo de email está em execução no Outlook ou no Outlook na Web usando a propriedade mailbox.diagnostics.hostName. Você pode determinar que versão do Outlook está em execução usando a propriedade mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="f90d7-p126">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f90d7-392">Parâmetros</span><span class="sxs-lookup"><span data-stu-id="f90d7-392">Parameters</span></span>

|<span data-ttu-id="f90d7-393">Nome</span><span class="sxs-lookup"><span data-stu-id="f90d7-393">Name</span></span>| <span data-ttu-id="f90d7-394">Tipo</span><span class="sxs-lookup"><span data-stu-id="f90d7-394">Type</span></span>| <span data-ttu-id="f90d7-395">Atributos</span><span class="sxs-lookup"><span data-stu-id="f90d7-395">Attributes</span></span>| <span data-ttu-id="f90d7-396">Descrição</span><span class="sxs-lookup"><span data-stu-id="f90d7-396">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="f90d7-397">String</span><span class="sxs-lookup"><span data-stu-id="f90d7-397">String</span></span>||<span data-ttu-id="f90d7-398">A solicitação do EWS.</span><span class="sxs-lookup"><span data-stu-id="f90d7-398">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="f90d7-399">function</span><span class="sxs-lookup"><span data-stu-id="f90d7-399">function</span></span>||<span data-ttu-id="f90d7-400">Quando o método for concluído, a função passada ao parâmetro `callback` é chamada com um único parâmetro, `asyncResult`, que é um objeto [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f90d7-400">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f90d7-401">O resultado XML da chamada do EWS é fornecido como uma cadeia de caracteres na propriedade `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f90d7-401">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="f90d7-402">Se o resultado exceder 1 MB de tamanho, será exibida uma mensagem de erro.</span><span class="sxs-lookup"><span data-stu-id="f90d7-402">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="f90d7-403">Objeto</span><span class="sxs-lookup"><span data-stu-id="f90d7-403">Object</span></span>| <span data-ttu-id="f90d7-404">&lt;opcional&gt;</span><span class="sxs-lookup"><span data-stu-id="f90d7-404">&lt;optional&gt;</span></span>|<span data-ttu-id="f90d7-405">Quaisquer dados de estado que são passados ao método assíncrono.</span><span class="sxs-lookup"><span data-stu-id="f90d7-405">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f90d7-406">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f90d7-406">Requirements</span></span>

|<span data-ttu-id="f90d7-407">Requisito</span><span class="sxs-lookup"><span data-stu-id="f90d7-407">Requirement</span></span>| <span data-ttu-id="f90d7-408">Valor</span><span class="sxs-lookup"><span data-stu-id="f90d7-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="f90d7-409">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f90d7-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f90d7-410">1.0</span><span class="sxs-lookup"><span data-stu-id="f90d7-410">1.0</span></span>|
|[<span data-ttu-id="f90d7-411">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f90d7-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f90d7-412">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="f90d7-412">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="f90d7-413">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f90d7-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f90d7-414">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f90d7-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f90d7-415">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f90d7-415">Example</span></span>

<span data-ttu-id="f90d7-416">O exemplo a seguir chama `makeEwsRequestAsync` para usar a operação `GetItem` para obter o assunto de um item.</span><span class="sxs-lookup"><span data-stu-id="f90d7-416">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
