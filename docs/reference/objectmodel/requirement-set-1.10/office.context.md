---
title: Office.context - conjunto de requisitos 1.10
description: Office. Membros do objeto context disponível para Outlook de entrada usando o conjunto de requisitos da API de Caixa de Correio 1.10.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: cb189dc3b7b51357dee8ac83bc61795b3ec47ae5
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592024"
---
# <a name="context-mailbox-requirement-set-110"></a><span data-ttu-id="dfc1f-103">context (Conjunto de requisitos de caixa de correio 1.10)</span><span class="sxs-lookup"><span data-stu-id="dfc1f-103">context (Mailbox requirement set 1.10)</span></span>

### <a name="officecontext"></a><span data-ttu-id="dfc1f-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="dfc1f-104">[Office](office.md).context</span></span>

<span data-ttu-id="dfc1f-105">Office.context fornece interfaces compartilhadas que são usadas por complementos em todos os Office aplicativos.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="dfc1f-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="dfc1f-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dfc1f-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfc1f-107">Requirements</span></span>

|<span data-ttu-id="dfc1f-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfc1f-108">Requirement</span></span>| <span data-ttu-id="dfc1f-109">Valor</span><span class="sxs-lookup"><span data-stu-id="dfc1f-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfc1f-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfc1f-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfc1f-111">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-111">1.1</span></span>|
|[<span data-ttu-id="dfc1f-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfc1f-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfc1f-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="dfc1f-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="dfc1f-114">Properties</span></span>

| <span data-ttu-id="dfc1f-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="dfc1f-115">Property</span></span> | <span data-ttu-id="dfc1f-116">Modos</span><span class="sxs-lookup"><span data-stu-id="dfc1f-116">Modes</span></span> | <span data-ttu-id="dfc1f-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="dfc1f-117">Return type</span></span> | <span data-ttu-id="dfc1f-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="dfc1f-118">Minimum</span></span><br><span data-ttu-id="dfc1f-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="dfc1f-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dfc1f-120">auth</span><span class="sxs-lookup"><span data-stu-id="dfc1f-120">auth</span></span>](#auth-auth) | <span data-ttu-id="dfc1f-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfc1f-121">Compose</span></span><br><span data-ttu-id="dfc1f-122">Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-122">Read</span></span> | [<span data-ttu-id="dfc1f-123">Auth</span><span class="sxs-lookup"><span data-stu-id="dfc1f-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="dfc1f-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="dfc1f-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="dfc1f-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="dfc1f-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="dfc1f-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfc1f-126">Compose</span></span><br><span data-ttu-id="dfc1f-127">Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-127">Read</span></span> | <span data-ttu-id="dfc1f-128">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dfc1f-128">String</span></span> | [<span data-ttu-id="dfc1f-129">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfc1f-130">diagnostics</span><span class="sxs-lookup"><span data-stu-id="dfc1f-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="dfc1f-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfc1f-131">Compose</span></span><br><span data-ttu-id="dfc1f-132">Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-132">Read</span></span> | [<span data-ttu-id="dfc1f-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="dfc1f-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="dfc1f-134">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfc1f-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="dfc1f-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="dfc1f-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfc1f-136">Compose</span></span><br><span data-ttu-id="dfc1f-137">Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-137">Read</span></span> | <span data-ttu-id="dfc1f-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dfc1f-138">String</span></span> | [<span data-ttu-id="dfc1f-139">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfc1f-140">host</span><span class="sxs-lookup"><span data-stu-id="dfc1f-140">host</span></span>](#host-hosttype) | <span data-ttu-id="dfc1f-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfc1f-141">Compose</span></span><br><span data-ttu-id="dfc1f-142">Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-142">Read</span></span> | [<span data-ttu-id="dfc1f-143">HostType</span><span class="sxs-lookup"><span data-stu-id="dfc1f-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="dfc1f-144">1.5</span><span class="sxs-lookup"><span data-stu-id="dfc1f-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="dfc1f-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="dfc1f-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="dfc1f-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfc1f-146">Compose</span></span><br><span data-ttu-id="dfc1f-147">Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-147">Read</span></span> | [<span data-ttu-id="dfc1f-148">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="dfc1f-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="dfc1f-149">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfc1f-150">platform</span><span class="sxs-lookup"><span data-stu-id="dfc1f-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="dfc1f-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfc1f-151">Compose</span></span><br><span data-ttu-id="dfc1f-152">Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-152">Read</span></span> | [<span data-ttu-id="dfc1f-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="dfc1f-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="dfc1f-154">1.5</span><span class="sxs-lookup"><span data-stu-id="dfc1f-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="dfc1f-155">requirements</span><span class="sxs-lookup"><span data-stu-id="dfc1f-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="dfc1f-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfc1f-156">Compose</span></span><br><span data-ttu-id="dfc1f-157">Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-157">Read</span></span> | [<span data-ttu-id="dfc1f-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="dfc1f-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="dfc1f-159">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfc1f-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="dfc1f-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="dfc1f-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfc1f-161">Compose</span></span><br><span data-ttu-id="dfc1f-162">Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-162">Read</span></span> | [<span data-ttu-id="dfc1f-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="dfc1f-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="dfc1f-164">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfc1f-165">ui</span><span class="sxs-lookup"><span data-stu-id="dfc1f-165">ui</span></span>](#ui-ui) | <span data-ttu-id="dfc1f-166">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfc1f-166">Compose</span></span><br><span data-ttu-id="dfc1f-167">Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-167">Read</span></span> | [<span data-ttu-id="dfc1f-168">UI</span><span class="sxs-lookup"><span data-stu-id="dfc1f-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="dfc1f-169">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="dfc1f-170">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="dfc1f-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="dfc1f-171">auth: [Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="dfc1f-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="dfc1f-172">Oferece suporte a [SSO (login único)](../../../outlook/authenticate-a-user-with-an-sso-token.md) fornecendo um método que permite ao aplicativo Office obter um token de acesso ao aplicativo Web do complemento.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="dfc1f-173">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="dfc1f-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-174">Type</span></span>

*   [<span data-ttu-id="dfc1f-175">Auth</span><span class="sxs-lookup"><span data-stu-id="dfc1f-175">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="dfc1f-176">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfc1f-176">Requirements</span></span>

|<span data-ttu-id="dfc1f-177">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfc1f-177">Requirement</span></span>| <span data-ttu-id="dfc1f-178">Valor</span><span class="sxs-lookup"><span data-stu-id="dfc1f-178">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfc1f-179">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfc1f-179">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfc1f-180">1.10</span><span class="sxs-lookup"><span data-stu-id="dfc1f-180">1.10</span></span>|
|[<span data-ttu-id="dfc1f-181">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfc1f-181">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfc1f-182">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-182">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfc1f-183">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-183">Example</span></span>

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

#### <a name="contentlanguage-string"></a><span data-ttu-id="dfc1f-184">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="dfc1f-184">contentLanguage: String</span></span>

<span data-ttu-id="dfc1f-185">Obtém a localidade (idioma) especificada pelo usuário para editar o item.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-185">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="dfc1f-186">O `contentLanguage` valor reflete a **configuração** atual de Idioma de Edição especificada com opções de > de arquivo **> idioma** no aplicativo Office cliente.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-186">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="dfc1f-187">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-187">Type</span></span>

*   <span data-ttu-id="dfc1f-188">String</span><span class="sxs-lookup"><span data-stu-id="dfc1f-188">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dfc1f-189">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfc1f-189">Requirements</span></span>

|<span data-ttu-id="dfc1f-190">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfc1f-190">Requirement</span></span>| <span data-ttu-id="dfc1f-191">Valor</span><span class="sxs-lookup"><span data-stu-id="dfc1f-191">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfc1f-192">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfc1f-192">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfc1f-193">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-193">1.1</span></span>|
|[<span data-ttu-id="dfc1f-194">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfc1f-194">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfc1f-195">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-195">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfc1f-196">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-196">Example</span></span>

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="dfc1f-197">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="dfc1f-197">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="dfc1f-198">Obtém informações sobre o ambiente no qual o complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-198">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="dfc1f-199">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-199">Type</span></span>

*   [<span data-ttu-id="dfc1f-200">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="dfc1f-200">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="dfc1f-201">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfc1f-201">Requirements</span></span>

|<span data-ttu-id="dfc1f-202">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfc1f-202">Requirement</span></span>| <span data-ttu-id="dfc1f-203">Valor</span><span class="sxs-lookup"><span data-stu-id="dfc1f-203">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfc1f-204">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfc1f-204">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfc1f-205">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-205">1.1</span></span>|
|[<span data-ttu-id="dfc1f-206">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfc1f-206">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfc1f-207">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-207">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfc1f-208">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-208">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="dfc1f-209">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="dfc1f-209">displayLanguage: String</span></span>

<span data-ttu-id="dfc1f-210">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente Office cliente.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-210">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="dfc1f-211">O valor reflete a configuração atual de Idioma de Exibição especificada com Opções > > Idioma no aplicativo Office `displayLanguage` cliente.  </span><span class="sxs-lookup"><span data-stu-id="dfc1f-211">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="dfc1f-212">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-212">Type</span></span>

*   <span data-ttu-id="dfc1f-213">String</span><span class="sxs-lookup"><span data-stu-id="dfc1f-213">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dfc1f-214">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfc1f-214">Requirements</span></span>

|<span data-ttu-id="dfc1f-215">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfc1f-215">Requirement</span></span>| <span data-ttu-id="dfc1f-216">Valor</span><span class="sxs-lookup"><span data-stu-id="dfc1f-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfc1f-217">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfc1f-217">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfc1f-218">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-218">1.1</span></span>|
|[<span data-ttu-id="dfc1f-219">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfc1f-219">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfc1f-220">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfc1f-221">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-221">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="host-hosttype"></a><span data-ttu-id="dfc1f-222">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="dfc1f-222">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="dfc1f-223">Obtém o Office aplicativo que está hospedando o complemento.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-223">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="dfc1f-224">Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter o host.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-224">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="dfc1f-225">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-225">Type</span></span>

*   [<span data-ttu-id="dfc1f-226">HostType</span><span class="sxs-lookup"><span data-stu-id="dfc1f-226">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="dfc1f-227">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfc1f-227">Requirements</span></span>

|<span data-ttu-id="dfc1f-228">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfc1f-228">Requirement</span></span>| <span data-ttu-id="dfc1f-229">Valor</span><span class="sxs-lookup"><span data-stu-id="dfc1f-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfc1f-230">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfc1f-230">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfc1f-231">1,5</span><span class="sxs-lookup"><span data-stu-id="dfc1f-231">1.5</span></span>|
|[<span data-ttu-id="dfc1f-232">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfc1f-232">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfc1f-233">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfc1f-234">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-234">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="dfc1f-235">plataforma: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="dfc1f-235">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="dfc1f-236">Fornece a plataforma na qual o complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-236">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="dfc1f-237">Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter a plataforma.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-237">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="dfc1f-238">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-238">Type</span></span>

*   [<span data-ttu-id="dfc1f-239">PlatformType</span><span class="sxs-lookup"><span data-stu-id="dfc1f-239">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="dfc1f-240">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfc1f-240">Requirements</span></span>

|<span data-ttu-id="dfc1f-241">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfc1f-241">Requirement</span></span>| <span data-ttu-id="dfc1f-242">Valor</span><span class="sxs-lookup"><span data-stu-id="dfc1f-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfc1f-243">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfc1f-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfc1f-244">1,5</span><span class="sxs-lookup"><span data-stu-id="dfc1f-244">1.5</span></span>|
|[<span data-ttu-id="dfc1f-245">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfc1f-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfc1f-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfc1f-247">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-247">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="dfc1f-248">requirements: [RequirementsSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="dfc1f-248">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="dfc1f-249">Fornece um método para determinar quais conjuntos de requisitos são suportados no aplicativo e na plataforma atual.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-249">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="dfc1f-250">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-250">Type</span></span>

*   [<span data-ttu-id="dfc1f-251">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="dfc1f-251">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="dfc1f-252">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfc1f-252">Requirements</span></span>

|<span data-ttu-id="dfc1f-253">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfc1f-253">Requirement</span></span>| <span data-ttu-id="dfc1f-254">Valor</span><span class="sxs-lookup"><span data-stu-id="dfc1f-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfc1f-255">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfc1f-255">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfc1f-256">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-256">1.1</span></span>|
|[<span data-ttu-id="dfc1f-257">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfc1f-257">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfc1f-258">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfc1f-259">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-259">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="dfc1f-260">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="dfc1f-260">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="dfc1f-261">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-261">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="dfc1f-262">O objeto permite que você armazene e acesse dados para um complemento de email armazenado na caixa de correio de um usuário, de modo que está disponível para esse complemento quando ele está sendo executado de qualquer cliente Outlook usado para acessar essa caixa de `RoamingSettings` correio.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-262">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="dfc1f-263">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-263">Type</span></span>

*   [<span data-ttu-id="dfc1f-264">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="dfc1f-264">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="dfc1f-265">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfc1f-265">Requirements</span></span>

|<span data-ttu-id="dfc1f-266">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfc1f-266">Requirement</span></span>| <span data-ttu-id="dfc1f-267">Valor</span><span class="sxs-lookup"><span data-stu-id="dfc1f-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfc1f-268">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfc1f-268">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfc1f-269">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-269">1.1</span></span>|
|[<span data-ttu-id="dfc1f-270">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-270">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="dfc1f-271">Restrito</span><span class="sxs-lookup"><span data-stu-id="dfc1f-271">Restricted</span></span>|
|[<span data-ttu-id="dfc1f-272">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfc1f-272">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfc1f-273">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="dfc1f-274">interface do usuário: [interface do usuário](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="dfc1f-274">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="dfc1f-275">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus Office de usuário.</span><span class="sxs-lookup"><span data-stu-id="dfc1f-275">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="dfc1f-276">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfc1f-276">Type</span></span>

*   [<span data-ttu-id="dfc1f-277">UI</span><span class="sxs-lookup"><span data-stu-id="dfc1f-277">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="dfc1f-278">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfc1f-278">Requirements</span></span>

|<span data-ttu-id="dfc1f-279">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfc1f-279">Requirement</span></span>| <span data-ttu-id="dfc1f-280">Valor</span><span class="sxs-lookup"><span data-stu-id="dfc1f-280">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfc1f-281">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfc1f-281">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfc1f-282">1.1</span><span class="sxs-lookup"><span data-stu-id="dfc1f-282">1.1</span></span>|
|[<span data-ttu-id="dfc1f-283">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfc1f-283">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfc1f-284">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfc1f-284">Compose or Read</span></span>|
