---
title: Office.context - conjunto de requisitos 1.9
description: Office. Membros do objeto Context disponíveis para Outlook de entrada usando o conjunto de requisitos da API de Caixa de Correio 1.9.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: f45eec7ce638f4bbb97ad4be9f2ba089905c631d
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590516"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="f7f40-103">context (Conjunto de requisitos de caixa de correio 1.9)</span><span class="sxs-lookup"><span data-stu-id="f7f40-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="f7f40-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="f7f40-104">[Office](office.md).context</span></span>

<span data-ttu-id="f7f40-105">Office.context fornece interfaces compartilhadas que são usadas por complementos em todos os Office aplicativos.</span><span class="sxs-lookup"><span data-stu-id="f7f40-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="f7f40-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="f7f40-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f7f40-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7f40-107">Requirements</span></span>

|<span data-ttu-id="f7f40-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7f40-108">Requirement</span></span>| <span data-ttu-id="f7f40-109">Valor</span><span class="sxs-lookup"><span data-stu-id="f7f40-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7f40-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7f40-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7f40-111">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-111">1.1</span></span>|
|[<span data-ttu-id="f7f40-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7f40-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7f40-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="f7f40-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="f7f40-114">Properties</span></span>

| <span data-ttu-id="f7f40-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="f7f40-115">Property</span></span> | <span data-ttu-id="f7f40-116">Modos</span><span class="sxs-lookup"><span data-stu-id="f7f40-116">Modes</span></span> | <span data-ttu-id="f7f40-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="f7f40-117">Return type</span></span> | <span data-ttu-id="f7f40-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="f7f40-118">Minimum</span></span><br><span data-ttu-id="f7f40-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="f7f40-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f7f40-120">auth</span><span class="sxs-lookup"><span data-stu-id="f7f40-120">auth</span></span>](#auth-auth) | <span data-ttu-id="f7f40-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="f7f40-121">Compose</span></span><br><span data-ttu-id="f7f40-122">Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-122">Read</span></span> | [<span data-ttu-id="f7f40-123">Auth</span><span class="sxs-lookup"><span data-stu-id="f7f40-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f7f40-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="f7f40-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="f7f40-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="f7f40-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="f7f40-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="f7f40-126">Compose</span></span><br><span data-ttu-id="f7f40-127">Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-127">Read</span></span> | <span data-ttu-id="f7f40-128">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f7f40-128">String</span></span> | [<span data-ttu-id="f7f40-129">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f7f40-130">diagnostics</span><span class="sxs-lookup"><span data-stu-id="f7f40-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="f7f40-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="f7f40-131">Compose</span></span><br><span data-ttu-id="f7f40-132">Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-132">Read</span></span> | [<span data-ttu-id="f7f40-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="f7f40-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f7f40-134">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f7f40-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="f7f40-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="f7f40-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="f7f40-136">Compose</span></span><br><span data-ttu-id="f7f40-137">Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-137">Read</span></span> | <span data-ttu-id="f7f40-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f7f40-138">String</span></span> | [<span data-ttu-id="f7f40-139">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f7f40-140">host</span><span class="sxs-lookup"><span data-stu-id="f7f40-140">host</span></span>](#host-hosttype) | <span data-ttu-id="f7f40-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="f7f40-141">Compose</span></span><br><span data-ttu-id="f7f40-142">Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-142">Read</span></span> | [<span data-ttu-id="f7f40-143">HostType</span><span class="sxs-lookup"><span data-stu-id="f7f40-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f7f40-144">1.5</span><span class="sxs-lookup"><span data-stu-id="f7f40-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="f7f40-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="f7f40-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="f7f40-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="f7f40-146">Compose</span></span><br><span data-ttu-id="f7f40-147">Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-147">Read</span></span> | [<span data-ttu-id="f7f40-148">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="f7f40-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f7f40-149">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f7f40-150">platform</span><span class="sxs-lookup"><span data-stu-id="f7f40-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="f7f40-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="f7f40-151">Compose</span></span><br><span data-ttu-id="f7f40-152">Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-152">Read</span></span> | [<span data-ttu-id="f7f40-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="f7f40-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f7f40-154">1.5</span><span class="sxs-lookup"><span data-stu-id="f7f40-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="f7f40-155">requirements</span><span class="sxs-lookup"><span data-stu-id="f7f40-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="f7f40-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="f7f40-156">Compose</span></span><br><span data-ttu-id="f7f40-157">Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-157">Read</span></span> | [<span data-ttu-id="f7f40-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="f7f40-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f7f40-159">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f7f40-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="f7f40-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="f7f40-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="f7f40-161">Compose</span></span><br><span data-ttu-id="f7f40-162">Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-162">Read</span></span> | [<span data-ttu-id="f7f40-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="f7f40-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f7f40-164">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f7f40-165">ui</span><span class="sxs-lookup"><span data-stu-id="f7f40-165">ui</span></span>](#ui-ui) | <span data-ttu-id="f7f40-166">Escrever</span><span class="sxs-lookup"><span data-stu-id="f7f40-166">Compose</span></span><br><span data-ttu-id="f7f40-167">Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-167">Read</span></span> | [<span data-ttu-id="f7f40-168">UI</span><span class="sxs-lookup"><span data-stu-id="f7f40-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f7f40-169">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="f7f40-170">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="f7f40-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="f7f40-171">auth: [Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="f7f40-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="f7f40-172">Oferece suporte a [SSO (login único)](../../../outlook/authenticate-a-user-with-an-sso-token.md) fornecendo um método que permite ao aplicativo Office obter um token de acesso ao aplicativo Web do complemento.</span><span class="sxs-lookup"><span data-stu-id="f7f40-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="f7f40-173">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="f7f40-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="f7f40-174">Consulte [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="f7f40-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="f7f40-175">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7f40-175">Type</span></span>

*   [<span data-ttu-id="f7f40-176">Auth</span><span class="sxs-lookup"><span data-stu-id="f7f40-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="f7f40-177">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7f40-177">Requirements</span></span>

|<span data-ttu-id="f7f40-178">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7f40-178">Requirement</span></span>| <span data-ttu-id="f7f40-179">Valor</span><span class="sxs-lookup"><span data-stu-id="f7f40-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7f40-180">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7f40-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7f40-181">N/D</span><span class="sxs-lookup"><span data-stu-id="f7f40-181">N/A</span></span>|
|[<span data-ttu-id="f7f40-182">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7f40-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7f40-183">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f7f40-184">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f7f40-184">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="f7f40-185">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="f7f40-185">contentLanguage: String</span></span>

<span data-ttu-id="f7f40-186">Obtém a localidade (idioma) especificada pelo usuário para editar o item.</span><span class="sxs-lookup"><span data-stu-id="f7f40-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="f7f40-187">O `contentLanguage` valor reflete a **configuração** atual de Idioma de Edição especificada com opções de > de arquivo **> idioma** no aplicativo Office cliente.</span><span class="sxs-lookup"><span data-stu-id="f7f40-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="f7f40-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7f40-188">Type</span></span>

*   <span data-ttu-id="f7f40-189">String</span><span class="sxs-lookup"><span data-stu-id="f7f40-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f7f40-190">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7f40-190">Requirements</span></span>

|<span data-ttu-id="f7f40-191">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7f40-191">Requirement</span></span>| <span data-ttu-id="f7f40-192">Valor</span><span class="sxs-lookup"><span data-stu-id="f7f40-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7f40-193">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7f40-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7f40-194">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-194">1.1</span></span>|
|[<span data-ttu-id="f7f40-195">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7f40-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7f40-196">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f7f40-197">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f7f40-197">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="f7f40-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="f7f40-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="f7f40-199">Obtém informações sobre o ambiente no qual o complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="f7f40-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="f7f40-200">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7f40-200">Type</span></span>

*   [<span data-ttu-id="f7f40-201">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="f7f40-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="f7f40-202">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7f40-202">Requirements</span></span>

|<span data-ttu-id="f7f40-203">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7f40-203">Requirement</span></span>| <span data-ttu-id="f7f40-204">Valor</span><span class="sxs-lookup"><span data-stu-id="f7f40-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7f40-205">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7f40-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7f40-206">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-206">1.1</span></span>|
|[<span data-ttu-id="f7f40-207">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7f40-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7f40-208">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f7f40-209">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f7f40-209">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="f7f40-210">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="f7f40-210">displayLanguage: String</span></span>

<span data-ttu-id="f7f40-211">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente Office cliente.</span><span class="sxs-lookup"><span data-stu-id="f7f40-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="f7f40-212">O valor reflete a configuração atual de Idioma de Exibição especificada com Opções > > Idioma no aplicativo Office `displayLanguage` cliente.  </span><span class="sxs-lookup"><span data-stu-id="f7f40-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="f7f40-213">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7f40-213">Type</span></span>

*   <span data-ttu-id="f7f40-214">String</span><span class="sxs-lookup"><span data-stu-id="f7f40-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f7f40-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7f40-215">Requirements</span></span>

|<span data-ttu-id="f7f40-216">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7f40-216">Requirement</span></span>| <span data-ttu-id="f7f40-217">Valor</span><span class="sxs-lookup"><span data-stu-id="f7f40-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7f40-218">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7f40-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7f40-219">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-219">1.1</span></span>|
|[<span data-ttu-id="f7f40-220">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7f40-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7f40-221">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f7f40-222">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f7f40-222">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="f7f40-223">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="f7f40-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="f7f40-224">Obtém o Office aplicativo que está hospedando o complemento.</span><span class="sxs-lookup"><span data-stu-id="f7f40-224">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="f7f40-225">Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter a plataforma.</span><span class="sxs-lookup"><span data-stu-id="f7f40-225">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="f7f40-226">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7f40-226">Type</span></span>

*   [<span data-ttu-id="f7f40-227">HostType</span><span class="sxs-lookup"><span data-stu-id="f7f40-227">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="f7f40-228">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7f40-228">Requirements</span></span>

|<span data-ttu-id="f7f40-229">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7f40-229">Requirement</span></span>| <span data-ttu-id="f7f40-230">Valor</span><span class="sxs-lookup"><span data-stu-id="f7f40-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7f40-231">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7f40-231">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7f40-232">1,5</span><span class="sxs-lookup"><span data-stu-id="f7f40-232">1.5</span></span>|
|[<span data-ttu-id="f7f40-233">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7f40-233">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7f40-234">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-234">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f7f40-235">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f7f40-235">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="f7f40-236">plataforma: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="f7f40-236">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="f7f40-237">Fornece a plataforma na qual o complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="f7f40-237">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="f7f40-238">Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter a plataforma.</span><span class="sxs-lookup"><span data-stu-id="f7f40-238">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="f7f40-239">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7f40-239">Type</span></span>

*   [<span data-ttu-id="f7f40-240">PlatformType</span><span class="sxs-lookup"><span data-stu-id="f7f40-240">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="f7f40-241">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7f40-241">Requirements</span></span>

|<span data-ttu-id="f7f40-242">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7f40-242">Requirement</span></span>| <span data-ttu-id="f7f40-243">Valor</span><span class="sxs-lookup"><span data-stu-id="f7f40-243">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7f40-244">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7f40-244">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7f40-245">1,5</span><span class="sxs-lookup"><span data-stu-id="f7f40-245">1.5</span></span>|
|[<span data-ttu-id="f7f40-246">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7f40-246">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7f40-247">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-247">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f7f40-248">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f7f40-248">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="f7f40-249">requirements: [RequirementsSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="f7f40-249">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="f7f40-250">Fornece um método para determinar quais conjuntos de requisitos são suportados no aplicativo e na plataforma atual.</span><span class="sxs-lookup"><span data-stu-id="f7f40-250">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="f7f40-251">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7f40-251">Type</span></span>

*   [<span data-ttu-id="f7f40-252">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="f7f40-252">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="f7f40-253">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7f40-253">Requirements</span></span>

|<span data-ttu-id="f7f40-254">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7f40-254">Requirement</span></span>| <span data-ttu-id="f7f40-255">Valor</span><span class="sxs-lookup"><span data-stu-id="f7f40-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7f40-256">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7f40-256">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7f40-257">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-257">1.1</span></span>|
|[<span data-ttu-id="f7f40-258">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7f40-258">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7f40-259">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f7f40-260">Exemplo</span><span class="sxs-lookup"><span data-stu-id="f7f40-260">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="f7f40-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="f7f40-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="f7f40-262">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="f7f40-262">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="f7f40-263">O objeto permite que você armazene e acesse dados para um complemento de email armazenado na caixa de correio de um usuário, de modo que está disponível para esse complemento quando ele está sendo executado de qualquer cliente Outlook usado para acessar essa caixa de `RoamingSettings` correio.</span><span class="sxs-lookup"><span data-stu-id="f7f40-263">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="f7f40-264">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7f40-264">Type</span></span>

*   [<span data-ttu-id="f7f40-265">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="f7f40-265">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="f7f40-266">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7f40-266">Requirements</span></span>

|<span data-ttu-id="f7f40-267">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7f40-267">Requirement</span></span>| <span data-ttu-id="f7f40-268">Valor</span><span class="sxs-lookup"><span data-stu-id="f7f40-268">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7f40-269">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7f40-269">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7f40-270">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-270">1.1</span></span>|
|[<span data-ttu-id="f7f40-271">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="f7f40-271">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="f7f40-272">Restrito</span><span class="sxs-lookup"><span data-stu-id="f7f40-272">Restricted</span></span>|
|[<span data-ttu-id="f7f40-273">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7f40-273">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7f40-274">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-274">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="f7f40-275">interface do usuário: [interface do usuário](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="f7f40-275">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="f7f40-276">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus Office de usuário.</span><span class="sxs-lookup"><span data-stu-id="f7f40-276">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="f7f40-277">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7f40-277">Type</span></span>

*   [<span data-ttu-id="f7f40-278">UI</span><span class="sxs-lookup"><span data-stu-id="f7f40-278">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="f7f40-279">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7f40-279">Requirements</span></span>

|<span data-ttu-id="f7f40-280">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7f40-280">Requirement</span></span>| <span data-ttu-id="f7f40-281">Valor</span><span class="sxs-lookup"><span data-stu-id="f7f40-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7f40-282">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7f40-282">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7f40-283">1.1</span><span class="sxs-lookup"><span data-stu-id="f7f40-283">1.1</span></span>|
|[<span data-ttu-id="f7f40-284">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7f40-284">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7f40-285">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7f40-285">Compose or Read</span></span>|
