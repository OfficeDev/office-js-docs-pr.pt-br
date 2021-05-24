---
title: Office.context - conjunto de requisitos de visualização
description: Office. Membros do objeto Context disponíveis para Outlook de usuário usando conjunto de requisitos de visualização da API de Caixa de Correio.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 59b1cce579afe69384e41a6f31cc70c8cec25bea
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591069"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="9b18e-103">context (Conjunto de requisitos de visualização de caixa de correio)</span><span class="sxs-lookup"><span data-stu-id="9b18e-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="9b18e-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="9b18e-104">[Office](office.md).context</span></span>

<span data-ttu-id="9b18e-105">Office.context fornece interfaces compartilhadas que são usadas por complementos em todos os Office aplicativos.</span><span class="sxs-lookup"><span data-stu-id="9b18e-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="9b18e-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="9b18e-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b18e-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9b18e-107">Requirements</span></span>

|<span data-ttu-id="9b18e-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="9b18e-108">Requirement</span></span>| <span data-ttu-id="9b18e-109">Valor</span><span class="sxs-lookup"><span data-stu-id="9b18e-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b18e-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9b18e-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b18e-111">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-111">1.1</span></span>|
|[<span data-ttu-id="9b18e-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9b18e-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b18e-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="9b18e-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="9b18e-114">Properties</span></span>

| <span data-ttu-id="9b18e-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="9b18e-115">Property</span></span> | <span data-ttu-id="9b18e-116">Modos</span><span class="sxs-lookup"><span data-stu-id="9b18e-116">Modes</span></span> | <span data-ttu-id="9b18e-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="9b18e-117">Return type</span></span> | <span data-ttu-id="9b18e-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="9b18e-118">Minimum</span></span><br><span data-ttu-id="9b18e-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="9b18e-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9b18e-120">auth</span><span class="sxs-lookup"><span data-stu-id="9b18e-120">auth</span></span>](#auth-auth) | <span data-ttu-id="9b18e-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="9b18e-121">Compose</span></span><br><span data-ttu-id="9b18e-122">Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-122">Read</span></span> | [<span data-ttu-id="9b18e-123">Auth</span><span class="sxs-lookup"><span data-stu-id="9b18e-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="9b18e-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="9b18e-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="9b18e-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="9b18e-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="9b18e-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="9b18e-126">Compose</span></span><br><span data-ttu-id="9b18e-127">Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-127">Read</span></span> | <span data-ttu-id="9b18e-128">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9b18e-128">String</span></span> | [<span data-ttu-id="9b18e-129">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9b18e-130">diagnostics</span><span class="sxs-lookup"><span data-stu-id="9b18e-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="9b18e-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="9b18e-131">Compose</span></span><br><span data-ttu-id="9b18e-132">Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-132">Read</span></span> | [<span data-ttu-id="9b18e-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="9b18e-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="9b18e-134">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9b18e-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="9b18e-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="9b18e-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="9b18e-136">Compose</span></span><br><span data-ttu-id="9b18e-137">Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-137">Read</span></span> | <span data-ttu-id="9b18e-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9b18e-138">String</span></span> | [<span data-ttu-id="9b18e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9b18e-140">host</span><span class="sxs-lookup"><span data-stu-id="9b18e-140">host</span></span>](#host-hosttype) | <span data-ttu-id="9b18e-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="9b18e-141">Compose</span></span><br><span data-ttu-id="9b18e-142">Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-142">Read</span></span> | [<span data-ttu-id="9b18e-143">HostType</span><span class="sxs-lookup"><span data-stu-id="9b18e-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="9b18e-144">1.5</span><span class="sxs-lookup"><span data-stu-id="9b18e-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="9b18e-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="9b18e-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="9b18e-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="9b18e-146">Compose</span></span><br><span data-ttu-id="9b18e-147">Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-147">Read</span></span> | [<span data-ttu-id="9b18e-148">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="9b18e-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="9b18e-149">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9b18e-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="9b18e-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="9b18e-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="9b18e-151">Compose</span></span><br><span data-ttu-id="9b18e-152">Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-152">Read</span></span> | [<span data-ttu-id="9b18e-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="9b18e-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="9b18e-154">Visualização</span><span class="sxs-lookup"><span data-stu-id="9b18e-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="9b18e-155">platform</span><span class="sxs-lookup"><span data-stu-id="9b18e-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="9b18e-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="9b18e-156">Compose</span></span><br><span data-ttu-id="9b18e-157">Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-157">Read</span></span> | [<span data-ttu-id="9b18e-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9b18e-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="9b18e-159">1.5</span><span class="sxs-lookup"><span data-stu-id="9b18e-159">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="9b18e-160">requirements</span><span class="sxs-lookup"><span data-stu-id="9b18e-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="9b18e-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="9b18e-161">Compose</span></span><br><span data-ttu-id="9b18e-162">Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-162">Read</span></span> | [<span data-ttu-id="9b18e-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="9b18e-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="9b18e-164">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9b18e-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="9b18e-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="9b18e-166">Escrever</span><span class="sxs-lookup"><span data-stu-id="9b18e-166">Compose</span></span><br><span data-ttu-id="9b18e-167">Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-167">Read</span></span> | [<span data-ttu-id="9b18e-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9b18e-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="9b18e-169">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9b18e-170">ui</span><span class="sxs-lookup"><span data-stu-id="9b18e-170">ui</span></span>](#ui-ui) | <span data-ttu-id="9b18e-171">Escrever</span><span class="sxs-lookup"><span data-stu-id="9b18e-171">Compose</span></span><br><span data-ttu-id="9b18e-172">Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-172">Read</span></span> | [<span data-ttu-id="9b18e-173">UI</span><span class="sxs-lookup"><span data-stu-id="9b18e-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="9b18e-174">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="9b18e-175">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="9b18e-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="9b18e-176">auth: [Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="9b18e-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="9b18e-177">Oferece suporte a [SSO (login único)](../../../outlook/authenticate-a-user-with-an-sso-token.md) fornecendo um método que permite ao aplicativo Office obter um token de acesso ao aplicativo Web do complemento.</span><span class="sxs-lookup"><span data-stu-id="9b18e-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="9b18e-178">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="9b18e-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="9b18e-179">Tipo</span><span class="sxs-lookup"><span data-stu-id="9b18e-179">Type</span></span>

*   [<span data-ttu-id="9b18e-180">Auth</span><span class="sxs-lookup"><span data-stu-id="9b18e-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="9b18e-181">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9b18e-181">Requirements</span></span>

|<span data-ttu-id="9b18e-182">Requisito</span><span class="sxs-lookup"><span data-stu-id="9b18e-182">Requirement</span></span>| <span data-ttu-id="9b18e-183">Valor</span><span class="sxs-lookup"><span data-stu-id="9b18e-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b18e-184">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9b18e-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b18e-185">Visualização</span><span class="sxs-lookup"><span data-stu-id="9b18e-185">Preview</span></span>|
|[<span data-ttu-id="9b18e-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9b18e-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b18e-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b18e-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9b18e-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="9b18e-189">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="9b18e-189">contentLanguage: String</span></span>

<span data-ttu-id="9b18e-190">Obtém a localidade (idioma) especificada pelo usuário para editar o item.</span><span class="sxs-lookup"><span data-stu-id="9b18e-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="9b18e-191">O `contentLanguage` valor reflete a **configuração** atual de Idioma de Edição especificada com opções de > de arquivo **> idioma** no aplicativo Office cliente.</span><span class="sxs-lookup"><span data-stu-id="9b18e-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="9b18e-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="9b18e-192">Type</span></span>

*   <span data-ttu-id="9b18e-193">String</span><span class="sxs-lookup"><span data-stu-id="9b18e-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b18e-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9b18e-194">Requirements</span></span>

|<span data-ttu-id="9b18e-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="9b18e-195">Requirement</span></span>| <span data-ttu-id="9b18e-196">Valor</span><span class="sxs-lookup"><span data-stu-id="9b18e-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b18e-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9b18e-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b18e-198">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-198">1.1</span></span>|
|[<span data-ttu-id="9b18e-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9b18e-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b18e-200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b18e-201">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9b18e-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="9b18e-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="9b18e-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="9b18e-203">Obtém informações sobre o ambiente no qual o complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="9b18e-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="9b18e-204">Tipo</span><span class="sxs-lookup"><span data-stu-id="9b18e-204">Type</span></span>

*   [<span data-ttu-id="9b18e-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="9b18e-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="9b18e-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9b18e-206">Requirements</span></span>

|<span data-ttu-id="9b18e-207">Requisito</span><span class="sxs-lookup"><span data-stu-id="9b18e-207">Requirement</span></span>| <span data-ttu-id="9b18e-208">Valor</span><span class="sxs-lookup"><span data-stu-id="9b18e-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b18e-209">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9b18e-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b18e-210">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-210">1.1</span></span>|
|[<span data-ttu-id="9b18e-211">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9b18e-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b18e-212">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b18e-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9b18e-213">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="9b18e-214">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="9b18e-214">displayLanguage: String</span></span>

<span data-ttu-id="9b18e-215">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente Office cliente.</span><span class="sxs-lookup"><span data-stu-id="9b18e-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="9b18e-216">O valor reflete a configuração atual de Idioma de Exibição especificada com Opções > > Idioma no aplicativo Office `displayLanguage` cliente.  </span><span class="sxs-lookup"><span data-stu-id="9b18e-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="9b18e-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="9b18e-217">Type</span></span>

*   <span data-ttu-id="9b18e-218">String</span><span class="sxs-lookup"><span data-stu-id="9b18e-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b18e-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9b18e-219">Requirements</span></span>

|<span data-ttu-id="9b18e-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="9b18e-220">Requirement</span></span>| <span data-ttu-id="9b18e-221">Valor</span><span class="sxs-lookup"><span data-stu-id="9b18e-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b18e-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9b18e-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b18e-223">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-223">1.1</span></span>|
|[<span data-ttu-id="9b18e-224">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9b18e-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b18e-225">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b18e-226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9b18e-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="9b18e-227">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="9b18e-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="9b18e-228">Obtém o Office aplicativo que está hospedando o complemento.</span><span class="sxs-lookup"><span data-stu-id="9b18e-228">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="9b18e-229">Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter o host.</span><span class="sxs-lookup"><span data-stu-id="9b18e-229">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="9b18e-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="9b18e-230">Type</span></span>

*   [<span data-ttu-id="9b18e-231">HostType</span><span class="sxs-lookup"><span data-stu-id="9b18e-231">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="9b18e-232">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9b18e-232">Requirements</span></span>

|<span data-ttu-id="9b18e-233">Requisito</span><span class="sxs-lookup"><span data-stu-id="9b18e-233">Requirement</span></span>| <span data-ttu-id="9b18e-234">Valor</span><span class="sxs-lookup"><span data-stu-id="9b18e-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b18e-235">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9b18e-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b18e-236">1,5</span><span class="sxs-lookup"><span data-stu-id="9b18e-236">1.5</span></span>|
|[<span data-ttu-id="9b18e-237">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9b18e-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b18e-238">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b18e-239">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9b18e-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="9b18e-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="9b18e-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="9b18e-241">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="9b18e-241">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="9b18e-242">Esse membro só tem suporte em Outlook no Windows.</span><span class="sxs-lookup"><span data-stu-id="9b18e-242">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="9b18e-243">O uso de cores de tema Office permite coordenar o esquema de cores do seu add **> Office-in** com o tema atual do Office selecionado pelo usuário com a interface do usuário > Office Conta > Office, que é aplicada em todos os aplicativos cliente Office.</span><span class="sxs-lookup"><span data-stu-id="9b18e-243">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="9b18e-244">Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="9b18e-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="9b18e-245">Tipo</span><span class="sxs-lookup"><span data-stu-id="9b18e-245">Type</span></span>

*   [<span data-ttu-id="9b18e-246">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="9b18e-246">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="9b18e-247">Propriedades</span><span class="sxs-lookup"><span data-stu-id="9b18e-247">Properties</span></span>

|<span data-ttu-id="9b18e-248">Nome</span><span class="sxs-lookup"><span data-stu-id="9b18e-248">Name</span></span>| <span data-ttu-id="9b18e-249">Tipo</span><span class="sxs-lookup"><span data-stu-id="9b18e-249">Type</span></span>| <span data-ttu-id="9b18e-250">Descrição</span><span class="sxs-lookup"><span data-stu-id="9b18e-250">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="9b18e-251">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9b18e-251">String</span></span>|<span data-ttu-id="9b18e-252">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="9b18e-252">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="9b18e-253">String</span><span class="sxs-lookup"><span data-stu-id="9b18e-253">String</span></span>|<span data-ttu-id="9b18e-254">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="9b18e-254">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="9b18e-255">String</span><span class="sxs-lookup"><span data-stu-id="9b18e-255">String</span></span>|<span data-ttu-id="9b18e-256">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="9b18e-256">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="9b18e-257">String</span><span class="sxs-lookup"><span data-stu-id="9b18e-257">String</span></span>|<span data-ttu-id="9b18e-258">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="9b18e-258">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9b18e-259">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9b18e-259">Requirements</span></span>

|<span data-ttu-id="9b18e-260">Requisito</span><span class="sxs-lookup"><span data-stu-id="9b18e-260">Requirement</span></span>| <span data-ttu-id="9b18e-261">Valor</span><span class="sxs-lookup"><span data-stu-id="9b18e-261">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b18e-262">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9b18e-262">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b18e-263">Visualização</span><span class="sxs-lookup"><span data-stu-id="9b18e-263">Preview</span></span>|
|[<span data-ttu-id="9b18e-264">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9b18e-264">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b18e-265">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-265">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b18e-266">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9b18e-266">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="9b18e-267">plataforma: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="9b18e-267">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="9b18e-268">Fornece a plataforma na qual o complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="9b18e-268">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="9b18e-269">Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter a plataforma.</span><span class="sxs-lookup"><span data-stu-id="9b18e-269">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="9b18e-270">Tipo</span><span class="sxs-lookup"><span data-stu-id="9b18e-270">Type</span></span>

*   [<span data-ttu-id="9b18e-271">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9b18e-271">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="9b18e-272">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9b18e-272">Requirements</span></span>

|<span data-ttu-id="9b18e-273">Requisito</span><span class="sxs-lookup"><span data-stu-id="9b18e-273">Requirement</span></span>| <span data-ttu-id="9b18e-274">Valor</span><span class="sxs-lookup"><span data-stu-id="9b18e-274">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b18e-275">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9b18e-275">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b18e-276">1,5</span><span class="sxs-lookup"><span data-stu-id="9b18e-276">1.5</span></span>|
|[<span data-ttu-id="9b18e-277">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9b18e-277">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b18e-278">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b18e-279">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9b18e-279">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="9b18e-280">requirements: [RequirementsSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="9b18e-280">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="9b18e-281">Fornece um método para determinar quais conjuntos de requisitos são suportados no aplicativo e na plataforma atual.</span><span class="sxs-lookup"><span data-stu-id="9b18e-281">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="9b18e-282">Tipo</span><span class="sxs-lookup"><span data-stu-id="9b18e-282">Type</span></span>

*   [<span data-ttu-id="9b18e-283">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="9b18e-283">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="9b18e-284">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9b18e-284">Requirements</span></span>

|<span data-ttu-id="9b18e-285">Requisito</span><span class="sxs-lookup"><span data-stu-id="9b18e-285">Requirement</span></span>| <span data-ttu-id="9b18e-286">Valor</span><span class="sxs-lookup"><span data-stu-id="9b18e-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b18e-287">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9b18e-287">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b18e-288">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-288">1.1</span></span>|
|[<span data-ttu-id="9b18e-289">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9b18e-289">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b18e-290">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b18e-291">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9b18e-291">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="9b18e-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="9b18e-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="9b18e-293">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="9b18e-293">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="9b18e-294">O objeto permite que você armazene e acesse dados para um complemento de email armazenado na caixa de correio de um usuário, de modo que está disponível para esse complemento quando ele está sendo executado de qualquer cliente Outlook usado para acessar essa caixa de `RoamingSettings` correio.</span><span class="sxs-lookup"><span data-stu-id="9b18e-294">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="9b18e-295">Tipo</span><span class="sxs-lookup"><span data-stu-id="9b18e-295">Type</span></span>

*   [<span data-ttu-id="9b18e-296">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9b18e-296">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="9b18e-297">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9b18e-297">Requirements</span></span>

|<span data-ttu-id="9b18e-298">Requisito</span><span class="sxs-lookup"><span data-stu-id="9b18e-298">Requirement</span></span>| <span data-ttu-id="9b18e-299">Valor</span><span class="sxs-lookup"><span data-stu-id="9b18e-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b18e-300">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9b18e-300">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b18e-301">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-301">1.1</span></span>|
|[<span data-ttu-id="9b18e-302">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9b18e-302">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="9b18e-303">Restrito</span><span class="sxs-lookup"><span data-stu-id="9b18e-303">Restricted</span></span>|
|[<span data-ttu-id="9b18e-304">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9b18e-304">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b18e-305">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-305">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="9b18e-306">interface do usuário: [interface do usuário](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="9b18e-306">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="9b18e-307">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus Office de usuário.</span><span class="sxs-lookup"><span data-stu-id="9b18e-307">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="9b18e-308">Tipo</span><span class="sxs-lookup"><span data-stu-id="9b18e-308">Type</span></span>

*   [<span data-ttu-id="9b18e-309">UI</span><span class="sxs-lookup"><span data-stu-id="9b18e-309">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="9b18e-310">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9b18e-310">Requirements</span></span>

|<span data-ttu-id="9b18e-311">Requisito</span><span class="sxs-lookup"><span data-stu-id="9b18e-311">Requirement</span></span>| <span data-ttu-id="9b18e-312">Valor</span><span class="sxs-lookup"><span data-stu-id="9b18e-312">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b18e-313">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9b18e-313">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b18e-314">1.1</span><span class="sxs-lookup"><span data-stu-id="9b18e-314">1.1</span></span>|
|[<span data-ttu-id="9b18e-315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9b18e-315">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b18e-316">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9b18e-316">Compose or Read</span></span>|
