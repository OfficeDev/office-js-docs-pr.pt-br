---
title: Office. Context – conjunto de requisitos de visualização
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de visualização da API da caixa de correio.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 8370df907aa3ab0534254057860c187cec583e6c
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570783"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="8e448-103">contexto (conjunto de requisitos de visualização da caixa de correio)</span><span class="sxs-lookup"><span data-stu-id="8e448-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="8e448-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="8e448-104">[Office](office.md).context</span></span>

<span data-ttu-id="8e448-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="8e448-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="8e448-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="8e448-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e448-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e448-107">Requirements</span></span>

|<span data-ttu-id="8e448-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="8e448-108">Requirement</span></span>| <span data-ttu-id="8e448-109">Valor</span><span class="sxs-lookup"><span data-stu-id="8e448-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e448-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8e448-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e448-111">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-111">1.1</span></span>|
|[<span data-ttu-id="8e448-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8e448-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8e448-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8e448-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="8e448-114">Properties</span></span>

| <span data-ttu-id="8e448-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="8e448-115">Property</span></span> | <span data-ttu-id="8e448-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="8e448-116">Modes</span></span> | <span data-ttu-id="8e448-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="8e448-117">Return type</span></span> | <span data-ttu-id="8e448-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="8e448-118">Minimum</span></span><br><span data-ttu-id="8e448-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="8e448-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8e448-120">autentica</span><span class="sxs-lookup"><span data-stu-id="8e448-120">auth</span></span>](#auth-auth) | <span data-ttu-id="8e448-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="8e448-121">Compose</span></span><br><span data-ttu-id="8e448-122">Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-122">Read</span></span> | [<span data-ttu-id="8e448-123">Auth</span><span class="sxs-lookup"><span data-stu-id="8e448-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="8e448-124">IdentityAPI 1,3</span><span class="sxs-lookup"><span data-stu-id="8e448-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="8e448-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="8e448-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="8e448-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="8e448-126">Compose</span></span><br><span data-ttu-id="8e448-127">Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-127">Read</span></span> | <span data-ttu-id="8e448-128">String</span><span class="sxs-lookup"><span data-stu-id="8e448-128">String</span></span> | [<span data-ttu-id="8e448-129">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8e448-130">la</span><span class="sxs-lookup"><span data-stu-id="8e448-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="8e448-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="8e448-131">Compose</span></span><br><span data-ttu-id="8e448-132">Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-132">Read</span></span> | [<span data-ttu-id="8e448-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="8e448-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="8e448-134">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8e448-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="8e448-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="8e448-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="8e448-136">Compose</span></span><br><span data-ttu-id="8e448-137">Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-137">Read</span></span> | <span data-ttu-id="8e448-138">String</span><span class="sxs-lookup"><span data-stu-id="8e448-138">String</span></span> | [<span data-ttu-id="8e448-139">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8e448-140">principal</span><span class="sxs-lookup"><span data-stu-id="8e448-140">host</span></span>](#host-hosttype) | <span data-ttu-id="8e448-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="8e448-141">Compose</span></span><br><span data-ttu-id="8e448-142">Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-142">Read</span></span> | [<span data-ttu-id="8e448-143">HostType</span><span class="sxs-lookup"><span data-stu-id="8e448-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="8e448-144">1,5</span><span class="sxs-lookup"><span data-stu-id="8e448-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="8e448-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="8e448-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="8e448-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="8e448-146">Compose</span></span><br><span data-ttu-id="8e448-147">Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-147">Read</span></span> | [<span data-ttu-id="8e448-148">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="8e448-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="8e448-149">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8e448-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="8e448-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="8e448-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="8e448-151">Compose</span></span><br><span data-ttu-id="8e448-152">Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-152">Read</span></span> | [<span data-ttu-id="8e448-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="8e448-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="8e448-154">Visualização</span><span class="sxs-lookup"><span data-stu-id="8e448-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="8e448-155">plataforma</span><span class="sxs-lookup"><span data-stu-id="8e448-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="8e448-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="8e448-156">Compose</span></span><br><span data-ttu-id="8e448-157">Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-157">Read</span></span> | [<span data-ttu-id="8e448-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="8e448-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="8e448-159">1,5</span><span class="sxs-lookup"><span data-stu-id="8e448-159">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="8e448-160">atende</span><span class="sxs-lookup"><span data-stu-id="8e448-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="8e448-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="8e448-161">Compose</span></span><br><span data-ttu-id="8e448-162">Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-162">Read</span></span> | [<span data-ttu-id="8e448-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="8e448-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="8e448-164">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8e448-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="8e448-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="8e448-166">Escrever</span><span class="sxs-lookup"><span data-stu-id="8e448-166">Compose</span></span><br><span data-ttu-id="8e448-167">Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-167">Read</span></span> | [<span data-ttu-id="8e448-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8e448-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="8e448-169">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8e448-170">ui</span><span class="sxs-lookup"><span data-stu-id="8e448-170">ui</span></span>](#ui-ui) | <span data-ttu-id="8e448-171">Escrever</span><span class="sxs-lookup"><span data-stu-id="8e448-171">Compose</span></span><br><span data-ttu-id="8e448-172">Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-172">Read</span></span> | [<span data-ttu-id="8e448-173">UI</span><span class="sxs-lookup"><span data-stu-id="8e448-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="8e448-174">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="8e448-175">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="8e448-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="8e448-176">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="8e448-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="8e448-177">Oferece suporte a [logon único (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , fornecendo um método que permite que o aplicativo do Office obtenha um token de acesso para o aplicativo Web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="8e448-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="8e448-178">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="8e448-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="8e448-179">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e448-179">Type</span></span>

*   [<span data-ttu-id="8e448-180">Auth</span><span class="sxs-lookup"><span data-stu-id="8e448-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="8e448-181">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e448-181">Requirements</span></span>

|<span data-ttu-id="8e448-182">Requisito</span><span class="sxs-lookup"><span data-stu-id="8e448-182">Requirement</span></span>| <span data-ttu-id="8e448-183">Valor</span><span class="sxs-lookup"><span data-stu-id="8e448-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e448-184">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8e448-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e448-185">Visualização</span><span class="sxs-lookup"><span data-stu-id="8e448-185">Preview</span></span>|
|[<span data-ttu-id="8e448-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8e448-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8e448-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e448-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8e448-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="8e448-189">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8e448-189">contentLanguage: String</span></span>

<span data-ttu-id="8e448-190">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="8e448-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="8e448-191">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="8e448-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="8e448-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e448-192">Type</span></span>

*   <span data-ttu-id="8e448-193">String</span><span class="sxs-lookup"><span data-stu-id="8e448-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e448-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e448-194">Requirements</span></span>

|<span data-ttu-id="8e448-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="8e448-195">Requirement</span></span>| <span data-ttu-id="8e448-196">Valor</span><span class="sxs-lookup"><span data-stu-id="8e448-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e448-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8e448-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e448-198">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-198">1.1</span></span>|
|[<span data-ttu-id="8e448-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8e448-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8e448-200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e448-201">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8e448-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="8e448-202">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="8e448-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="8e448-203">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="8e448-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="8e448-204">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e448-204">Type</span></span>

*   [<span data-ttu-id="8e448-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="8e448-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="8e448-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e448-206">Requirements</span></span>

|<span data-ttu-id="8e448-207">Requisito</span><span class="sxs-lookup"><span data-stu-id="8e448-207">Requirement</span></span>| <span data-ttu-id="8e448-208">Valor</span><span class="sxs-lookup"><span data-stu-id="8e448-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e448-209">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8e448-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e448-210">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-210">1.1</span></span>|
|[<span data-ttu-id="8e448-211">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8e448-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8e448-212">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e448-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8e448-213">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="8e448-214">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8e448-214">displayLanguage: String</span></span>

<span data-ttu-id="8e448-215">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="8e448-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="8e448-216">O `displayLanguage` valor reflete a configuração de **idioma de exibição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="8e448-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="8e448-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e448-217">Type</span></span>

*   <span data-ttu-id="8e448-218">String</span><span class="sxs-lookup"><span data-stu-id="8e448-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e448-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e448-219">Requirements</span></span>

|<span data-ttu-id="8e448-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="8e448-220">Requirement</span></span>| <span data-ttu-id="8e448-221">Valor</span><span class="sxs-lookup"><span data-stu-id="8e448-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e448-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8e448-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e448-223">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-223">1.1</span></span>|
|[<span data-ttu-id="8e448-224">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8e448-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8e448-225">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e448-226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8e448-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="8e448-227">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="8e448-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="8e448-228">Obtém o aplicativo do Office que está hospedando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="8e448-228">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="8e448-229">Como alternativa, você pode usar a propriedade [Office. Context. Diagnostics](#diagnostics-contextinformation) para obter o host.</span><span class="sxs-lookup"><span data-stu-id="8e448-229">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="8e448-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e448-230">Type</span></span>

*   [<span data-ttu-id="8e448-231">HostType</span><span class="sxs-lookup"><span data-stu-id="8e448-231">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="8e448-232">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e448-232">Requirements</span></span>

|<span data-ttu-id="8e448-233">Requisito</span><span class="sxs-lookup"><span data-stu-id="8e448-233">Requirement</span></span>| <span data-ttu-id="8e448-234">Valor</span><span class="sxs-lookup"><span data-stu-id="8e448-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e448-235">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8e448-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e448-236">1,5</span><span class="sxs-lookup"><span data-stu-id="8e448-236">1.5</span></span>|
|[<span data-ttu-id="8e448-237">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8e448-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8e448-238">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e448-239">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8e448-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="8e448-240">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="8e448-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="8e448-241">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="8e448-241">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="8e448-242">Só há suporte para esse membro no Outlook no Windows.</span><span class="sxs-lookup"><span data-stu-id="8e448-242">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="8e448-243">O uso de cores de temas do Office permite coordenar o esquema de cores do seu suplemento com o tema atual do Office selecionado pelo usuário com a **conta de arquivo > office > Office Theme UI**, que é aplicada em todos os aplicativos cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="8e448-243">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="8e448-244">Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="8e448-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="8e448-245">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e448-245">Type</span></span>

*   [<span data-ttu-id="8e448-246">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="8e448-246">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="8e448-247">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="8e448-247">Properties:</span></span>

|<span data-ttu-id="8e448-248">Nome</span><span class="sxs-lookup"><span data-stu-id="8e448-248">Name</span></span>| <span data-ttu-id="8e448-249">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e448-249">Type</span></span>| <span data-ttu-id="8e448-250">Descrição</span><span class="sxs-lookup"><span data-stu-id="8e448-250">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="8e448-251">String</span><span class="sxs-lookup"><span data-stu-id="8e448-251">String</span></span>|<span data-ttu-id="8e448-252">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="8e448-252">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="8e448-253">String</span><span class="sxs-lookup"><span data-stu-id="8e448-253">String</span></span>|<span data-ttu-id="8e448-254">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="8e448-254">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="8e448-255">String</span><span class="sxs-lookup"><span data-stu-id="8e448-255">String</span></span>|<span data-ttu-id="8e448-256">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="8e448-256">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="8e448-257">String</span><span class="sxs-lookup"><span data-stu-id="8e448-257">String</span></span>|<span data-ttu-id="8e448-258">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="8e448-258">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e448-259">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e448-259">Requirements</span></span>

|<span data-ttu-id="8e448-260">Requisito</span><span class="sxs-lookup"><span data-stu-id="8e448-260">Requirement</span></span>| <span data-ttu-id="8e448-261">Valor</span><span class="sxs-lookup"><span data-stu-id="8e448-261">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e448-262">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8e448-262">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e448-263">Visualização</span><span class="sxs-lookup"><span data-stu-id="8e448-263">Preview</span></span>|
|[<span data-ttu-id="8e448-264">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8e448-264">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8e448-265">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-265">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e448-266">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8e448-266">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="8e448-267">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="8e448-267">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="8e448-268">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="8e448-268">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="8e448-269">Como alternativa, você pode usar a propriedade [Office. Context. Diagnostics](#diagnostics-contextinformation) para obter a plataforma.</span><span class="sxs-lookup"><span data-stu-id="8e448-269">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="8e448-270">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e448-270">Type</span></span>

*   [<span data-ttu-id="8e448-271">PlatformType</span><span class="sxs-lookup"><span data-stu-id="8e448-271">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="8e448-272">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e448-272">Requirements</span></span>

|<span data-ttu-id="8e448-273">Requisito</span><span class="sxs-lookup"><span data-stu-id="8e448-273">Requirement</span></span>| <span data-ttu-id="8e448-274">Valor</span><span class="sxs-lookup"><span data-stu-id="8e448-274">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e448-275">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8e448-275">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e448-276">1,5</span><span class="sxs-lookup"><span data-stu-id="8e448-276">1.5</span></span>|
|[<span data-ttu-id="8e448-277">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8e448-277">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8e448-278">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e448-279">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8e448-279">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="8e448-280">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="8e448-280">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="8e448-281">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o aplicativo atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="8e448-281">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="8e448-282">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e448-282">Type</span></span>

*   [<span data-ttu-id="8e448-283">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="8e448-283">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="8e448-284">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e448-284">Requirements</span></span>

|<span data-ttu-id="8e448-285">Requisito</span><span class="sxs-lookup"><span data-stu-id="8e448-285">Requirement</span></span>| <span data-ttu-id="8e448-286">Valor</span><span class="sxs-lookup"><span data-stu-id="8e448-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e448-287">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8e448-287">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e448-288">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-288">1.1</span></span>|
|[<span data-ttu-id="8e448-289">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8e448-289">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8e448-290">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e448-291">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8e448-291">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="8e448-292">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="8e448-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="8e448-293">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="8e448-293">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="8e448-294">O `RoamingSettings` objeto permite armazenar e acessar dados de um suplemento de email armazenado na caixa de correio de um usuário, para que esteja disponível para esse suplemento quando ele estiver sendo executado a partir de qualquer cliente do Outlook usado para acessar a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="8e448-294">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="8e448-295">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e448-295">Type</span></span>

*   [<span data-ttu-id="8e448-296">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8e448-296">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="8e448-297">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e448-297">Requirements</span></span>

|<span data-ttu-id="8e448-298">Requisito</span><span class="sxs-lookup"><span data-stu-id="8e448-298">Requirement</span></span>| <span data-ttu-id="8e448-299">Valor</span><span class="sxs-lookup"><span data-stu-id="8e448-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e448-300">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8e448-300">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e448-301">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-301">1.1</span></span>|
|[<span data-ttu-id="8e448-302">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8e448-302">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="8e448-303">Restrito</span><span class="sxs-lookup"><span data-stu-id="8e448-303">Restricted</span></span>|
|[<span data-ttu-id="8e448-304">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8e448-304">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8e448-305">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-305">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="8e448-306">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="8e448-306">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="8e448-307">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="8e448-307">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="8e448-308">Tipo</span><span class="sxs-lookup"><span data-stu-id="8e448-308">Type</span></span>

*   [<span data-ttu-id="8e448-309">UI</span><span class="sxs-lookup"><span data-stu-id="8e448-309">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="8e448-310">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8e448-310">Requirements</span></span>

|<span data-ttu-id="8e448-311">Requisito</span><span class="sxs-lookup"><span data-stu-id="8e448-311">Requirement</span></span>| <span data-ttu-id="8e448-312">Valor</span><span class="sxs-lookup"><span data-stu-id="8e448-312">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e448-313">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8e448-313">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8e448-314">1.1</span><span class="sxs-lookup"><span data-stu-id="8e448-314">1.1</span></span>|
|[<span data-ttu-id="8e448-315">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8e448-315">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8e448-316">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8e448-316">Compose or Read</span></span>|
