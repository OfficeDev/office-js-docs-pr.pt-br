---
title: Office. Context – conjunto de requisitos de visualização
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de visualização da API da caixa de correio.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 8286434d2cbfc11cf0d16f8bd014b4760f0337ff
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626404"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="890aa-103">contexto (conjunto de requisitos de visualização da caixa de correio)</span><span class="sxs-lookup"><span data-stu-id="890aa-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="890aa-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="890aa-104">[Office](office.md).context</span></span>

<span data-ttu-id="890aa-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="890aa-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="890aa-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="890aa-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="890aa-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="890aa-107">Requirements</span></span>

|<span data-ttu-id="890aa-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="890aa-108">Requirement</span></span>| <span data-ttu-id="890aa-109">Valor</span><span class="sxs-lookup"><span data-stu-id="890aa-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="890aa-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="890aa-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="890aa-111">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-111">1.1</span></span>|
|[<span data-ttu-id="890aa-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="890aa-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="890aa-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="890aa-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="890aa-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="890aa-114">Properties</span></span>

| <span data-ttu-id="890aa-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="890aa-115">Property</span></span> | <span data-ttu-id="890aa-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="890aa-116">Modes</span></span> | <span data-ttu-id="890aa-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="890aa-117">Return type</span></span> | <span data-ttu-id="890aa-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="890aa-118">Minimum</span></span><br><span data-ttu-id="890aa-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="890aa-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="890aa-120">autentica</span><span class="sxs-lookup"><span data-stu-id="890aa-120">auth</span></span>](#auth-auth) | <span data-ttu-id="890aa-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="890aa-121">Compose</span></span><br><span data-ttu-id="890aa-122">Leitura</span><span class="sxs-lookup"><span data-stu-id="890aa-122">Read</span></span> | [<span data-ttu-id="890aa-123">Auth</span><span class="sxs-lookup"><span data-stu-id="890aa-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="890aa-124">IdentityAPI 1,3</span><span class="sxs-lookup"><span data-stu-id="890aa-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="890aa-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="890aa-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="890aa-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="890aa-126">Compose</span></span><br><span data-ttu-id="890aa-127">Leitura</span><span class="sxs-lookup"><span data-stu-id="890aa-127">Read</span></span> | <span data-ttu-id="890aa-128">String</span><span class="sxs-lookup"><span data-stu-id="890aa-128">String</span></span> | [<span data-ttu-id="890aa-129">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="890aa-130">la</span><span class="sxs-lookup"><span data-stu-id="890aa-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="890aa-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="890aa-131">Compose</span></span><br><span data-ttu-id="890aa-132">Leitura</span><span class="sxs-lookup"><span data-stu-id="890aa-132">Read</span></span> | [<span data-ttu-id="890aa-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="890aa-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="890aa-134">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="890aa-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="890aa-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="890aa-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="890aa-136">Compose</span></span><br><span data-ttu-id="890aa-137">Leitura</span><span class="sxs-lookup"><span data-stu-id="890aa-137">Read</span></span> | <span data-ttu-id="890aa-138">String</span><span class="sxs-lookup"><span data-stu-id="890aa-138">String</span></span> | [<span data-ttu-id="890aa-139">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="890aa-140">principal</span><span class="sxs-lookup"><span data-stu-id="890aa-140">host</span></span>](#host-hosttype) | <span data-ttu-id="890aa-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="890aa-141">Compose</span></span><br><span data-ttu-id="890aa-142">Leitura</span><span class="sxs-lookup"><span data-stu-id="890aa-142">Read</span></span> | [<span data-ttu-id="890aa-143">HostType</span><span class="sxs-lookup"><span data-stu-id="890aa-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="890aa-144">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="890aa-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="890aa-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="890aa-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="890aa-146">Compose</span></span><br><span data-ttu-id="890aa-147">Leitura</span><span class="sxs-lookup"><span data-stu-id="890aa-147">Read</span></span> | [<span data-ttu-id="890aa-148">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="890aa-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="890aa-149">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="890aa-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="890aa-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="890aa-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="890aa-151">Compose</span></span><br><span data-ttu-id="890aa-152">Leitura</span><span class="sxs-lookup"><span data-stu-id="890aa-152">Read</span></span> | [<span data-ttu-id="890aa-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="890aa-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="890aa-154">Visualização</span><span class="sxs-lookup"><span data-stu-id="890aa-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="890aa-155">plataforma</span><span class="sxs-lookup"><span data-stu-id="890aa-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="890aa-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="890aa-156">Compose</span></span><br><span data-ttu-id="890aa-157">Leitura</span><span class="sxs-lookup"><span data-stu-id="890aa-157">Read</span></span> | [<span data-ttu-id="890aa-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="890aa-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="890aa-159">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="890aa-160">atende</span><span class="sxs-lookup"><span data-stu-id="890aa-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="890aa-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="890aa-161">Compose</span></span><br><span data-ttu-id="890aa-162">Leitura</span><span class="sxs-lookup"><span data-stu-id="890aa-162">Read</span></span> | [<span data-ttu-id="890aa-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="890aa-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="890aa-164">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="890aa-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="890aa-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="890aa-166">Escrever</span><span class="sxs-lookup"><span data-stu-id="890aa-166">Compose</span></span><br><span data-ttu-id="890aa-167">Leitura</span><span class="sxs-lookup"><span data-stu-id="890aa-167">Read</span></span> | [<span data-ttu-id="890aa-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="890aa-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="890aa-169">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="890aa-170">ui</span><span class="sxs-lookup"><span data-stu-id="890aa-170">ui</span></span>](#ui-ui) | <span data-ttu-id="890aa-171">Escrever</span><span class="sxs-lookup"><span data-stu-id="890aa-171">Compose</span></span><br><span data-ttu-id="890aa-172">Leitura</span><span class="sxs-lookup"><span data-stu-id="890aa-172">Read</span></span> | [<span data-ttu-id="890aa-173">UI</span><span class="sxs-lookup"><span data-stu-id="890aa-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="890aa-174">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="890aa-175">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="890aa-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="890aa-176">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="890aa-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="890aa-177">Oferece suporte a [logon único (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , fornecendo um método que permite que o aplicativo do Office obtenha um token de acesso para o aplicativo Web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="890aa-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="890aa-178">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="890aa-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="890aa-179">Tipo</span><span class="sxs-lookup"><span data-stu-id="890aa-179">Type</span></span>

*   [<span data-ttu-id="890aa-180">Auth</span><span class="sxs-lookup"><span data-stu-id="890aa-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="890aa-181">Requisitos</span><span class="sxs-lookup"><span data-stu-id="890aa-181">Requirements</span></span>

|<span data-ttu-id="890aa-182">Requisito</span><span class="sxs-lookup"><span data-stu-id="890aa-182">Requirement</span></span>| <span data-ttu-id="890aa-183">Valor</span><span class="sxs-lookup"><span data-stu-id="890aa-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="890aa-184">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="890aa-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="890aa-185">Visualização</span><span class="sxs-lookup"><span data-stu-id="890aa-185">Preview</span></span>|
|[<span data-ttu-id="890aa-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="890aa-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="890aa-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="890aa-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="890aa-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="890aa-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="890aa-189">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="890aa-189">contentLanguage: String</span></span>

<span data-ttu-id="890aa-190">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="890aa-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="890aa-191">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="890aa-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="890aa-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="890aa-192">Type</span></span>

*   <span data-ttu-id="890aa-193">String</span><span class="sxs-lookup"><span data-stu-id="890aa-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="890aa-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="890aa-194">Requirements</span></span>

|<span data-ttu-id="890aa-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="890aa-195">Requirement</span></span>| <span data-ttu-id="890aa-196">Valor</span><span class="sxs-lookup"><span data-stu-id="890aa-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="890aa-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="890aa-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="890aa-198">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-198">1.1</span></span>|
|[<span data-ttu-id="890aa-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="890aa-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="890aa-200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="890aa-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="890aa-201">Exemplo</span><span class="sxs-lookup"><span data-stu-id="890aa-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="890aa-202">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="890aa-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="890aa-203">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="890aa-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="890aa-204">Tipo</span><span class="sxs-lookup"><span data-stu-id="890aa-204">Type</span></span>

*   [<span data-ttu-id="890aa-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="890aa-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="890aa-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="890aa-206">Requirements</span></span>

|<span data-ttu-id="890aa-207">Requisito</span><span class="sxs-lookup"><span data-stu-id="890aa-207">Requirement</span></span>| <span data-ttu-id="890aa-208">Valor</span><span class="sxs-lookup"><span data-stu-id="890aa-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="890aa-209">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="890aa-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="890aa-210">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-210">1.1</span></span>|
|[<span data-ttu-id="890aa-211">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="890aa-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="890aa-212">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="890aa-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="890aa-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="890aa-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="890aa-214">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="890aa-214">displayLanguage: String</span></span>

<span data-ttu-id="890aa-215">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="890aa-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="890aa-216">O `displayLanguage` valor reflete a configuração de **idioma de exibição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="890aa-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="890aa-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="890aa-217">Type</span></span>

*   <span data-ttu-id="890aa-218">String</span><span class="sxs-lookup"><span data-stu-id="890aa-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="890aa-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="890aa-219">Requirements</span></span>

|<span data-ttu-id="890aa-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="890aa-220">Requirement</span></span>| <span data-ttu-id="890aa-221">Valor</span><span class="sxs-lookup"><span data-stu-id="890aa-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="890aa-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="890aa-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="890aa-223">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-223">1.1</span></span>|
|[<span data-ttu-id="890aa-224">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="890aa-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="890aa-225">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="890aa-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="890aa-226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="890aa-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="890aa-227">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="890aa-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="890aa-228">Obtém o aplicativo do Office que está hospedando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="890aa-228">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="890aa-229">Tipo</span><span class="sxs-lookup"><span data-stu-id="890aa-229">Type</span></span>

*   [<span data-ttu-id="890aa-230">HostType</span><span class="sxs-lookup"><span data-stu-id="890aa-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="890aa-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="890aa-231">Requirements</span></span>

|<span data-ttu-id="890aa-232">Requisito</span><span class="sxs-lookup"><span data-stu-id="890aa-232">Requirement</span></span>| <span data-ttu-id="890aa-233">Valor</span><span class="sxs-lookup"><span data-stu-id="890aa-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="890aa-234">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="890aa-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="890aa-235">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-235">1.1</span></span>|
|[<span data-ttu-id="890aa-236">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="890aa-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="890aa-237">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="890aa-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="890aa-238">Exemplo</span><span class="sxs-lookup"><span data-stu-id="890aa-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="890aa-239">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="890aa-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="890aa-240">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="890aa-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="890aa-241">Só há suporte para esse membro no Outlook no Windows.</span><span class="sxs-lookup"><span data-stu-id="890aa-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="890aa-242">O uso de cores de temas do Office permite coordenar o esquema de cores do seu suplemento com o tema atual do Office selecionado pelo usuário com a **conta de arquivo > office > Office Theme UI**, que é aplicada em todos os aplicativos cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="890aa-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="890aa-243">Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="890aa-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="890aa-244">Tipo</span><span class="sxs-lookup"><span data-stu-id="890aa-244">Type</span></span>

*   [<span data-ttu-id="890aa-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="890aa-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="890aa-246">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="890aa-246">Properties:</span></span>

|<span data-ttu-id="890aa-247">Nome</span><span class="sxs-lookup"><span data-stu-id="890aa-247">Name</span></span>| <span data-ttu-id="890aa-248">Tipo</span><span class="sxs-lookup"><span data-stu-id="890aa-248">Type</span></span>| <span data-ttu-id="890aa-249">Descrição</span><span class="sxs-lookup"><span data-stu-id="890aa-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="890aa-250">String</span><span class="sxs-lookup"><span data-stu-id="890aa-250">String</span></span>|<span data-ttu-id="890aa-251">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="890aa-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="890aa-252">String</span><span class="sxs-lookup"><span data-stu-id="890aa-252">String</span></span>|<span data-ttu-id="890aa-253">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="890aa-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="890aa-254">String</span><span class="sxs-lookup"><span data-stu-id="890aa-254">String</span></span>|<span data-ttu-id="890aa-255">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="890aa-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="890aa-256">String</span><span class="sxs-lookup"><span data-stu-id="890aa-256">String</span></span>|<span data-ttu-id="890aa-257">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="890aa-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="890aa-258">Requisitos</span><span class="sxs-lookup"><span data-stu-id="890aa-258">Requirements</span></span>

|<span data-ttu-id="890aa-259">Requisito</span><span class="sxs-lookup"><span data-stu-id="890aa-259">Requirement</span></span>| <span data-ttu-id="890aa-260">Valor</span><span class="sxs-lookup"><span data-stu-id="890aa-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="890aa-261">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="890aa-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="890aa-262">Visualização</span><span class="sxs-lookup"><span data-stu-id="890aa-262">Preview</span></span>|
|[<span data-ttu-id="890aa-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="890aa-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="890aa-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="890aa-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="890aa-265">Exemplo</span><span class="sxs-lookup"><span data-stu-id="890aa-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="890aa-266">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="890aa-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="890aa-267">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="890aa-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="890aa-268">Tipo</span><span class="sxs-lookup"><span data-stu-id="890aa-268">Type</span></span>

*   [<span data-ttu-id="890aa-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="890aa-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="890aa-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="890aa-270">Requirements</span></span>

|<span data-ttu-id="890aa-271">Requisito</span><span class="sxs-lookup"><span data-stu-id="890aa-271">Requirement</span></span>| <span data-ttu-id="890aa-272">Valor</span><span class="sxs-lookup"><span data-stu-id="890aa-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="890aa-273">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="890aa-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="890aa-274">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-274">1.1</span></span>|
|[<span data-ttu-id="890aa-275">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="890aa-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="890aa-276">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="890aa-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="890aa-277">Exemplo</span><span class="sxs-lookup"><span data-stu-id="890aa-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="890aa-278">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="890aa-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="890aa-279">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o aplicativo atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="890aa-279">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="890aa-280">Tipo</span><span class="sxs-lookup"><span data-stu-id="890aa-280">Type</span></span>

*   [<span data-ttu-id="890aa-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="890aa-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="890aa-282">Requisitos</span><span class="sxs-lookup"><span data-stu-id="890aa-282">Requirements</span></span>

|<span data-ttu-id="890aa-283">Requisito</span><span class="sxs-lookup"><span data-stu-id="890aa-283">Requirement</span></span>| <span data-ttu-id="890aa-284">Valor</span><span class="sxs-lookup"><span data-stu-id="890aa-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="890aa-285">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="890aa-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="890aa-286">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-286">1.1</span></span>|
|[<span data-ttu-id="890aa-287">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="890aa-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="890aa-288">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="890aa-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="890aa-289">Exemplo</span><span class="sxs-lookup"><span data-stu-id="890aa-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="890aa-290">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="890aa-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="890aa-291">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="890aa-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="890aa-292">O `RoamingSettings` objeto permite armazenar e acessar dados de um suplemento de email armazenado na caixa de correio de um usuário, para que esteja disponível para esse suplemento quando ele estiver sendo executado a partir de qualquer cliente do Outlook usado para acessar a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="890aa-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="890aa-293">Tipo</span><span class="sxs-lookup"><span data-stu-id="890aa-293">Type</span></span>

*   [<span data-ttu-id="890aa-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="890aa-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="890aa-295">Requisitos</span><span class="sxs-lookup"><span data-stu-id="890aa-295">Requirements</span></span>

|<span data-ttu-id="890aa-296">Requisito</span><span class="sxs-lookup"><span data-stu-id="890aa-296">Requirement</span></span>| <span data-ttu-id="890aa-297">Valor</span><span class="sxs-lookup"><span data-stu-id="890aa-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="890aa-298">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="890aa-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="890aa-299">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-299">1.1</span></span>|
|[<span data-ttu-id="890aa-300">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="890aa-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="890aa-301">Restrito</span><span class="sxs-lookup"><span data-stu-id="890aa-301">Restricted</span></span>|
|[<span data-ttu-id="890aa-302">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="890aa-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="890aa-303">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="890aa-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="890aa-304">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="890aa-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="890aa-305">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="890aa-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="890aa-306">Tipo</span><span class="sxs-lookup"><span data-stu-id="890aa-306">Type</span></span>

*   [<span data-ttu-id="890aa-307">UI</span><span class="sxs-lookup"><span data-stu-id="890aa-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="890aa-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="890aa-308">Requirements</span></span>

|<span data-ttu-id="890aa-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="890aa-309">Requirement</span></span>| <span data-ttu-id="890aa-310">Valor</span><span class="sxs-lookup"><span data-stu-id="890aa-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="890aa-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="890aa-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="890aa-312">1.1</span><span class="sxs-lookup"><span data-stu-id="890aa-312">1.1</span></span>|
|[<span data-ttu-id="890aa-313">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="890aa-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="890aa-314">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="890aa-314">Compose or Read</span></span>|
