---
title: Office. Context – conjunto de requisitos 1,9
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,9.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 6b2657d1e608bd1820d3814d9a6bfab67681824c
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628039"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="dfaf4-103">contexto (conjunto de requisitos de caixa de correio 1,9)</span><span class="sxs-lookup"><span data-stu-id="dfaf4-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="dfaf4-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="dfaf4-104">[Office](office.md).context</span></span>

<span data-ttu-id="dfaf4-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="dfaf4-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="dfaf4-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dfaf4-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfaf4-107">Requirements</span></span>

|<span data-ttu-id="dfaf4-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfaf4-108">Requirement</span></span>| <span data-ttu-id="dfaf4-109">Valor</span><span class="sxs-lookup"><span data-stu-id="dfaf4-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfaf4-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfaf4-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfaf4-111">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-111">1.1</span></span>|
|[<span data-ttu-id="dfaf4-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfaf4-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfaf4-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfaf4-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="dfaf4-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="dfaf4-114">Properties</span></span>

| <span data-ttu-id="dfaf4-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="dfaf4-115">Property</span></span> | <span data-ttu-id="dfaf4-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="dfaf4-116">Modes</span></span> | <span data-ttu-id="dfaf4-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="dfaf4-117">Return type</span></span> | <span data-ttu-id="dfaf4-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="dfaf4-118">Minimum</span></span><br><span data-ttu-id="dfaf4-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="dfaf4-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dfaf4-120">autentica</span><span class="sxs-lookup"><span data-stu-id="dfaf4-120">auth</span></span>](#auth-auth) | <span data-ttu-id="dfaf4-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfaf4-121">Compose</span></span><br><span data-ttu-id="dfaf4-122">Leitura</span><span class="sxs-lookup"><span data-stu-id="dfaf4-122">Read</span></span> | [<span data-ttu-id="dfaf4-123">Auth</span><span class="sxs-lookup"><span data-stu-id="dfaf4-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dfaf4-124">IdentityAPI 1,3</span><span class="sxs-lookup"><span data-stu-id="dfaf4-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="dfaf4-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="dfaf4-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="dfaf4-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfaf4-126">Compose</span></span><br><span data-ttu-id="dfaf4-127">Leitura</span><span class="sxs-lookup"><span data-stu-id="dfaf4-127">Read</span></span> | <span data-ttu-id="dfaf4-128">String</span><span class="sxs-lookup"><span data-stu-id="dfaf4-128">String</span></span> | [<span data-ttu-id="dfaf4-129">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfaf4-130">la</span><span class="sxs-lookup"><span data-stu-id="dfaf4-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="dfaf4-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfaf4-131">Compose</span></span><br><span data-ttu-id="dfaf4-132">Leitura</span><span class="sxs-lookup"><span data-stu-id="dfaf4-132">Read</span></span> | [<span data-ttu-id="dfaf4-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="dfaf4-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dfaf4-134">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfaf4-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="dfaf4-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="dfaf4-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfaf4-136">Compose</span></span><br><span data-ttu-id="dfaf4-137">Leitura</span><span class="sxs-lookup"><span data-stu-id="dfaf4-137">Read</span></span> | <span data-ttu-id="dfaf4-138">String</span><span class="sxs-lookup"><span data-stu-id="dfaf4-138">String</span></span> | [<span data-ttu-id="dfaf4-139">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfaf4-140">principal</span><span class="sxs-lookup"><span data-stu-id="dfaf4-140">host</span></span>](#host-hosttype) | <span data-ttu-id="dfaf4-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfaf4-141">Compose</span></span><br><span data-ttu-id="dfaf4-142">Leitura</span><span class="sxs-lookup"><span data-stu-id="dfaf4-142">Read</span></span> | [<span data-ttu-id="dfaf4-143">HostType</span><span class="sxs-lookup"><span data-stu-id="dfaf4-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dfaf4-144">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfaf4-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="dfaf4-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="dfaf4-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfaf4-146">Compose</span></span><br><span data-ttu-id="dfaf4-147">Leitura</span><span class="sxs-lookup"><span data-stu-id="dfaf4-147">Read</span></span> | [<span data-ttu-id="dfaf4-148">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="dfaf4-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dfaf4-149">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfaf4-150">plataforma</span><span class="sxs-lookup"><span data-stu-id="dfaf4-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="dfaf4-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfaf4-151">Compose</span></span><br><span data-ttu-id="dfaf4-152">Leitura</span><span class="sxs-lookup"><span data-stu-id="dfaf4-152">Read</span></span> | [<span data-ttu-id="dfaf4-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="dfaf4-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dfaf4-154">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfaf4-155">atende</span><span class="sxs-lookup"><span data-stu-id="dfaf4-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="dfaf4-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfaf4-156">Compose</span></span><br><span data-ttu-id="dfaf4-157">Leitura</span><span class="sxs-lookup"><span data-stu-id="dfaf4-157">Read</span></span> | [<span data-ttu-id="dfaf4-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="dfaf4-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dfaf4-159">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfaf4-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="dfaf4-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="dfaf4-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfaf4-161">Compose</span></span><br><span data-ttu-id="dfaf4-162">Leitura</span><span class="sxs-lookup"><span data-stu-id="dfaf4-162">Read</span></span> | [<span data-ttu-id="dfaf4-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="dfaf4-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dfaf4-164">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dfaf4-165">ui</span><span class="sxs-lookup"><span data-stu-id="dfaf4-165">ui</span></span>](#ui-ui) | <span data-ttu-id="dfaf4-166">Escrever</span><span class="sxs-lookup"><span data-stu-id="dfaf4-166">Compose</span></span><br><span data-ttu-id="dfaf4-167">Leitura</span><span class="sxs-lookup"><span data-stu-id="dfaf4-167">Read</span></span> | [<span data-ttu-id="dfaf4-168">UI</span><span class="sxs-lookup"><span data-stu-id="dfaf4-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dfaf4-169">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="dfaf4-170">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="dfaf4-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="dfaf4-171">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="dfaf4-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="dfaf4-172">Oferece suporte a [logon único (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , fornecendo um método que permite que o aplicativo do Office obtenha um token de acesso para o aplicativo Web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="dfaf4-173">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="dfaf4-174">Confira [IdentityAPI 1,3 conjunto de requisitos](../../requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="dfaf4-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="dfaf4-175">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-175">Type</span></span>

*   [<span data-ttu-id="dfaf4-176">Auth</span><span class="sxs-lookup"><span data-stu-id="dfaf4-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="dfaf4-177">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfaf4-177">Requirements</span></span>

|<span data-ttu-id="dfaf4-178">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfaf4-178">Requirement</span></span>| <span data-ttu-id="dfaf4-179">Valor</span><span class="sxs-lookup"><span data-stu-id="dfaf4-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfaf4-180">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfaf4-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfaf4-181">N/D</span><span class="sxs-lookup"><span data-stu-id="dfaf4-181">N/A</span></span>|
|[<span data-ttu-id="dfaf4-182">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfaf4-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfaf4-183">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfaf4-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfaf4-184">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-184">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="dfaf4-185">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dfaf4-185">contentLanguage: String</span></span>

<span data-ttu-id="dfaf4-186">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="dfaf4-187">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="dfaf4-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-188">Type</span></span>

*   <span data-ttu-id="dfaf4-189">String</span><span class="sxs-lookup"><span data-stu-id="dfaf4-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dfaf4-190">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfaf4-190">Requirements</span></span>

|<span data-ttu-id="dfaf4-191">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfaf4-191">Requirement</span></span>| <span data-ttu-id="dfaf4-192">Valor</span><span class="sxs-lookup"><span data-stu-id="dfaf4-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfaf4-193">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfaf4-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfaf4-194">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-194">1.1</span></span>|
|[<span data-ttu-id="dfaf4-195">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfaf4-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfaf4-196">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfaf4-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfaf4-197">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-197">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="dfaf4-198">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="dfaf4-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="dfaf4-199">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="dfaf4-200">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-200">Type</span></span>

*   [<span data-ttu-id="dfaf4-201">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="dfaf4-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="dfaf4-202">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfaf4-202">Requirements</span></span>

|<span data-ttu-id="dfaf4-203">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfaf4-203">Requirement</span></span>| <span data-ttu-id="dfaf4-204">Valor</span><span class="sxs-lookup"><span data-stu-id="dfaf4-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfaf4-205">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfaf4-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfaf4-206">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-206">1.1</span></span>|
|[<span data-ttu-id="dfaf4-207">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfaf4-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfaf4-208">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfaf4-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfaf4-209">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-209">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="dfaf4-210">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dfaf4-210">displayLanguage: String</span></span>

<span data-ttu-id="dfaf4-211">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="dfaf4-212">O `displayLanguage` valor reflete a configuração de **idioma de exibição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="dfaf4-213">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-213">Type</span></span>

*   <span data-ttu-id="dfaf4-214">String</span><span class="sxs-lookup"><span data-stu-id="dfaf4-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dfaf4-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfaf4-215">Requirements</span></span>

|<span data-ttu-id="dfaf4-216">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfaf4-216">Requirement</span></span>| <span data-ttu-id="dfaf4-217">Valor</span><span class="sxs-lookup"><span data-stu-id="dfaf4-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfaf4-218">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfaf4-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfaf4-219">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-219">1.1</span></span>|
|[<span data-ttu-id="dfaf4-220">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfaf4-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfaf4-221">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfaf4-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfaf4-222">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-222">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="dfaf4-223">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="dfaf4-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="dfaf4-224">Obtém o aplicativo do Office que está hospedando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-224">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="dfaf4-225">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-225">Type</span></span>

*   [<span data-ttu-id="dfaf4-226">HostType</span><span class="sxs-lookup"><span data-stu-id="dfaf4-226">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="dfaf4-227">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfaf4-227">Requirements</span></span>

|<span data-ttu-id="dfaf4-228">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfaf4-228">Requirement</span></span>| <span data-ttu-id="dfaf4-229">Valor</span><span class="sxs-lookup"><span data-stu-id="dfaf4-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfaf4-230">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfaf4-230">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfaf4-231">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-231">1.1</span></span>|
|[<span data-ttu-id="dfaf4-232">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfaf4-232">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfaf4-233">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfaf4-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfaf4-234">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-234">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="dfaf4-235">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="dfaf4-235">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="dfaf4-236">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-236">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="dfaf4-237">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-237">Type</span></span>

*   [<span data-ttu-id="dfaf4-238">PlatformType</span><span class="sxs-lookup"><span data-stu-id="dfaf4-238">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="dfaf4-239">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfaf4-239">Requirements</span></span>

|<span data-ttu-id="dfaf4-240">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfaf4-240">Requirement</span></span>| <span data-ttu-id="dfaf4-241">Valor</span><span class="sxs-lookup"><span data-stu-id="dfaf4-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfaf4-242">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfaf4-242">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfaf4-243">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-243">1.1</span></span>|
|[<span data-ttu-id="dfaf4-244">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfaf4-244">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfaf4-245">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfaf4-245">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfaf4-246">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-246">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="dfaf4-247">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="dfaf4-247">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="dfaf4-248">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o aplicativo atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-248">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="dfaf4-249">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-249">Type</span></span>

*   [<span data-ttu-id="dfaf4-250">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="dfaf4-250">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="dfaf4-251">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfaf4-251">Requirements</span></span>

|<span data-ttu-id="dfaf4-252">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfaf4-252">Requirement</span></span>| <span data-ttu-id="dfaf4-253">Valor</span><span class="sxs-lookup"><span data-stu-id="dfaf4-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfaf4-254">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfaf4-254">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfaf4-255">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-255">1.1</span></span>|
|[<span data-ttu-id="dfaf4-256">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfaf4-256">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfaf4-257">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfaf4-257">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dfaf4-258">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-258">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="dfaf4-259">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="dfaf4-259">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="dfaf4-260">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-260">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="dfaf4-261">O `RoamingSettings` objeto permite armazenar e acessar dados de um suplemento de email armazenado na caixa de correio de um usuário, para que esteja disponível para esse suplemento quando ele estiver sendo executado a partir de qualquer cliente do Outlook usado para acessar a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-261">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="dfaf4-262">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-262">Type</span></span>

*   [<span data-ttu-id="dfaf4-263">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="dfaf4-263">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="dfaf4-264">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfaf4-264">Requirements</span></span>

|<span data-ttu-id="dfaf4-265">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfaf4-265">Requirement</span></span>| <span data-ttu-id="dfaf4-266">Valor</span><span class="sxs-lookup"><span data-stu-id="dfaf4-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfaf4-267">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfaf4-267">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfaf4-268">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-268">1.1</span></span>|
|[<span data-ttu-id="dfaf4-269">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-269">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="dfaf4-270">Restrito</span><span class="sxs-lookup"><span data-stu-id="dfaf4-270">Restricted</span></span>|
|[<span data-ttu-id="dfaf4-271">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfaf4-271">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfaf4-272">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfaf4-272">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="dfaf4-273">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="dfaf4-273">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="dfaf4-274">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="dfaf4-274">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="dfaf4-275">Tipo</span><span class="sxs-lookup"><span data-stu-id="dfaf4-275">Type</span></span>

*   [<span data-ttu-id="dfaf4-276">UI</span><span class="sxs-lookup"><span data-stu-id="dfaf4-276">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="dfaf4-277">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dfaf4-277">Requirements</span></span>

|<span data-ttu-id="dfaf4-278">Requisito</span><span class="sxs-lookup"><span data-stu-id="dfaf4-278">Requirement</span></span>| <span data-ttu-id="dfaf4-279">Valor</span><span class="sxs-lookup"><span data-stu-id="dfaf4-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="dfaf4-280">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dfaf4-280">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dfaf4-281">1.1</span><span class="sxs-lookup"><span data-stu-id="dfaf4-281">1.1</span></span>|
|[<span data-ttu-id="dfaf4-282">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dfaf4-282">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dfaf4-283">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dfaf4-283">Compose or Read</span></span>|
