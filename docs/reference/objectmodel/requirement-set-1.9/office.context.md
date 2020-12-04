---
title: Office. Context – conjunto de requisitos 1,9
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,9.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 3a8a9fe65ebf3c5a5ee63766f71dfce8e3f8d905
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570720"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="8f2e5-103">contexto (conjunto de requisitos de caixa de correio 1,9)</span><span class="sxs-lookup"><span data-stu-id="8f2e5-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="8f2e5-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="8f2e5-104">[Office](office.md).context</span></span>

<span data-ttu-id="8f2e5-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="8f2e5-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="8f2e5-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f2e5-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f2e5-107">Requirements</span></span>

|<span data-ttu-id="8f2e5-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f2e5-108">Requirement</span></span>| <span data-ttu-id="8f2e5-109">Valor</span><span class="sxs-lookup"><span data-stu-id="8f2e5-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f2e5-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f2e5-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f2e5-111">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-111">1.1</span></span>|
|[<span data-ttu-id="8f2e5-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f2e5-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f2e5-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8f2e5-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="8f2e5-114">Properties</span></span>

| <span data-ttu-id="8f2e5-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="8f2e5-115">Property</span></span> | <span data-ttu-id="8f2e5-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="8f2e5-116">Modes</span></span> | <span data-ttu-id="8f2e5-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="8f2e5-117">Return type</span></span> | <span data-ttu-id="8f2e5-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="8f2e5-118">Minimum</span></span><br><span data-ttu-id="8f2e5-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="8f2e5-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8f2e5-120">autentica</span><span class="sxs-lookup"><span data-stu-id="8f2e5-120">auth</span></span>](#auth-auth) | <span data-ttu-id="8f2e5-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="8f2e5-121">Compose</span></span><br><span data-ttu-id="8f2e5-122">Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-122">Read</span></span> | [<span data-ttu-id="8f2e5-123">Auth</span><span class="sxs-lookup"><span data-stu-id="8f2e5-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="8f2e5-124">IdentityAPI 1,3</span><span class="sxs-lookup"><span data-stu-id="8f2e5-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="8f2e5-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="8f2e5-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="8f2e5-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="8f2e5-126">Compose</span></span><br><span data-ttu-id="8f2e5-127">Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-127">Read</span></span> | <span data-ttu-id="8f2e5-128">String</span><span class="sxs-lookup"><span data-stu-id="8f2e5-128">String</span></span> | [<span data-ttu-id="8f2e5-129">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f2e5-130">la</span><span class="sxs-lookup"><span data-stu-id="8f2e5-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="8f2e5-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="8f2e5-131">Compose</span></span><br><span data-ttu-id="8f2e5-132">Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-132">Read</span></span> | [<span data-ttu-id="8f2e5-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="8f2e5-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="8f2e5-134">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f2e5-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="8f2e5-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="8f2e5-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="8f2e5-136">Compose</span></span><br><span data-ttu-id="8f2e5-137">Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-137">Read</span></span> | <span data-ttu-id="8f2e5-138">String</span><span class="sxs-lookup"><span data-stu-id="8f2e5-138">String</span></span> | [<span data-ttu-id="8f2e5-139">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f2e5-140">principal</span><span class="sxs-lookup"><span data-stu-id="8f2e5-140">host</span></span>](#host-hosttype) | <span data-ttu-id="8f2e5-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="8f2e5-141">Compose</span></span><br><span data-ttu-id="8f2e5-142">Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-142">Read</span></span> | [<span data-ttu-id="8f2e5-143">HostType</span><span class="sxs-lookup"><span data-stu-id="8f2e5-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="8f2e5-144">1,5</span><span class="sxs-lookup"><span data-stu-id="8f2e5-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="8f2e5-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="8f2e5-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="8f2e5-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="8f2e5-146">Compose</span></span><br><span data-ttu-id="8f2e5-147">Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-147">Read</span></span> | [<span data-ttu-id="8f2e5-148">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="8f2e5-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="8f2e5-149">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f2e5-150">plataforma</span><span class="sxs-lookup"><span data-stu-id="8f2e5-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="8f2e5-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="8f2e5-151">Compose</span></span><br><span data-ttu-id="8f2e5-152">Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-152">Read</span></span> | [<span data-ttu-id="8f2e5-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="8f2e5-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="8f2e5-154">1,5</span><span class="sxs-lookup"><span data-stu-id="8f2e5-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="8f2e5-155">atende</span><span class="sxs-lookup"><span data-stu-id="8f2e5-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="8f2e5-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="8f2e5-156">Compose</span></span><br><span data-ttu-id="8f2e5-157">Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-157">Read</span></span> | [<span data-ttu-id="8f2e5-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="8f2e5-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="8f2e5-159">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f2e5-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="8f2e5-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="8f2e5-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="8f2e5-161">Compose</span></span><br><span data-ttu-id="8f2e5-162">Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-162">Read</span></span> | [<span data-ttu-id="8f2e5-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8f2e5-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="8f2e5-164">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8f2e5-165">ui</span><span class="sxs-lookup"><span data-stu-id="8f2e5-165">ui</span></span>](#ui-ui) | <span data-ttu-id="8f2e5-166">Escrever</span><span class="sxs-lookup"><span data-stu-id="8f2e5-166">Compose</span></span><br><span data-ttu-id="8f2e5-167">Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-167">Read</span></span> | [<span data-ttu-id="8f2e5-168">UI</span><span class="sxs-lookup"><span data-stu-id="8f2e5-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="8f2e5-169">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="8f2e5-170">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="8f2e5-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="8f2e5-171">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="8f2e5-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="8f2e5-172">Oferece suporte a [logon único (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , fornecendo um método que permite que o aplicativo do Office obtenha um token de acesso para o aplicativo Web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="8f2e5-173">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="8f2e5-174">Confira [IdentityAPI 1,3 conjunto de requisitos](../../requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="8f2e5-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="8f2e5-175">Tipo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-175">Type</span></span>

*   [<span data-ttu-id="8f2e5-176">Auth</span><span class="sxs-lookup"><span data-stu-id="8f2e5-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="8f2e5-177">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f2e5-177">Requirements</span></span>

|<span data-ttu-id="8f2e5-178">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f2e5-178">Requirement</span></span>| <span data-ttu-id="8f2e5-179">Valor</span><span class="sxs-lookup"><span data-stu-id="8f2e5-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f2e5-180">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f2e5-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f2e5-181">N/D</span><span class="sxs-lookup"><span data-stu-id="8f2e5-181">N/A</span></span>|
|[<span data-ttu-id="8f2e5-182">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f2e5-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f2e5-183">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f2e5-184">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-184">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="8f2e5-185">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8f2e5-185">contentLanguage: String</span></span>

<span data-ttu-id="8f2e5-186">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="8f2e5-187">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="8f2e5-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-188">Type</span></span>

*   <span data-ttu-id="8f2e5-189">String</span><span class="sxs-lookup"><span data-stu-id="8f2e5-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f2e5-190">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f2e5-190">Requirements</span></span>

|<span data-ttu-id="8f2e5-191">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f2e5-191">Requirement</span></span>| <span data-ttu-id="8f2e5-192">Valor</span><span class="sxs-lookup"><span data-stu-id="8f2e5-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f2e5-193">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f2e5-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f2e5-194">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-194">1.1</span></span>|
|[<span data-ttu-id="8f2e5-195">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f2e5-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f2e5-196">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f2e5-197">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-197">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="8f2e5-198">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="8f2e5-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="8f2e5-199">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="8f2e5-200">Tipo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-200">Type</span></span>

*   [<span data-ttu-id="8f2e5-201">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="8f2e5-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="8f2e5-202">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f2e5-202">Requirements</span></span>

|<span data-ttu-id="8f2e5-203">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f2e5-203">Requirement</span></span>| <span data-ttu-id="8f2e5-204">Valor</span><span class="sxs-lookup"><span data-stu-id="8f2e5-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f2e5-205">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f2e5-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f2e5-206">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-206">1.1</span></span>|
|[<span data-ttu-id="8f2e5-207">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f2e5-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f2e5-208">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f2e5-209">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-209">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="8f2e5-210">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8f2e5-210">displayLanguage: String</span></span>

<span data-ttu-id="8f2e5-211">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="8f2e5-212">O `displayLanguage` valor reflete a configuração de **idioma de exibição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="8f2e5-213">Tipo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-213">Type</span></span>

*   <span data-ttu-id="8f2e5-214">String</span><span class="sxs-lookup"><span data-stu-id="8f2e5-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8f2e5-215">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f2e5-215">Requirements</span></span>

|<span data-ttu-id="8f2e5-216">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f2e5-216">Requirement</span></span>| <span data-ttu-id="8f2e5-217">Valor</span><span class="sxs-lookup"><span data-stu-id="8f2e5-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f2e5-218">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f2e5-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f2e5-219">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-219">1.1</span></span>|
|[<span data-ttu-id="8f2e5-220">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f2e5-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f2e5-221">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f2e5-222">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-222">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="8f2e5-223">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="8f2e5-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="8f2e5-224">Obtém o aplicativo do Office que está hospedando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-224">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="8f2e5-225">Como alternativa, você pode usar a propriedade [Office. Context. Diagnostics](#diagnostics-contextinformation) para obter a plataforma.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-225">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="8f2e5-226">Tipo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-226">Type</span></span>

*   [<span data-ttu-id="8f2e5-227">HostType</span><span class="sxs-lookup"><span data-stu-id="8f2e5-227">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="8f2e5-228">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f2e5-228">Requirements</span></span>

|<span data-ttu-id="8f2e5-229">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f2e5-229">Requirement</span></span>| <span data-ttu-id="8f2e5-230">Valor</span><span class="sxs-lookup"><span data-stu-id="8f2e5-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f2e5-231">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f2e5-231">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f2e5-232">1,5</span><span class="sxs-lookup"><span data-stu-id="8f2e5-232">1.5</span></span>|
|[<span data-ttu-id="8f2e5-233">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f2e5-233">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f2e5-234">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-234">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f2e5-235">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-235">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="8f2e5-236">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="8f2e5-236">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="8f2e5-237">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-237">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="8f2e5-238">Como alternativa, você pode usar a propriedade [Office. Context. Diagnostics](#diagnostics-contextinformation) para obter a plataforma.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-238">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="8f2e5-239">Tipo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-239">Type</span></span>

*   [<span data-ttu-id="8f2e5-240">PlatformType</span><span class="sxs-lookup"><span data-stu-id="8f2e5-240">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="8f2e5-241">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f2e5-241">Requirements</span></span>

|<span data-ttu-id="8f2e5-242">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f2e5-242">Requirement</span></span>| <span data-ttu-id="8f2e5-243">Valor</span><span class="sxs-lookup"><span data-stu-id="8f2e5-243">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f2e5-244">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f2e5-244">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f2e5-245">1,5</span><span class="sxs-lookup"><span data-stu-id="8f2e5-245">1.5</span></span>|
|[<span data-ttu-id="8f2e5-246">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f2e5-246">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f2e5-247">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-247">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f2e5-248">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-248">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="8f2e5-249">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="8f2e5-249">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="8f2e5-250">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o aplicativo atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-250">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="8f2e5-251">Tipo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-251">Type</span></span>

*   [<span data-ttu-id="8f2e5-252">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="8f2e5-252">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="8f2e5-253">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f2e5-253">Requirements</span></span>

|<span data-ttu-id="8f2e5-254">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f2e5-254">Requirement</span></span>| <span data-ttu-id="8f2e5-255">Valor</span><span class="sxs-lookup"><span data-stu-id="8f2e5-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f2e5-256">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f2e5-256">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f2e5-257">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-257">1.1</span></span>|
|[<span data-ttu-id="8f2e5-258">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f2e5-258">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f2e5-259">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8f2e5-260">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-260">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="8f2e5-261">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="8f2e5-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="8f2e5-262">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-262">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="8f2e5-263">O `RoamingSettings` objeto permite armazenar e acessar dados de um suplemento de email armazenado na caixa de correio de um usuário, para que esteja disponível para esse suplemento quando ele estiver sendo executado a partir de qualquer cliente do Outlook usado para acessar a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-263">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="8f2e5-264">Tipo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-264">Type</span></span>

*   [<span data-ttu-id="8f2e5-265">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8f2e5-265">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="8f2e5-266">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f2e5-266">Requirements</span></span>

|<span data-ttu-id="8f2e5-267">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f2e5-267">Requirement</span></span>| <span data-ttu-id="8f2e5-268">Valor</span><span class="sxs-lookup"><span data-stu-id="8f2e5-268">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f2e5-269">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f2e5-269">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f2e5-270">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-270">1.1</span></span>|
|[<span data-ttu-id="8f2e5-271">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-271">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="8f2e5-272">Restrito</span><span class="sxs-lookup"><span data-stu-id="8f2e5-272">Restricted</span></span>|
|[<span data-ttu-id="8f2e5-273">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f2e5-273">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f2e5-274">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-274">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="8f2e5-275">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="8f2e5-275">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="8f2e5-276">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="8f2e5-276">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="8f2e5-277">Tipo</span><span class="sxs-lookup"><span data-stu-id="8f2e5-277">Type</span></span>

*   [<span data-ttu-id="8f2e5-278">UI</span><span class="sxs-lookup"><span data-stu-id="8f2e5-278">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="8f2e5-279">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8f2e5-279">Requirements</span></span>

|<span data-ttu-id="8f2e5-280">Requisito</span><span class="sxs-lookup"><span data-stu-id="8f2e5-280">Requirement</span></span>| <span data-ttu-id="8f2e5-281">Valor</span><span class="sxs-lookup"><span data-stu-id="8f2e5-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="8f2e5-282">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8f2e5-282">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8f2e5-283">1.1</span><span class="sxs-lookup"><span data-stu-id="8f2e5-283">1.1</span></span>|
|[<span data-ttu-id="8f2e5-284">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8f2e5-284">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8f2e5-285">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8f2e5-285">Compose or Read</span></span>|
