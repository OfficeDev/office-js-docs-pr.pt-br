---
title: Office. Context – conjunto de requisitos de visualização
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de visualização da API da caixa de correio.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 5987f81b0b4790b74bde092fc3de44df4fa3ed16
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293811"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="440ff-103">contexto (conjunto de requisitos de visualização da caixa de correio)</span><span class="sxs-lookup"><span data-stu-id="440ff-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="440ff-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="440ff-104">[Office](office.md).context</span></span>

<span data-ttu-id="440ff-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="440ff-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="440ff-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-preview).</span><span class="sxs-lookup"><span data-stu-id="440ff-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="440ff-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="440ff-107">Requirements</span></span>

|<span data-ttu-id="440ff-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="440ff-108">Requirement</span></span>| <span data-ttu-id="440ff-109">Valor</span><span class="sxs-lookup"><span data-stu-id="440ff-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="440ff-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="440ff-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="440ff-111">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-111">1.1</span></span>|
|[<span data-ttu-id="440ff-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="440ff-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="440ff-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="440ff-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="440ff-114">Properties</span></span>

| <span data-ttu-id="440ff-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="440ff-115">Property</span></span> | <span data-ttu-id="440ff-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="440ff-116">Modes</span></span> | <span data-ttu-id="440ff-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="440ff-117">Return type</span></span> | <span data-ttu-id="440ff-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="440ff-118">Minimum</span></span><br><span data-ttu-id="440ff-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="440ff-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="440ff-120">autentica</span><span class="sxs-lookup"><span data-stu-id="440ff-120">auth</span></span>](#auth-auth) | <span data-ttu-id="440ff-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="440ff-121">Compose</span></span><br><span data-ttu-id="440ff-122">Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-122">Read</span></span> | [<span data-ttu-id="440ff-123">Auth</span><span class="sxs-lookup"><span data-stu-id="440ff-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="440ff-124">Visualização</span><span class="sxs-lookup"><span data-stu-id="440ff-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="440ff-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="440ff-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="440ff-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="440ff-126">Compose</span></span><br><span data-ttu-id="440ff-127">Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-127">Read</span></span> | <span data-ttu-id="440ff-128">String</span><span class="sxs-lookup"><span data-stu-id="440ff-128">String</span></span> | [<span data-ttu-id="440ff-129">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="440ff-130">la</span><span class="sxs-lookup"><span data-stu-id="440ff-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="440ff-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="440ff-131">Compose</span></span><br><span data-ttu-id="440ff-132">Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-132">Read</span></span> | [<span data-ttu-id="440ff-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="440ff-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="440ff-134">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="440ff-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="440ff-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="440ff-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="440ff-136">Compose</span></span><br><span data-ttu-id="440ff-137">Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-137">Read</span></span> | <span data-ttu-id="440ff-138">String</span><span class="sxs-lookup"><span data-stu-id="440ff-138">String</span></span> | [<span data-ttu-id="440ff-139">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="440ff-140">principal</span><span class="sxs-lookup"><span data-stu-id="440ff-140">host</span></span>](#host-hosttype) | <span data-ttu-id="440ff-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="440ff-141">Compose</span></span><br><span data-ttu-id="440ff-142">Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-142">Read</span></span> | [<span data-ttu-id="440ff-143">HostType</span><span class="sxs-lookup"><span data-stu-id="440ff-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="440ff-144">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="440ff-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="440ff-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="440ff-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="440ff-146">Compose</span></span><br><span data-ttu-id="440ff-147">Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-147">Read</span></span> | [<span data-ttu-id="440ff-148">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="440ff-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="440ff-149">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="440ff-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="440ff-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="440ff-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="440ff-151">Compose</span></span><br><span data-ttu-id="440ff-152">Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-152">Read</span></span> | [<span data-ttu-id="440ff-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="440ff-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="440ff-154">Visualização</span><span class="sxs-lookup"><span data-stu-id="440ff-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="440ff-155">plataforma</span><span class="sxs-lookup"><span data-stu-id="440ff-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="440ff-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="440ff-156">Compose</span></span><br><span data-ttu-id="440ff-157">Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-157">Read</span></span> | [<span data-ttu-id="440ff-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="440ff-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="440ff-159">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="440ff-160">atende</span><span class="sxs-lookup"><span data-stu-id="440ff-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="440ff-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="440ff-161">Compose</span></span><br><span data-ttu-id="440ff-162">Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-162">Read</span></span> | [<span data-ttu-id="440ff-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="440ff-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="440ff-164">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="440ff-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="440ff-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="440ff-166">Escrever</span><span class="sxs-lookup"><span data-stu-id="440ff-166">Compose</span></span><br><span data-ttu-id="440ff-167">Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-167">Read</span></span> | [<span data-ttu-id="440ff-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="440ff-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="440ff-169">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="440ff-170">ui</span><span class="sxs-lookup"><span data-stu-id="440ff-170">ui</span></span>](#ui-ui) | <span data-ttu-id="440ff-171">Escrever</span><span class="sxs-lookup"><span data-stu-id="440ff-171">Compose</span></span><br><span data-ttu-id="440ff-172">Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-172">Read</span></span> | [<span data-ttu-id="440ff-173">UI</span><span class="sxs-lookup"><span data-stu-id="440ff-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="440ff-174">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="440ff-175">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="440ff-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="440ff-176">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="440ff-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="440ff-177">Oferece suporte a [logon único (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , fornecendo um método que permite que o aplicativo do Office obtenha um token de acesso para o aplicativo Web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="440ff-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="440ff-178">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="440ff-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="440ff-179">Tipo</span><span class="sxs-lookup"><span data-stu-id="440ff-179">Type</span></span>

*   [<span data-ttu-id="440ff-180">Auth</span><span class="sxs-lookup"><span data-stu-id="440ff-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="440ff-181">Requisitos</span><span class="sxs-lookup"><span data-stu-id="440ff-181">Requirements</span></span>

|<span data-ttu-id="440ff-182">Requisito</span><span class="sxs-lookup"><span data-stu-id="440ff-182">Requirement</span></span>| <span data-ttu-id="440ff-183">Valor</span><span class="sxs-lookup"><span data-stu-id="440ff-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="440ff-184">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="440ff-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="440ff-185">Visualização</span><span class="sxs-lookup"><span data-stu-id="440ff-185">Preview</span></span>|
|[<span data-ttu-id="440ff-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="440ff-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="440ff-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="440ff-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="440ff-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="440ff-189">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="440ff-189">contentLanguage: String</span></span>

<span data-ttu-id="440ff-190">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="440ff-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="440ff-191">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="440ff-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="440ff-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="440ff-192">Type</span></span>

*   <span data-ttu-id="440ff-193">String</span><span class="sxs-lookup"><span data-stu-id="440ff-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="440ff-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="440ff-194">Requirements</span></span>

|<span data-ttu-id="440ff-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="440ff-195">Requirement</span></span>| <span data-ttu-id="440ff-196">Valor</span><span class="sxs-lookup"><span data-stu-id="440ff-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="440ff-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="440ff-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="440ff-198">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-198">1.1</span></span>|
|[<span data-ttu-id="440ff-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="440ff-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="440ff-200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="440ff-201">Exemplo</span><span class="sxs-lookup"><span data-stu-id="440ff-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="440ff-202">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="440ff-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="440ff-203">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="440ff-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="440ff-204">Tipo</span><span class="sxs-lookup"><span data-stu-id="440ff-204">Type</span></span>

*   [<span data-ttu-id="440ff-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="440ff-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="440ff-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="440ff-206">Requirements</span></span>

|<span data-ttu-id="440ff-207">Requisito</span><span class="sxs-lookup"><span data-stu-id="440ff-207">Requirement</span></span>| <span data-ttu-id="440ff-208">Valor</span><span class="sxs-lookup"><span data-stu-id="440ff-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="440ff-209">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="440ff-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="440ff-210">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-210">1.1</span></span>|
|[<span data-ttu-id="440ff-211">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="440ff-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="440ff-212">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="440ff-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="440ff-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="440ff-214">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="440ff-214">displayLanguage: String</span></span>

<span data-ttu-id="440ff-215">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="440ff-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="440ff-216">O `displayLanguage` valor reflete a configuração de **idioma de exibição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="440ff-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="440ff-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="440ff-217">Type</span></span>

*   <span data-ttu-id="440ff-218">String</span><span class="sxs-lookup"><span data-stu-id="440ff-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="440ff-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="440ff-219">Requirements</span></span>

|<span data-ttu-id="440ff-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="440ff-220">Requirement</span></span>| <span data-ttu-id="440ff-221">Valor</span><span class="sxs-lookup"><span data-stu-id="440ff-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="440ff-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="440ff-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="440ff-223">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-223">1.1</span></span>|
|[<span data-ttu-id="440ff-224">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="440ff-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="440ff-225">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="440ff-226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="440ff-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="440ff-227">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="440ff-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="440ff-228">Obtém o aplicativo do Office que está hospedando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="440ff-228">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="440ff-229">Tipo</span><span class="sxs-lookup"><span data-stu-id="440ff-229">Type</span></span>

*   [<span data-ttu-id="440ff-230">HostType</span><span class="sxs-lookup"><span data-stu-id="440ff-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="440ff-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="440ff-231">Requirements</span></span>

|<span data-ttu-id="440ff-232">Requisito</span><span class="sxs-lookup"><span data-stu-id="440ff-232">Requirement</span></span>| <span data-ttu-id="440ff-233">Valor</span><span class="sxs-lookup"><span data-stu-id="440ff-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="440ff-234">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="440ff-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="440ff-235">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-235">1.1</span></span>|
|[<span data-ttu-id="440ff-236">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="440ff-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="440ff-237">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="440ff-238">Exemplo</span><span class="sxs-lookup"><span data-stu-id="440ff-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="440ff-239">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="440ff-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="440ff-240">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="440ff-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="440ff-241">Só há suporte para esse membro no Outlook no Windows.</span><span class="sxs-lookup"><span data-stu-id="440ff-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="440ff-242">O uso de cores de temas do Office permite coordenar o esquema de cores do seu suplemento com o tema atual do Office selecionado pelo usuário com a **conta de arquivo > office > Office Theme UI**, que é aplicada em todos os aplicativos cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="440ff-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="440ff-243">Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="440ff-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="440ff-244">Tipo</span><span class="sxs-lookup"><span data-stu-id="440ff-244">Type</span></span>

*   [<span data-ttu-id="440ff-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="440ff-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="440ff-246">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="440ff-246">Properties:</span></span>

|<span data-ttu-id="440ff-247">Nome</span><span class="sxs-lookup"><span data-stu-id="440ff-247">Name</span></span>| <span data-ttu-id="440ff-248">Tipo</span><span class="sxs-lookup"><span data-stu-id="440ff-248">Type</span></span>| <span data-ttu-id="440ff-249">Descrição</span><span class="sxs-lookup"><span data-stu-id="440ff-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="440ff-250">String</span><span class="sxs-lookup"><span data-stu-id="440ff-250">String</span></span>|<span data-ttu-id="440ff-251">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="440ff-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="440ff-252">String</span><span class="sxs-lookup"><span data-stu-id="440ff-252">String</span></span>|<span data-ttu-id="440ff-253">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="440ff-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="440ff-254">String</span><span class="sxs-lookup"><span data-stu-id="440ff-254">String</span></span>|<span data-ttu-id="440ff-255">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="440ff-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="440ff-256">String</span><span class="sxs-lookup"><span data-stu-id="440ff-256">String</span></span>|<span data-ttu-id="440ff-257">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="440ff-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="440ff-258">Requisitos</span><span class="sxs-lookup"><span data-stu-id="440ff-258">Requirements</span></span>

|<span data-ttu-id="440ff-259">Requisito</span><span class="sxs-lookup"><span data-stu-id="440ff-259">Requirement</span></span>| <span data-ttu-id="440ff-260">Valor</span><span class="sxs-lookup"><span data-stu-id="440ff-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="440ff-261">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="440ff-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="440ff-262">Visualização</span><span class="sxs-lookup"><span data-stu-id="440ff-262">Preview</span></span>|
|[<span data-ttu-id="440ff-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="440ff-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="440ff-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="440ff-265">Exemplo</span><span class="sxs-lookup"><span data-stu-id="440ff-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="440ff-266">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="440ff-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="440ff-267">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="440ff-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="440ff-268">Tipo</span><span class="sxs-lookup"><span data-stu-id="440ff-268">Type</span></span>

*   [<span data-ttu-id="440ff-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="440ff-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="440ff-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="440ff-270">Requirements</span></span>

|<span data-ttu-id="440ff-271">Requisito</span><span class="sxs-lookup"><span data-stu-id="440ff-271">Requirement</span></span>| <span data-ttu-id="440ff-272">Valor</span><span class="sxs-lookup"><span data-stu-id="440ff-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="440ff-273">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="440ff-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="440ff-274">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-274">1.1</span></span>|
|[<span data-ttu-id="440ff-275">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="440ff-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="440ff-276">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="440ff-277">Exemplo</span><span class="sxs-lookup"><span data-stu-id="440ff-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="440ff-278">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="440ff-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="440ff-279">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o aplicativo atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="440ff-279">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="440ff-280">Tipo</span><span class="sxs-lookup"><span data-stu-id="440ff-280">Type</span></span>

*   [<span data-ttu-id="440ff-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="440ff-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="440ff-282">Requisitos</span><span class="sxs-lookup"><span data-stu-id="440ff-282">Requirements</span></span>

|<span data-ttu-id="440ff-283">Requisito</span><span class="sxs-lookup"><span data-stu-id="440ff-283">Requirement</span></span>| <span data-ttu-id="440ff-284">Valor</span><span class="sxs-lookup"><span data-stu-id="440ff-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="440ff-285">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="440ff-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="440ff-286">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-286">1.1</span></span>|
|[<span data-ttu-id="440ff-287">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="440ff-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="440ff-288">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="440ff-289">Exemplo</span><span class="sxs-lookup"><span data-stu-id="440ff-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="440ff-290">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="440ff-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="440ff-291">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="440ff-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="440ff-292">O `RoamingSettings` objeto permite armazenar e acessar dados de um suplemento de email armazenado na caixa de correio de um usuário, para que esteja disponível para esse suplemento quando ele estiver sendo executado a partir de qualquer cliente do Outlook usado para acessar a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="440ff-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="440ff-293">Tipo</span><span class="sxs-lookup"><span data-stu-id="440ff-293">Type</span></span>

*   [<span data-ttu-id="440ff-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="440ff-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="440ff-295">Requisitos</span><span class="sxs-lookup"><span data-stu-id="440ff-295">Requirements</span></span>

|<span data-ttu-id="440ff-296">Requisito</span><span class="sxs-lookup"><span data-stu-id="440ff-296">Requirement</span></span>| <span data-ttu-id="440ff-297">Valor</span><span class="sxs-lookup"><span data-stu-id="440ff-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="440ff-298">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="440ff-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="440ff-299">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-299">1.1</span></span>|
|[<span data-ttu-id="440ff-300">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="440ff-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="440ff-301">Restrito</span><span class="sxs-lookup"><span data-stu-id="440ff-301">Restricted</span></span>|
|[<span data-ttu-id="440ff-302">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="440ff-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="440ff-303">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="440ff-304">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="440ff-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="440ff-305">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="440ff-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="440ff-306">Tipo</span><span class="sxs-lookup"><span data-stu-id="440ff-306">Type</span></span>

*   [<span data-ttu-id="440ff-307">UI</span><span class="sxs-lookup"><span data-stu-id="440ff-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="440ff-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="440ff-308">Requirements</span></span>

|<span data-ttu-id="440ff-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="440ff-309">Requirement</span></span>| <span data-ttu-id="440ff-310">Valor</span><span class="sxs-lookup"><span data-stu-id="440ff-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="440ff-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="440ff-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="440ff-312">1.1</span><span class="sxs-lookup"><span data-stu-id="440ff-312">1.1</span></span>|
|[<span data-ttu-id="440ff-313">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="440ff-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="440ff-314">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="440ff-314">Compose or Read</span></span>|
