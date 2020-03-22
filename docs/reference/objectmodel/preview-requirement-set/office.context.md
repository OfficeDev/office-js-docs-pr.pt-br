---
title: Office. Context – conjunto de requisitos de visualização
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de visualização da API da caixa de correio.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: c61769cb1ae98097ffabb8b3ef19b2f82257c2b1
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890862"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="16951-103">contexto (conjunto de requisitos de visualização da caixa de correio)</span><span class="sxs-lookup"><span data-stu-id="16951-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="16951-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="16951-104">[Office](office.md).context</span></span>

<span data-ttu-id="16951-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="16951-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="16951-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-preview).</span><span class="sxs-lookup"><span data-stu-id="16951-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="16951-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="16951-107">Requirements</span></span>

|<span data-ttu-id="16951-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="16951-108">Requirement</span></span>| <span data-ttu-id="16951-109">Valor</span><span class="sxs-lookup"><span data-stu-id="16951-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="16951-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="16951-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16951-111">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-111">1.1</span></span>|
|[<span data-ttu-id="16951-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="16951-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16951-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="16951-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="16951-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="16951-114">Properties</span></span>

| <span data-ttu-id="16951-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="16951-115">Property</span></span> | <span data-ttu-id="16951-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="16951-116">Modes</span></span> | <span data-ttu-id="16951-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="16951-117">Return type</span></span> | <span data-ttu-id="16951-118">Mínimo</span><span class="sxs-lookup"><span data-stu-id="16951-118">Minimum</span></span><br><span data-ttu-id="16951-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="16951-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="16951-120">autentica</span><span class="sxs-lookup"><span data-stu-id="16951-120">auth</span></span>](#auth-auth) | <span data-ttu-id="16951-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="16951-121">Compose</span></span><br><span data-ttu-id="16951-122">Ler</span><span class="sxs-lookup"><span data-stu-id="16951-122">Read</span></span> | [<span data-ttu-id="16951-123">Auth</span><span class="sxs-lookup"><span data-stu-id="16951-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="16951-124">Visualização</span><span class="sxs-lookup"><span data-stu-id="16951-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="16951-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="16951-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="16951-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="16951-126">Compose</span></span><br><span data-ttu-id="16951-127">Ler</span><span class="sxs-lookup"><span data-stu-id="16951-127">Read</span></span> | <span data-ttu-id="16951-128">String</span><span class="sxs-lookup"><span data-stu-id="16951-128">String</span></span> | [<span data-ttu-id="16951-129">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="16951-130">la</span><span class="sxs-lookup"><span data-stu-id="16951-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="16951-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="16951-131">Compose</span></span><br><span data-ttu-id="16951-132">Ler</span><span class="sxs-lookup"><span data-stu-id="16951-132">Read</span></span> | [<span data-ttu-id="16951-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="16951-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="16951-134">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="16951-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="16951-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="16951-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="16951-136">Compose</span></span><br><span data-ttu-id="16951-137">Ler</span><span class="sxs-lookup"><span data-stu-id="16951-137">Read</span></span> | <span data-ttu-id="16951-138">String</span><span class="sxs-lookup"><span data-stu-id="16951-138">String</span></span> | [<span data-ttu-id="16951-139">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="16951-140">principal</span><span class="sxs-lookup"><span data-stu-id="16951-140">host</span></span>](#host-hosttype) | <span data-ttu-id="16951-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="16951-141">Compose</span></span><br><span data-ttu-id="16951-142">Ler</span><span class="sxs-lookup"><span data-stu-id="16951-142">Read</span></span> | [<span data-ttu-id="16951-143">HostType</span><span class="sxs-lookup"><span data-stu-id="16951-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="16951-144">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="16951-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="16951-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="16951-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="16951-146">Compose</span></span><br><span data-ttu-id="16951-147">Ler</span><span class="sxs-lookup"><span data-stu-id="16951-147">Read</span></span> | [<span data-ttu-id="16951-148">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="16951-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="16951-149">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="16951-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="16951-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="16951-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="16951-151">Compose</span></span><br><span data-ttu-id="16951-152">Ler</span><span class="sxs-lookup"><span data-stu-id="16951-152">Read</span></span> | [<span data-ttu-id="16951-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="16951-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="16951-154">Visualização</span><span class="sxs-lookup"><span data-stu-id="16951-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="16951-155">plataforma</span><span class="sxs-lookup"><span data-stu-id="16951-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="16951-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="16951-156">Compose</span></span><br><span data-ttu-id="16951-157">Ler</span><span class="sxs-lookup"><span data-stu-id="16951-157">Read</span></span> | [<span data-ttu-id="16951-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="16951-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="16951-159">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="16951-160">atende</span><span class="sxs-lookup"><span data-stu-id="16951-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="16951-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="16951-161">Compose</span></span><br><span data-ttu-id="16951-162">Ler</span><span class="sxs-lookup"><span data-stu-id="16951-162">Read</span></span> | [<span data-ttu-id="16951-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="16951-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="16951-164">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="16951-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="16951-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="16951-166">Escrever</span><span class="sxs-lookup"><span data-stu-id="16951-166">Compose</span></span><br><span data-ttu-id="16951-167">Ler</span><span class="sxs-lookup"><span data-stu-id="16951-167">Read</span></span> | [<span data-ttu-id="16951-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="16951-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="16951-169">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="16951-170">ui</span><span class="sxs-lookup"><span data-stu-id="16951-170">ui</span></span>](#ui-ui) | <span data-ttu-id="16951-171">Escrever</span><span class="sxs-lookup"><span data-stu-id="16951-171">Compose</span></span><br><span data-ttu-id="16951-172">Ler</span><span class="sxs-lookup"><span data-stu-id="16951-172">Read</span></span> | [<span data-ttu-id="16951-173">UI</span><span class="sxs-lookup"><span data-stu-id="16951-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="16951-174">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="16951-175">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="16951-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="16951-176">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="16951-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="16951-177">Oferece suporte a [logon único (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , fornecendo um método que permite que o host do Office obtenha um token de acesso para o aplicativo Web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="16951-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="16951-178">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="16951-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="16951-179">Tipo</span><span class="sxs-lookup"><span data-stu-id="16951-179">Type</span></span>

*   [<span data-ttu-id="16951-180">Auth</span><span class="sxs-lookup"><span data-stu-id="16951-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="16951-181">Requisitos</span><span class="sxs-lookup"><span data-stu-id="16951-181">Requirements</span></span>

|<span data-ttu-id="16951-182">Requisito</span><span class="sxs-lookup"><span data-stu-id="16951-182">Requirement</span></span>| <span data-ttu-id="16951-183">Valor</span><span class="sxs-lookup"><span data-stu-id="16951-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="16951-184">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="16951-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16951-185">Visualização</span><span class="sxs-lookup"><span data-stu-id="16951-185">Preview</span></span>|
|[<span data-ttu-id="16951-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="16951-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16951-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="16951-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="16951-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="16951-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="16951-189">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="16951-189">contentLanguage: String</span></span>

<span data-ttu-id="16951-190">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="16951-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="16951-191">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="16951-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="16951-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="16951-192">Type</span></span>

*   <span data-ttu-id="16951-193">String</span><span class="sxs-lookup"><span data-stu-id="16951-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="16951-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="16951-194">Requirements</span></span>

|<span data-ttu-id="16951-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="16951-195">Requirement</span></span>| <span data-ttu-id="16951-196">Valor</span><span class="sxs-lookup"><span data-stu-id="16951-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="16951-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="16951-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16951-198">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-198">1.1</span></span>|
|[<span data-ttu-id="16951-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="16951-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16951-200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="16951-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="16951-201">Exemplo</span><span class="sxs-lookup"><span data-stu-id="16951-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="16951-202">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="16951-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="16951-203">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="16951-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="16951-204">Tipo</span><span class="sxs-lookup"><span data-stu-id="16951-204">Type</span></span>

*   [<span data-ttu-id="16951-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="16951-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="16951-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="16951-206">Requirements</span></span>

|<span data-ttu-id="16951-207">Requisito</span><span class="sxs-lookup"><span data-stu-id="16951-207">Requirement</span></span>| <span data-ttu-id="16951-208">Valor</span><span class="sxs-lookup"><span data-stu-id="16951-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="16951-209">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="16951-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16951-210">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-210">1.1</span></span>|
|[<span data-ttu-id="16951-211">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="16951-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16951-212">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="16951-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="16951-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="16951-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="16951-214">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="16951-214">displayLanguage: String</span></span>

<span data-ttu-id="16951-215">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="16951-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="16951-216">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="16951-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="16951-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="16951-217">Type</span></span>

*   <span data-ttu-id="16951-218">String</span><span class="sxs-lookup"><span data-stu-id="16951-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="16951-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="16951-219">Requirements</span></span>

|<span data-ttu-id="16951-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="16951-220">Requirement</span></span>| <span data-ttu-id="16951-221">Valor</span><span class="sxs-lookup"><span data-stu-id="16951-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="16951-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="16951-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16951-223">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-223">1.1</span></span>|
|[<span data-ttu-id="16951-224">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="16951-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16951-225">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="16951-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="16951-226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="16951-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="16951-227">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="16951-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="16951-228">Obtém o host do aplicativo do Office no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="16951-228">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="16951-229">Tipo</span><span class="sxs-lookup"><span data-stu-id="16951-229">Type</span></span>

*   [<span data-ttu-id="16951-230">HostType</span><span class="sxs-lookup"><span data-stu-id="16951-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="16951-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="16951-231">Requirements</span></span>

|<span data-ttu-id="16951-232">Requisito</span><span class="sxs-lookup"><span data-stu-id="16951-232">Requirement</span></span>| <span data-ttu-id="16951-233">Valor</span><span class="sxs-lookup"><span data-stu-id="16951-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="16951-234">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="16951-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16951-235">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-235">1.1</span></span>|
|[<span data-ttu-id="16951-236">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="16951-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16951-237">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="16951-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="16951-238">Exemplo</span><span class="sxs-lookup"><span data-stu-id="16951-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="16951-239">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="16951-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="16951-240">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="16951-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="16951-241">Só há suporte para esse membro no Outlook no Windows.</span><span class="sxs-lookup"><span data-stu-id="16951-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="16951-242">O uso de cores de temas do Office permite coordenar o esquema de cores do seu suplemento com o tema atual do Office selecionado pelo usuário com a **conta de arquivo > office > Office Theme UI**, que é aplicada em todos os aplicativos host do Office.</span><span class="sxs-lookup"><span data-stu-id="16951-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="16951-243">Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="16951-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="16951-244">Tipo</span><span class="sxs-lookup"><span data-stu-id="16951-244">Type</span></span>

*   [<span data-ttu-id="16951-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="16951-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="16951-246">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="16951-246">Properties:</span></span>

|<span data-ttu-id="16951-247">Nome</span><span class="sxs-lookup"><span data-stu-id="16951-247">Name</span></span>| <span data-ttu-id="16951-248">Tipo</span><span class="sxs-lookup"><span data-stu-id="16951-248">Type</span></span>| <span data-ttu-id="16951-249">Descrição</span><span class="sxs-lookup"><span data-stu-id="16951-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="16951-250">String</span><span class="sxs-lookup"><span data-stu-id="16951-250">String</span></span>|<span data-ttu-id="16951-251">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="16951-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="16951-252">String</span><span class="sxs-lookup"><span data-stu-id="16951-252">String</span></span>|<span data-ttu-id="16951-253">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="16951-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="16951-254">String</span><span class="sxs-lookup"><span data-stu-id="16951-254">String</span></span>|<span data-ttu-id="16951-255">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="16951-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="16951-256">String</span><span class="sxs-lookup"><span data-stu-id="16951-256">String</span></span>|<span data-ttu-id="16951-257">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="16951-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="16951-258">Requisitos</span><span class="sxs-lookup"><span data-stu-id="16951-258">Requirements</span></span>

|<span data-ttu-id="16951-259">Requisito</span><span class="sxs-lookup"><span data-stu-id="16951-259">Requirement</span></span>| <span data-ttu-id="16951-260">Valor</span><span class="sxs-lookup"><span data-stu-id="16951-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="16951-261">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="16951-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16951-262">Visualização</span><span class="sxs-lookup"><span data-stu-id="16951-262">Preview</span></span>|
|[<span data-ttu-id="16951-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="16951-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16951-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="16951-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="16951-265">Exemplo</span><span class="sxs-lookup"><span data-stu-id="16951-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="16951-266">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="16951-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="16951-267">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="16951-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="16951-268">Tipo</span><span class="sxs-lookup"><span data-stu-id="16951-268">Type</span></span>

*   [<span data-ttu-id="16951-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="16951-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="16951-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="16951-270">Requirements</span></span>

|<span data-ttu-id="16951-271">Requisito</span><span class="sxs-lookup"><span data-stu-id="16951-271">Requirement</span></span>| <span data-ttu-id="16951-272">Valor</span><span class="sxs-lookup"><span data-stu-id="16951-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="16951-273">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="16951-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16951-274">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-274">1.1</span></span>|
|[<span data-ttu-id="16951-275">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="16951-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16951-276">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="16951-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="16951-277">Exemplo</span><span class="sxs-lookup"><span data-stu-id="16951-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="16951-278">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="16951-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="16951-279">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o host atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="16951-279">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="16951-280">Tipo</span><span class="sxs-lookup"><span data-stu-id="16951-280">Type</span></span>

*   [<span data-ttu-id="16951-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="16951-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="16951-282">Requisitos</span><span class="sxs-lookup"><span data-stu-id="16951-282">Requirements</span></span>

|<span data-ttu-id="16951-283">Requisito</span><span class="sxs-lookup"><span data-stu-id="16951-283">Requirement</span></span>| <span data-ttu-id="16951-284">Valor</span><span class="sxs-lookup"><span data-stu-id="16951-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="16951-285">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="16951-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16951-286">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-286">1.1</span></span>|
|[<span data-ttu-id="16951-287">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="16951-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16951-288">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="16951-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="16951-289">Exemplo</span><span class="sxs-lookup"><span data-stu-id="16951-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="16951-290">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="16951-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="16951-291">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="16951-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="16951-292">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="16951-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="16951-293">Tipo</span><span class="sxs-lookup"><span data-stu-id="16951-293">Type</span></span>

*   [<span data-ttu-id="16951-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="16951-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="16951-295">Requisitos</span><span class="sxs-lookup"><span data-stu-id="16951-295">Requirements</span></span>

|<span data-ttu-id="16951-296">Requisito</span><span class="sxs-lookup"><span data-stu-id="16951-296">Requirement</span></span>| <span data-ttu-id="16951-297">Valor</span><span class="sxs-lookup"><span data-stu-id="16951-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="16951-298">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="16951-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16951-299">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-299">1.1</span></span>|
|[<span data-ttu-id="16951-300">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="16951-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="16951-301">Restrito</span><span class="sxs-lookup"><span data-stu-id="16951-301">Restricted</span></span>|
|[<span data-ttu-id="16951-302">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="16951-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16951-303">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="16951-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="16951-304">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="16951-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="16951-305">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="16951-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="16951-306">Tipo</span><span class="sxs-lookup"><span data-stu-id="16951-306">Type</span></span>

*   [<span data-ttu-id="16951-307">UI</span><span class="sxs-lookup"><span data-stu-id="16951-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="16951-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="16951-308">Requirements</span></span>

|<span data-ttu-id="16951-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="16951-309">Requirement</span></span>| <span data-ttu-id="16951-310">Valor</span><span class="sxs-lookup"><span data-stu-id="16951-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="16951-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="16951-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="16951-312">1.1</span><span class="sxs-lookup"><span data-stu-id="16951-312">1.1</span></span>|
|[<span data-ttu-id="16951-313">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="16951-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="16951-314">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="16951-314">Compose or Read</span></span>|
