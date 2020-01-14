---
title: Office. Context – conjunto de requisitos de visualização
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 08f26de89624e6e06bc57382afe8e02b018029ca
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111148"
---
# <a name="context"></a><span data-ttu-id="09bb2-102">context</span><span class="sxs-lookup"><span data-stu-id="09bb2-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="09bb2-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="09bb2-103">[Office](office.md).context</span></span>

<span data-ttu-id="09bb2-104">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="09bb2-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="09bb2-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-preview).</span><span class="sxs-lookup"><span data-stu-id="09bb2-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="09bb2-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb2-106">Requirements</span></span>

|<span data-ttu-id="09bb2-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb2-107">Requirement</span></span>| <span data-ttu-id="09bb2-108">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb2-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb2-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb2-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="09bb2-110">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-110">1.1</span></span>|
|[<span data-ttu-id="09bb2-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb2-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb2-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb2-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="09bb2-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="09bb2-113">Properties</span></span>

| <span data-ttu-id="09bb2-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="09bb2-114">Property</span></span> | <span data-ttu-id="09bb2-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="09bb2-115">Modes</span></span> | <span data-ttu-id="09bb2-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="09bb2-116">Return type</span></span> | <span data-ttu-id="09bb2-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="09bb2-117">Minimum</span></span><br><span data-ttu-id="09bb2-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb2-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="09bb2-119">autentica</span><span class="sxs-lookup"><span data-stu-id="09bb2-119">auth</span></span>](#auth-auth) | <span data-ttu-id="09bb2-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="09bb2-120">Compose</span></span><br><span data-ttu-id="09bb2-121">Leitura</span><span class="sxs-lookup"><span data-stu-id="09bb2-121">Read</span></span> | [<span data-ttu-id="09bb2-122">Auth</span><span class="sxs-lookup"><span data-stu-id="09bb2-122">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="09bb2-123">Visualização</span><span class="sxs-lookup"><span data-stu-id="09bb2-123">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="09bb2-124">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="09bb2-124">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="09bb2-125">Escrever</span><span class="sxs-lookup"><span data-stu-id="09bb2-125">Compose</span></span><br><span data-ttu-id="09bb2-126">Leitura</span><span class="sxs-lookup"><span data-stu-id="09bb2-126">Read</span></span> | <span data-ttu-id="09bb2-127">String</span><span class="sxs-lookup"><span data-stu-id="09bb2-127">String</span></span> | [<span data-ttu-id="09bb2-128">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="09bb2-129">la</span><span class="sxs-lookup"><span data-stu-id="09bb2-129">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="09bb2-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="09bb2-130">Compose</span></span><br><span data-ttu-id="09bb2-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="09bb2-131">Read</span></span> | [<span data-ttu-id="09bb2-132">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="09bb2-132">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="09bb2-133">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="09bb2-134">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="09bb2-134">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="09bb2-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="09bb2-135">Compose</span></span><br><span data-ttu-id="09bb2-136">Leitura</span><span class="sxs-lookup"><span data-stu-id="09bb2-136">Read</span></span> | <span data-ttu-id="09bb2-137">String</span><span class="sxs-lookup"><span data-stu-id="09bb2-137">String</span></span> | [<span data-ttu-id="09bb2-138">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="09bb2-139">principal</span><span class="sxs-lookup"><span data-stu-id="09bb2-139">host</span></span>](#host-hosttype) | <span data-ttu-id="09bb2-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="09bb2-140">Compose</span></span><br><span data-ttu-id="09bb2-141">Leitura</span><span class="sxs-lookup"><span data-stu-id="09bb2-141">Read</span></span> | [<span data-ttu-id="09bb2-142">HostType</span><span class="sxs-lookup"><span data-stu-id="09bb2-142">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="09bb2-143">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="09bb2-144">mailbox</span><span class="sxs-lookup"><span data-stu-id="09bb2-144">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="09bb2-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="09bb2-145">Compose</span></span><br><span data-ttu-id="09bb2-146">Leitura</span><span class="sxs-lookup"><span data-stu-id="09bb2-146">Read</span></span> | [<span data-ttu-id="09bb2-147">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="09bb2-147">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="09bb2-148">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="09bb2-149">officeTheme</span><span class="sxs-lookup"><span data-stu-id="09bb2-149">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="09bb2-150">Escrever</span><span class="sxs-lookup"><span data-stu-id="09bb2-150">Compose</span></span><br><span data-ttu-id="09bb2-151">Leitura</span><span class="sxs-lookup"><span data-stu-id="09bb2-151">Read</span></span> | [<span data-ttu-id="09bb2-152">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="09bb2-152">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="09bb2-153">Visualização</span><span class="sxs-lookup"><span data-stu-id="09bb2-153">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="09bb2-154">plataforma</span><span class="sxs-lookup"><span data-stu-id="09bb2-154">platform</span></span>](#platform-platformtype) | <span data-ttu-id="09bb2-155">Escrever</span><span class="sxs-lookup"><span data-stu-id="09bb2-155">Compose</span></span><br><span data-ttu-id="09bb2-156">Leitura</span><span class="sxs-lookup"><span data-stu-id="09bb2-156">Read</span></span> | [<span data-ttu-id="09bb2-157">PlatformType</span><span class="sxs-lookup"><span data-stu-id="09bb2-157">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="09bb2-158">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="09bb2-159">atende</span><span class="sxs-lookup"><span data-stu-id="09bb2-159">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="09bb2-160">Escrever</span><span class="sxs-lookup"><span data-stu-id="09bb2-160">Compose</span></span><br><span data-ttu-id="09bb2-161">Leitura</span><span class="sxs-lookup"><span data-stu-id="09bb2-161">Read</span></span> | [<span data-ttu-id="09bb2-162">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="09bb2-162">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="09bb2-163">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="09bb2-164">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="09bb2-164">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="09bb2-165">Escrever</span><span class="sxs-lookup"><span data-stu-id="09bb2-165">Compose</span></span><br><span data-ttu-id="09bb2-166">Leitura</span><span class="sxs-lookup"><span data-stu-id="09bb2-166">Read</span></span> | [<span data-ttu-id="09bb2-167">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="09bb2-167">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="09bb2-168">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-168">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="09bb2-169">ui</span><span class="sxs-lookup"><span data-stu-id="09bb2-169">ui</span></span>](#ui-ui) | <span data-ttu-id="09bb2-170">Escrever</span><span class="sxs-lookup"><span data-stu-id="09bb2-170">Compose</span></span><br><span data-ttu-id="09bb2-171">Leitura</span><span class="sxs-lookup"><span data-stu-id="09bb2-171">Read</span></span> | [<span data-ttu-id="09bb2-172">UI</span><span class="sxs-lookup"><span data-stu-id="09bb2-172">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="09bb2-173">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-173">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="09bb2-174">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="09bb2-174">Property details</span></span>

#### <a name="auth-authjavascriptapiofficeofficeauth"></a><span data-ttu-id="09bb2-175">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="09bb2-175">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="09bb2-176">Oferece suporte a [logon único (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token) , fornecendo um método que permite que o host do Office obtenha um token de acesso para o aplicativo Web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="09bb2-176">Supports [single sign-on (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="09bb2-177">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="09bb2-177">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb2-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb2-178">Type</span></span>

*   [<span data-ttu-id="09bb2-179">Auth</span><span class="sxs-lookup"><span data-stu-id="09bb2-179">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="09bb2-180">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb2-180">Requirements</span></span>

|<span data-ttu-id="09bb2-181">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb2-181">Requirement</span></span>| <span data-ttu-id="09bb2-182">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb2-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb2-183">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb2-183">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="09bb2-184">Visualização</span><span class="sxs-lookup"><span data-stu-id="09bb2-184">Preview</span></span>|
|[<span data-ttu-id="09bb2-185">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb2-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb2-186">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb2-186">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09bb2-187">Exemplo</span><span class="sxs-lookup"><span data-stu-id="09bb2-187">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="09bb2-188">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="09bb2-188">contentLanguage: String</span></span>

<span data-ttu-id="09bb2-189">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="09bb2-189">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="09bb2-190">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="09bb2-190">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb2-191">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb2-191">Type</span></span>

*   <span data-ttu-id="09bb2-192">String</span><span class="sxs-lookup"><span data-stu-id="09bb2-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09bb2-193">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb2-193">Requirements</span></span>

|<span data-ttu-id="09bb2-194">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb2-194">Requirement</span></span>| <span data-ttu-id="09bb2-195">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb2-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb2-196">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb2-196">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="09bb2-197">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-197">1.1</span></span>|
|[<span data-ttu-id="09bb2-198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb2-198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb2-199">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb2-199">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09bb2-200">Exemplo</span><span class="sxs-lookup"><span data-stu-id="09bb2-200">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="09bb2-201">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="09bb2-201">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="09bb2-202">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="09bb2-202">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb2-203">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb2-203">Type</span></span>

*   [<span data-ttu-id="09bb2-204">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="09bb2-204">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="09bb2-205">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb2-205">Requirements</span></span>

|<span data-ttu-id="09bb2-206">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb2-206">Requirement</span></span>| <span data-ttu-id="09bb2-207">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb2-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb2-208">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb2-208">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="09bb2-209">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-209">1.1</span></span>|
|[<span data-ttu-id="09bb2-210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb2-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb2-211">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb2-211">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09bb2-212">Exemplo</span><span class="sxs-lookup"><span data-stu-id="09bb2-212">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="09bb2-213">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="09bb2-213">displayLanguage: String</span></span>

<span data-ttu-id="09bb2-214">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="09bb2-214">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="09bb2-215">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="09bb2-215">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb2-216">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb2-216">Type</span></span>

*   <span data-ttu-id="09bb2-217">String</span><span class="sxs-lookup"><span data-stu-id="09bb2-217">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09bb2-218">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb2-218">Requirements</span></span>

|<span data-ttu-id="09bb2-219">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb2-219">Requirement</span></span>| <span data-ttu-id="09bb2-220">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb2-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb2-221">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb2-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="09bb2-222">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-222">1.1</span></span>|
|[<span data-ttu-id="09bb2-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb2-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb2-224">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb2-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09bb2-225">Exemplo</span><span class="sxs-lookup"><span data-stu-id="09bb2-225">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="09bb2-226">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="09bb2-226">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="09bb2-227">Obtém o host do aplicativo do Office no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="09bb2-227">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb2-228">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb2-228">Type</span></span>

*   [<span data-ttu-id="09bb2-229">HostType</span><span class="sxs-lookup"><span data-stu-id="09bb2-229">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="09bb2-230">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb2-230">Requirements</span></span>

|<span data-ttu-id="09bb2-231">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb2-231">Requirement</span></span>| <span data-ttu-id="09bb2-232">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb2-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb2-233">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb2-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="09bb2-234">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-234">1.1</span></span>|
|[<span data-ttu-id="09bb2-235">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb2-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb2-236">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb2-236">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09bb2-237">Exemplo</span><span class="sxs-lookup"><span data-stu-id="09bb2-237">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officethemejavascriptapiofficeofficeofficetheme"></a><span data-ttu-id="09bb2-238">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="09bb2-238">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="09bb2-239">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="09bb2-239">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="09bb2-240">Só há suporte para esse membro no Outlook no Windows.</span><span class="sxs-lookup"><span data-stu-id="09bb2-240">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="09bb2-241">O uso de cores de temas do Office permite coordenar o esquema de cores do seu suplemento com o tema atual do Office selecionado pelo usuário com a **conta de arquivo > office > Office Theme UI**, que é aplicada em todos os aplicativos host do Office.</span><span class="sxs-lookup"><span data-stu-id="09bb2-241">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="09bb2-242">Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="09bb2-242">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb2-243">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb2-243">Type</span></span>

*   [<span data-ttu-id="09bb2-244">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="09bb2-244">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="09bb2-245">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="09bb2-245">Properties:</span></span>

|<span data-ttu-id="09bb2-246">Nome</span><span class="sxs-lookup"><span data-stu-id="09bb2-246">Name</span></span>| <span data-ttu-id="09bb2-247">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb2-247">Type</span></span>| <span data-ttu-id="09bb2-248">Descrição</span><span class="sxs-lookup"><span data-stu-id="09bb2-248">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="09bb2-249">String</span><span class="sxs-lookup"><span data-stu-id="09bb2-249">String</span></span>|<span data-ttu-id="09bb2-250">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="09bb2-250">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="09bb2-251">String</span><span class="sxs-lookup"><span data-stu-id="09bb2-251">String</span></span>|<span data-ttu-id="09bb2-252">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="09bb2-252">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="09bb2-253">String</span><span class="sxs-lookup"><span data-stu-id="09bb2-253">String</span></span>|<span data-ttu-id="09bb2-254">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="09bb2-254">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="09bb2-255">String</span><span class="sxs-lookup"><span data-stu-id="09bb2-255">String</span></span>|<span data-ttu-id="09bb2-256">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="09bb2-256">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09bb2-257">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb2-257">Requirements</span></span>

|<span data-ttu-id="09bb2-258">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb2-258">Requirement</span></span>| <span data-ttu-id="09bb2-259">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb2-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb2-260">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb2-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="09bb2-261">Visualização</span><span class="sxs-lookup"><span data-stu-id="09bb2-261">Preview</span></span>|
|[<span data-ttu-id="09bb2-262">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb2-262">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb2-263">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb2-263">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09bb2-264">Exemplo</span><span class="sxs-lookup"><span data-stu-id="09bb2-264">Example</span></span>

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

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="09bb2-265">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="09bb2-265">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="09bb2-266">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="09bb2-266">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb2-267">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb2-267">Type</span></span>

*   [<span data-ttu-id="09bb2-268">PlatformType</span><span class="sxs-lookup"><span data-stu-id="09bb2-268">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="09bb2-269">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb2-269">Requirements</span></span>

|<span data-ttu-id="09bb2-270">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb2-270">Requirement</span></span>| <span data-ttu-id="09bb2-271">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb2-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb2-272">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb2-272">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="09bb2-273">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-273">1.1</span></span>|
|[<span data-ttu-id="09bb2-274">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb2-274">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb2-275">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb2-275">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09bb2-276">Exemplo</span><span class="sxs-lookup"><span data-stu-id="09bb2-276">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="09bb2-277">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="09bb2-277">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="09bb2-278">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o host atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="09bb2-278">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb2-279">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb2-279">Type</span></span>

*   [<span data-ttu-id="09bb2-280">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="09bb2-280">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="09bb2-281">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb2-281">Requirements</span></span>

|<span data-ttu-id="09bb2-282">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb2-282">Requirement</span></span>| <span data-ttu-id="09bb2-283">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb2-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb2-284">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb2-284">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="09bb2-285">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-285">1.1</span></span>|
|[<span data-ttu-id="09bb2-286">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb2-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb2-287">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb2-287">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09bb2-288">Exemplo</span><span class="sxs-lookup"><span data-stu-id="09bb2-288">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="09bb2-289">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="09bb2-289">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="09bb2-290">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="09bb2-290">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="09bb2-291">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="09bb2-291">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb2-292">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb2-292">Type</span></span>

*   [<span data-ttu-id="09bb2-293">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="09bb2-293">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="09bb2-294">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb2-294">Requirements</span></span>

|<span data-ttu-id="09bb2-295">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb2-295">Requirement</span></span>| <span data-ttu-id="09bb2-296">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb2-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb2-297">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb2-297">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="09bb2-298">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-298">1.1</span></span>|
|[<span data-ttu-id="09bb2-299">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="09bb2-299">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09bb2-300">Restrito</span><span class="sxs-lookup"><span data-stu-id="09bb2-300">Restricted</span></span>|
|[<span data-ttu-id="09bb2-301">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb2-301">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb2-302">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb2-302">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="09bb2-303">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="09bb2-303">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="09bb2-304">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="09bb2-304">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb2-305">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb2-305">Type</span></span>

*   [<span data-ttu-id="09bb2-306">UI</span><span class="sxs-lookup"><span data-stu-id="09bb2-306">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="09bb2-307">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb2-307">Requirements</span></span>

|<span data-ttu-id="09bb2-308">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb2-308">Requirement</span></span>| <span data-ttu-id="09bb2-309">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb2-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb2-310">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb2-310">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="09bb2-311">1.1</span><span class="sxs-lookup"><span data-stu-id="09bb2-311">1.1</span></span>|
|[<span data-ttu-id="09bb2-312">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb2-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb2-313">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb2-313">Compose or Read</span></span>|
