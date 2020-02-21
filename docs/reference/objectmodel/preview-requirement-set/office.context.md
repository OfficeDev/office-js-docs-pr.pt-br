---
title: Office. Context – conjunto de requisitos de visualização
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9c2c661ce870e2007bd891aee040c6b3564f7b9e
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165514"
---
# <a name="context"></a><span data-ttu-id="cda8c-102">context</span><span class="sxs-lookup"><span data-stu-id="cda8c-102">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="cda8c-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="cda8c-103">[Office](office.md).context</span></span>

<span data-ttu-id="cda8c-104">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="cda8c-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="cda8c-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-preview).</span><span class="sxs-lookup"><span data-stu-id="cda8c-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="cda8c-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cda8c-106">Requirements</span></span>

|<span data-ttu-id="cda8c-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="cda8c-107">Requirement</span></span>| <span data-ttu-id="cda8c-108">Valor</span><span class="sxs-lookup"><span data-stu-id="cda8c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="cda8c-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cda8c-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cda8c-110">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-110">1.1</span></span>|
|[<span data-ttu-id="cda8c-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cda8c-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cda8c-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="cda8c-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="cda8c-113">Properties</span></span>

| <span data-ttu-id="cda8c-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="cda8c-114">Property</span></span> | <span data-ttu-id="cda8c-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="cda8c-115">Modes</span></span> | <span data-ttu-id="cda8c-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="cda8c-116">Return type</span></span> | <span data-ttu-id="cda8c-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="cda8c-117">Minimum</span></span><br><span data-ttu-id="cda8c-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="cda8c-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cda8c-119">autentica</span><span class="sxs-lookup"><span data-stu-id="cda8c-119">auth</span></span>](#auth-auth) | <span data-ttu-id="cda8c-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="cda8c-120">Compose</span></span><br><span data-ttu-id="cda8c-121">Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-121">Read</span></span> | [<span data-ttu-id="cda8c-122">Auth</span><span class="sxs-lookup"><span data-stu-id="cda8c-122">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="cda8c-123">Visualização</span><span class="sxs-lookup"><span data-stu-id="cda8c-123">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="cda8c-124">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="cda8c-124">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="cda8c-125">Escrever</span><span class="sxs-lookup"><span data-stu-id="cda8c-125">Compose</span></span><br><span data-ttu-id="cda8c-126">Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-126">Read</span></span> | <span data-ttu-id="cda8c-127">String</span><span class="sxs-lookup"><span data-stu-id="cda8c-127">String</span></span> | [<span data-ttu-id="cda8c-128">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cda8c-129">la</span><span class="sxs-lookup"><span data-stu-id="cda8c-129">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="cda8c-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="cda8c-130">Compose</span></span><br><span data-ttu-id="cda8c-131">Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-131">Read</span></span> | [<span data-ttu-id="cda8c-132">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="cda8c-132">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="cda8c-133">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cda8c-134">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="cda8c-134">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="cda8c-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="cda8c-135">Compose</span></span><br><span data-ttu-id="cda8c-136">Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-136">Read</span></span> | <span data-ttu-id="cda8c-137">String</span><span class="sxs-lookup"><span data-stu-id="cda8c-137">String</span></span> | [<span data-ttu-id="cda8c-138">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cda8c-139">principal</span><span class="sxs-lookup"><span data-stu-id="cda8c-139">host</span></span>](#host-hosttype) | <span data-ttu-id="cda8c-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="cda8c-140">Compose</span></span><br><span data-ttu-id="cda8c-141">Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-141">Read</span></span> | [<span data-ttu-id="cda8c-142">HostType</span><span class="sxs-lookup"><span data-stu-id="cda8c-142">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="cda8c-143">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cda8c-144">mailbox</span><span class="sxs-lookup"><span data-stu-id="cda8c-144">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="cda8c-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="cda8c-145">Compose</span></span><br><span data-ttu-id="cda8c-146">Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-146">Read</span></span> | [<span data-ttu-id="cda8c-147">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="cda8c-147">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="cda8c-148">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cda8c-149">officeTheme</span><span class="sxs-lookup"><span data-stu-id="cda8c-149">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="cda8c-150">Escrever</span><span class="sxs-lookup"><span data-stu-id="cda8c-150">Compose</span></span><br><span data-ttu-id="cda8c-151">Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-151">Read</span></span> | [<span data-ttu-id="cda8c-152">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="cda8c-152">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="cda8c-153">Visualização</span><span class="sxs-lookup"><span data-stu-id="cda8c-153">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="cda8c-154">plataforma</span><span class="sxs-lookup"><span data-stu-id="cda8c-154">platform</span></span>](#platform-platformtype) | <span data-ttu-id="cda8c-155">Escrever</span><span class="sxs-lookup"><span data-stu-id="cda8c-155">Compose</span></span><br><span data-ttu-id="cda8c-156">Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-156">Read</span></span> | [<span data-ttu-id="cda8c-157">PlatformType</span><span class="sxs-lookup"><span data-stu-id="cda8c-157">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="cda8c-158">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cda8c-159">atende</span><span class="sxs-lookup"><span data-stu-id="cda8c-159">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="cda8c-160">Escrever</span><span class="sxs-lookup"><span data-stu-id="cda8c-160">Compose</span></span><br><span data-ttu-id="cda8c-161">Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-161">Read</span></span> | [<span data-ttu-id="cda8c-162">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="cda8c-162">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="cda8c-163">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cda8c-164">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="cda8c-164">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="cda8c-165">Escrever</span><span class="sxs-lookup"><span data-stu-id="cda8c-165">Compose</span></span><br><span data-ttu-id="cda8c-166">Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-166">Read</span></span> | [<span data-ttu-id="cda8c-167">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="cda8c-167">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="cda8c-168">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-168">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cda8c-169">ui</span><span class="sxs-lookup"><span data-stu-id="cda8c-169">ui</span></span>](#ui-ui) | <span data-ttu-id="cda8c-170">Escrever</span><span class="sxs-lookup"><span data-stu-id="cda8c-170">Compose</span></span><br><span data-ttu-id="cda8c-171">Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-171">Read</span></span> | [<span data-ttu-id="cda8c-172">UI</span><span class="sxs-lookup"><span data-stu-id="cda8c-172">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="cda8c-173">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-173">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="cda8c-174">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="cda8c-174">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="cda8c-175">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="cda8c-175">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="cda8c-176">Oferece suporte a [logon único (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , fornecendo um método que permite que o host do Office obtenha um token de acesso para o aplicativo Web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="cda8c-176">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="cda8c-177">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="cda8c-177">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="cda8c-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="cda8c-178">Type</span></span>

*   [<span data-ttu-id="cda8c-179">Auth</span><span class="sxs-lookup"><span data-stu-id="cda8c-179">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="cda8c-180">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cda8c-180">Requirements</span></span>

|<span data-ttu-id="cda8c-181">Requisito</span><span class="sxs-lookup"><span data-stu-id="cda8c-181">Requirement</span></span>| <span data-ttu-id="cda8c-182">Valor</span><span class="sxs-lookup"><span data-stu-id="cda8c-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="cda8c-183">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cda8c-183">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cda8c-184">Visualização</span><span class="sxs-lookup"><span data-stu-id="cda8c-184">Preview</span></span>|
|[<span data-ttu-id="cda8c-185">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cda8c-185">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cda8c-186">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-186">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cda8c-187">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cda8c-187">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="cda8c-188">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="cda8c-188">contentLanguage: String</span></span>

<span data-ttu-id="cda8c-189">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="cda8c-189">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="cda8c-190">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="cda8c-190">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="cda8c-191">Tipo</span><span class="sxs-lookup"><span data-stu-id="cda8c-191">Type</span></span>

*   <span data-ttu-id="cda8c-192">String</span><span class="sxs-lookup"><span data-stu-id="cda8c-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cda8c-193">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cda8c-193">Requirements</span></span>

|<span data-ttu-id="cda8c-194">Requisito</span><span class="sxs-lookup"><span data-stu-id="cda8c-194">Requirement</span></span>| <span data-ttu-id="cda8c-195">Valor</span><span class="sxs-lookup"><span data-stu-id="cda8c-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="cda8c-196">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cda8c-196">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cda8c-197">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-197">1.1</span></span>|
|[<span data-ttu-id="cda8c-198">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cda8c-198">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cda8c-199">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-199">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cda8c-200">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cda8c-200">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="cda8c-201">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="cda8c-201">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="cda8c-202">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="cda8c-202">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="cda8c-203">Tipo</span><span class="sxs-lookup"><span data-stu-id="cda8c-203">Type</span></span>

*   [<span data-ttu-id="cda8c-204">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="cda8c-204">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="cda8c-205">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cda8c-205">Requirements</span></span>

|<span data-ttu-id="cda8c-206">Requisito</span><span class="sxs-lookup"><span data-stu-id="cda8c-206">Requirement</span></span>| <span data-ttu-id="cda8c-207">Valor</span><span class="sxs-lookup"><span data-stu-id="cda8c-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="cda8c-208">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cda8c-208">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cda8c-209">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-209">1.1</span></span>|
|[<span data-ttu-id="cda8c-210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cda8c-210">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cda8c-211">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-211">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cda8c-212">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cda8c-212">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="cda8c-213">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="cda8c-213">displayLanguage: String</span></span>

<span data-ttu-id="cda8c-214">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="cda8c-214">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="cda8c-215">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="cda8c-215">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="cda8c-216">Tipo</span><span class="sxs-lookup"><span data-stu-id="cda8c-216">Type</span></span>

*   <span data-ttu-id="cda8c-217">String</span><span class="sxs-lookup"><span data-stu-id="cda8c-217">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cda8c-218">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cda8c-218">Requirements</span></span>

|<span data-ttu-id="cda8c-219">Requisito</span><span class="sxs-lookup"><span data-stu-id="cda8c-219">Requirement</span></span>| <span data-ttu-id="cda8c-220">Valor</span><span class="sxs-lookup"><span data-stu-id="cda8c-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="cda8c-221">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cda8c-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cda8c-222">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-222">1.1</span></span>|
|[<span data-ttu-id="cda8c-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cda8c-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cda8c-224">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cda8c-225">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cda8c-225">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="cda8c-226">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="cda8c-226">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="cda8c-227">Obtém o host do aplicativo do Office no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="cda8c-227">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="cda8c-228">Tipo</span><span class="sxs-lookup"><span data-stu-id="cda8c-228">Type</span></span>

*   [<span data-ttu-id="cda8c-229">HostType</span><span class="sxs-lookup"><span data-stu-id="cda8c-229">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="cda8c-230">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cda8c-230">Requirements</span></span>

|<span data-ttu-id="cda8c-231">Requisito</span><span class="sxs-lookup"><span data-stu-id="cda8c-231">Requirement</span></span>| <span data-ttu-id="cda8c-232">Valor</span><span class="sxs-lookup"><span data-stu-id="cda8c-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="cda8c-233">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cda8c-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cda8c-234">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-234">1.1</span></span>|
|[<span data-ttu-id="cda8c-235">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cda8c-235">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cda8c-236">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-236">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cda8c-237">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cda8c-237">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="cda8c-238">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="cda8c-238">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="cda8c-239">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="cda8c-239">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="cda8c-240">Só há suporte para esse membro no Outlook no Windows.</span><span class="sxs-lookup"><span data-stu-id="cda8c-240">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="cda8c-241">O uso de cores de temas do Office permite coordenar o esquema de cores do seu suplemento com o tema atual do Office selecionado pelo usuário com a **conta de arquivo > office > Office Theme UI**, que é aplicada em todos os aplicativos host do Office.</span><span class="sxs-lookup"><span data-stu-id="cda8c-241">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="cda8c-242">Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="cda8c-242">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="cda8c-243">Tipo</span><span class="sxs-lookup"><span data-stu-id="cda8c-243">Type</span></span>

*   [<span data-ttu-id="cda8c-244">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="cda8c-244">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="cda8c-245">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="cda8c-245">Properties:</span></span>

|<span data-ttu-id="cda8c-246">Nome</span><span class="sxs-lookup"><span data-stu-id="cda8c-246">Name</span></span>| <span data-ttu-id="cda8c-247">Tipo</span><span class="sxs-lookup"><span data-stu-id="cda8c-247">Type</span></span>| <span data-ttu-id="cda8c-248">Descrição</span><span class="sxs-lookup"><span data-stu-id="cda8c-248">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="cda8c-249">String</span><span class="sxs-lookup"><span data-stu-id="cda8c-249">String</span></span>|<span data-ttu-id="cda8c-250">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="cda8c-250">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="cda8c-251">String</span><span class="sxs-lookup"><span data-stu-id="cda8c-251">String</span></span>|<span data-ttu-id="cda8c-252">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="cda8c-252">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="cda8c-253">String</span><span class="sxs-lookup"><span data-stu-id="cda8c-253">String</span></span>|<span data-ttu-id="cda8c-254">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="cda8c-254">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="cda8c-255">String</span><span class="sxs-lookup"><span data-stu-id="cda8c-255">String</span></span>|<span data-ttu-id="cda8c-256">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="cda8c-256">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cda8c-257">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cda8c-257">Requirements</span></span>

|<span data-ttu-id="cda8c-258">Requisito</span><span class="sxs-lookup"><span data-stu-id="cda8c-258">Requirement</span></span>| <span data-ttu-id="cda8c-259">Valor</span><span class="sxs-lookup"><span data-stu-id="cda8c-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="cda8c-260">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cda8c-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cda8c-261">Visualização</span><span class="sxs-lookup"><span data-stu-id="cda8c-261">Preview</span></span>|
|[<span data-ttu-id="cda8c-262">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cda8c-262">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cda8c-263">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-263">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cda8c-264">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cda8c-264">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="cda8c-265">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="cda8c-265">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="cda8c-266">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="cda8c-266">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="cda8c-267">Tipo</span><span class="sxs-lookup"><span data-stu-id="cda8c-267">Type</span></span>

*   [<span data-ttu-id="cda8c-268">PlatformType</span><span class="sxs-lookup"><span data-stu-id="cda8c-268">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="cda8c-269">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cda8c-269">Requirements</span></span>

|<span data-ttu-id="cda8c-270">Requisito</span><span class="sxs-lookup"><span data-stu-id="cda8c-270">Requirement</span></span>| <span data-ttu-id="cda8c-271">Valor</span><span class="sxs-lookup"><span data-stu-id="cda8c-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="cda8c-272">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cda8c-272">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cda8c-273">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-273">1.1</span></span>|
|[<span data-ttu-id="cda8c-274">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cda8c-274">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cda8c-275">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-275">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cda8c-276">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cda8c-276">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="cda8c-277">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="cda8c-277">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="cda8c-278">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o host atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="cda8c-278">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="cda8c-279">Tipo</span><span class="sxs-lookup"><span data-stu-id="cda8c-279">Type</span></span>

*   [<span data-ttu-id="cda8c-280">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="cda8c-280">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="cda8c-281">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cda8c-281">Requirements</span></span>

|<span data-ttu-id="cda8c-282">Requisito</span><span class="sxs-lookup"><span data-stu-id="cda8c-282">Requirement</span></span>| <span data-ttu-id="cda8c-283">Valor</span><span class="sxs-lookup"><span data-stu-id="cda8c-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="cda8c-284">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cda8c-284">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cda8c-285">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-285">1.1</span></span>|
|[<span data-ttu-id="cda8c-286">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cda8c-286">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cda8c-287">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-287">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cda8c-288">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cda8c-288">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="cda8c-289">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="cda8c-289">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="cda8c-290">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="cda8c-290">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="cda8c-291">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="cda8c-291">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="cda8c-292">Tipo</span><span class="sxs-lookup"><span data-stu-id="cda8c-292">Type</span></span>

*   [<span data-ttu-id="cda8c-293">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="cda8c-293">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="cda8c-294">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cda8c-294">Requirements</span></span>

|<span data-ttu-id="cda8c-295">Requisito</span><span class="sxs-lookup"><span data-stu-id="cda8c-295">Requirement</span></span>| <span data-ttu-id="cda8c-296">Valor</span><span class="sxs-lookup"><span data-stu-id="cda8c-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="cda8c-297">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cda8c-297">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cda8c-298">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-298">1.1</span></span>|
|[<span data-ttu-id="cda8c-299">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="cda8c-299">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="cda8c-300">Restrito</span><span class="sxs-lookup"><span data-stu-id="cda8c-300">Restricted</span></span>|
|[<span data-ttu-id="cda8c-301">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cda8c-301">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cda8c-302">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-302">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="cda8c-303">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="cda8c-303">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="cda8c-304">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="cda8c-304">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="cda8c-305">Tipo</span><span class="sxs-lookup"><span data-stu-id="cda8c-305">Type</span></span>

*   [<span data-ttu-id="cda8c-306">UI</span><span class="sxs-lookup"><span data-stu-id="cda8c-306">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="cda8c-307">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cda8c-307">Requirements</span></span>

|<span data-ttu-id="cda8c-308">Requisito</span><span class="sxs-lookup"><span data-stu-id="cda8c-308">Requirement</span></span>| <span data-ttu-id="cda8c-309">Valor</span><span class="sxs-lookup"><span data-stu-id="cda8c-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="cda8c-310">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cda8c-310">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cda8c-311">1.1</span><span class="sxs-lookup"><span data-stu-id="cda8c-311">1.1</span></span>|
|[<span data-ttu-id="cda8c-312">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cda8c-312">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cda8c-313">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cda8c-313">Compose or Read</span></span>|
