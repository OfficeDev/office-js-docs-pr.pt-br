---
title: Office. Context – conjunto de requisitos de visualização
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de visualização da API da caixa de correio.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 0e0ea973032bb5cd854856920e192522f90a26a1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612021"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="ac55d-103">contexto (conjunto de requisitos de visualização da caixa de correio)</span><span class="sxs-lookup"><span data-stu-id="ac55d-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="ac55d-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="ac55d-104">[Office](office.md).context</span></span>

<span data-ttu-id="ac55d-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="ac55d-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="ac55d-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-preview).</span><span class="sxs-lookup"><span data-stu-id="ac55d-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac55d-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac55d-107">Requirements</span></span>

|<span data-ttu-id="ac55d-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac55d-108">Requirement</span></span>| <span data-ttu-id="ac55d-109">Valor</span><span class="sxs-lookup"><span data-stu-id="ac55d-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac55d-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac55d-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac55d-111">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-111">1.1</span></span>|
|[<span data-ttu-id="ac55d-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac55d-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac55d-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac55d-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ac55d-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="ac55d-114">Properties</span></span>

| <span data-ttu-id="ac55d-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="ac55d-115">Property</span></span> | <span data-ttu-id="ac55d-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="ac55d-116">Modes</span></span> | <span data-ttu-id="ac55d-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="ac55d-117">Return type</span></span> | <span data-ttu-id="ac55d-118">Mínimo</span><span class="sxs-lookup"><span data-stu-id="ac55d-118">Minimum</span></span><br><span data-ttu-id="ac55d-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="ac55d-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ac55d-120">autentica</span><span class="sxs-lookup"><span data-stu-id="ac55d-120">auth</span></span>](#auth-auth) | <span data-ttu-id="ac55d-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac55d-121">Compose</span></span><br><span data-ttu-id="ac55d-122">Read</span><span class="sxs-lookup"><span data-stu-id="ac55d-122">Read</span></span> | [<span data-ttu-id="ac55d-123">Auth</span><span class="sxs-lookup"><span data-stu-id="ac55d-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="ac55d-124">Visualização</span><span class="sxs-lookup"><span data-stu-id="ac55d-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="ac55d-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="ac55d-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="ac55d-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac55d-126">Compose</span></span><br><span data-ttu-id="ac55d-127">Read</span><span class="sxs-lookup"><span data-stu-id="ac55d-127">Read</span></span> | <span data-ttu-id="ac55d-128">String</span><span class="sxs-lookup"><span data-stu-id="ac55d-128">String</span></span> | [<span data-ttu-id="ac55d-129">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac55d-130">la</span><span class="sxs-lookup"><span data-stu-id="ac55d-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="ac55d-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac55d-131">Compose</span></span><br><span data-ttu-id="ac55d-132">Read</span><span class="sxs-lookup"><span data-stu-id="ac55d-132">Read</span></span> | [<span data-ttu-id="ac55d-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ac55d-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="ac55d-134">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac55d-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="ac55d-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="ac55d-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac55d-136">Compose</span></span><br><span data-ttu-id="ac55d-137">Read</span><span class="sxs-lookup"><span data-stu-id="ac55d-137">Read</span></span> | <span data-ttu-id="ac55d-138">String</span><span class="sxs-lookup"><span data-stu-id="ac55d-138">String</span></span> | [<span data-ttu-id="ac55d-139">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac55d-140">principal</span><span class="sxs-lookup"><span data-stu-id="ac55d-140">host</span></span>](#host-hosttype) | <span data-ttu-id="ac55d-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac55d-141">Compose</span></span><br><span data-ttu-id="ac55d-142">Read</span><span class="sxs-lookup"><span data-stu-id="ac55d-142">Read</span></span> | [<span data-ttu-id="ac55d-143">HostType</span><span class="sxs-lookup"><span data-stu-id="ac55d-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="ac55d-144">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac55d-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="ac55d-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="ac55d-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac55d-146">Compose</span></span><br><span data-ttu-id="ac55d-147">Read</span><span class="sxs-lookup"><span data-stu-id="ac55d-147">Read</span></span> | [<span data-ttu-id="ac55d-148">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="ac55d-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="ac55d-149">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac55d-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="ac55d-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="ac55d-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac55d-151">Compose</span></span><br><span data-ttu-id="ac55d-152">Read</span><span class="sxs-lookup"><span data-stu-id="ac55d-152">Read</span></span> | [<span data-ttu-id="ac55d-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="ac55d-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="ac55d-154">Visualização</span><span class="sxs-lookup"><span data-stu-id="ac55d-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="ac55d-155">plataforma</span><span class="sxs-lookup"><span data-stu-id="ac55d-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="ac55d-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac55d-156">Compose</span></span><br><span data-ttu-id="ac55d-157">Read</span><span class="sxs-lookup"><span data-stu-id="ac55d-157">Read</span></span> | [<span data-ttu-id="ac55d-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ac55d-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="ac55d-159">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac55d-160">atende</span><span class="sxs-lookup"><span data-stu-id="ac55d-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="ac55d-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac55d-161">Compose</span></span><br><span data-ttu-id="ac55d-162">Read</span><span class="sxs-lookup"><span data-stu-id="ac55d-162">Read</span></span> | [<span data-ttu-id="ac55d-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ac55d-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="ac55d-164">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac55d-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="ac55d-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="ac55d-166">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac55d-166">Compose</span></span><br><span data-ttu-id="ac55d-167">Read</span><span class="sxs-lookup"><span data-stu-id="ac55d-167">Read</span></span> | [<span data-ttu-id="ac55d-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ac55d-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="ac55d-169">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac55d-170">ui</span><span class="sxs-lookup"><span data-stu-id="ac55d-170">ui</span></span>](#ui-ui) | <span data-ttu-id="ac55d-171">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac55d-171">Compose</span></span><br><span data-ttu-id="ac55d-172">Read</span><span class="sxs-lookup"><span data-stu-id="ac55d-172">Read</span></span> | [<span data-ttu-id="ac55d-173">UI</span><span class="sxs-lookup"><span data-stu-id="ac55d-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="ac55d-174">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="ac55d-175">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="ac55d-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="ac55d-176">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="ac55d-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="ac55d-177">Oferece suporte a [logon único (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , fornecendo um método que permite que o host do Office obtenha um token de acesso para o aplicativo Web do suplemento.</span><span class="sxs-lookup"><span data-stu-id="ac55d-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="ac55d-178">Indiretamente, isso também habilita o suplemento para acessar os dados do Microsoft Graph do usuário sem exigir que o usuário se conecte uma segunda vez.</span><span class="sxs-lookup"><span data-stu-id="ac55d-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="ac55d-179">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac55d-179">Type</span></span>

*   [<span data-ttu-id="ac55d-180">Auth</span><span class="sxs-lookup"><span data-stu-id="ac55d-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="ac55d-181">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac55d-181">Requirements</span></span>

|<span data-ttu-id="ac55d-182">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac55d-182">Requirement</span></span>| <span data-ttu-id="ac55d-183">Valor</span><span class="sxs-lookup"><span data-stu-id="ac55d-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac55d-184">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac55d-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac55d-185">Visualização</span><span class="sxs-lookup"><span data-stu-id="ac55d-185">Preview</span></span>|
|[<span data-ttu-id="ac55d-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac55d-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac55d-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac55d-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac55d-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ac55d-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="ac55d-189">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ac55d-189">contentLanguage: String</span></span>

<span data-ttu-id="ac55d-190">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="ac55d-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="ac55d-191">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="ac55d-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="ac55d-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac55d-192">Type</span></span>

*   <span data-ttu-id="ac55d-193">String</span><span class="sxs-lookup"><span data-stu-id="ac55d-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac55d-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac55d-194">Requirements</span></span>

|<span data-ttu-id="ac55d-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac55d-195">Requirement</span></span>| <span data-ttu-id="ac55d-196">Valor</span><span class="sxs-lookup"><span data-stu-id="ac55d-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac55d-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac55d-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac55d-198">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-198">1.1</span></span>|
|[<span data-ttu-id="ac55d-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac55d-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac55d-200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac55d-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac55d-201">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ac55d-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="ac55d-202">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="ac55d-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="ac55d-203">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="ac55d-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ac55d-204">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac55d-204">Type</span></span>

*   [<span data-ttu-id="ac55d-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ac55d-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="ac55d-206">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac55d-206">Requirements</span></span>

|<span data-ttu-id="ac55d-207">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac55d-207">Requirement</span></span>| <span data-ttu-id="ac55d-208">Valor</span><span class="sxs-lookup"><span data-stu-id="ac55d-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac55d-209">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac55d-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac55d-210">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-210">1.1</span></span>|
|[<span data-ttu-id="ac55d-211">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac55d-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac55d-212">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac55d-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac55d-213">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ac55d-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="ac55d-214">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ac55d-214">displayLanguage: String</span></span>

<span data-ttu-id="ac55d-215">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="ac55d-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="ac55d-216">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="ac55d-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="ac55d-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac55d-217">Type</span></span>

*   <span data-ttu-id="ac55d-218">String</span><span class="sxs-lookup"><span data-stu-id="ac55d-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac55d-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac55d-219">Requirements</span></span>

|<span data-ttu-id="ac55d-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac55d-220">Requirement</span></span>| <span data-ttu-id="ac55d-221">Valor</span><span class="sxs-lookup"><span data-stu-id="ac55d-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac55d-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac55d-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac55d-223">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-223">1.1</span></span>|
|[<span data-ttu-id="ac55d-224">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac55d-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac55d-225">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac55d-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac55d-226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ac55d-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="ac55d-227">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="ac55d-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="ac55d-228">Obtém o host do aplicativo do Office no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="ac55d-228">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ac55d-229">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac55d-229">Type</span></span>

*   [<span data-ttu-id="ac55d-230">HostType</span><span class="sxs-lookup"><span data-stu-id="ac55d-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="ac55d-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac55d-231">Requirements</span></span>

|<span data-ttu-id="ac55d-232">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac55d-232">Requirement</span></span>| <span data-ttu-id="ac55d-233">Valor</span><span class="sxs-lookup"><span data-stu-id="ac55d-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac55d-234">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac55d-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac55d-235">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-235">1.1</span></span>|
|[<span data-ttu-id="ac55d-236">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac55d-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac55d-237">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac55d-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac55d-238">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ac55d-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="ac55d-239">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="ac55d-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="ac55d-240">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="ac55d-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="ac55d-241">Só há suporte para esse membro no Outlook no Windows.</span><span class="sxs-lookup"><span data-stu-id="ac55d-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="ac55d-242">O uso de cores de temas do Office permite coordenar o esquema de cores do seu suplemento com o tema atual do Office selecionado pelo usuário com a **conta de arquivo > office > Office Theme UI**, que é aplicada em todos os aplicativos host do Office.</span><span class="sxs-lookup"><span data-stu-id="ac55d-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="ac55d-243">Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="ac55d-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ac55d-244">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac55d-244">Type</span></span>

*   [<span data-ttu-id="ac55d-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="ac55d-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="ac55d-246">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ac55d-246">Properties:</span></span>

|<span data-ttu-id="ac55d-247">Nome</span><span class="sxs-lookup"><span data-stu-id="ac55d-247">Name</span></span>| <span data-ttu-id="ac55d-248">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac55d-248">Type</span></span>| <span data-ttu-id="ac55d-249">Descrição</span><span class="sxs-lookup"><span data-stu-id="ac55d-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="ac55d-250">String</span><span class="sxs-lookup"><span data-stu-id="ac55d-250">String</span></span>|<span data-ttu-id="ac55d-251">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="ac55d-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="ac55d-252">String</span><span class="sxs-lookup"><span data-stu-id="ac55d-252">String</span></span>|<span data-ttu-id="ac55d-253">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="ac55d-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="ac55d-254">String</span><span class="sxs-lookup"><span data-stu-id="ac55d-254">String</span></span>|<span data-ttu-id="ac55d-255">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="ac55d-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="ac55d-256">String</span><span class="sxs-lookup"><span data-stu-id="ac55d-256">String</span></span>|<span data-ttu-id="ac55d-257">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="ac55d-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ac55d-258">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac55d-258">Requirements</span></span>

|<span data-ttu-id="ac55d-259">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac55d-259">Requirement</span></span>| <span data-ttu-id="ac55d-260">Valor</span><span class="sxs-lookup"><span data-stu-id="ac55d-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac55d-261">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac55d-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac55d-262">Visualização</span><span class="sxs-lookup"><span data-stu-id="ac55d-262">Preview</span></span>|
|[<span data-ttu-id="ac55d-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac55d-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac55d-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac55d-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac55d-265">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ac55d-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="ac55d-266">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="ac55d-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="ac55d-267">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="ac55d-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ac55d-268">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac55d-268">Type</span></span>

*   [<span data-ttu-id="ac55d-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ac55d-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="ac55d-270">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac55d-270">Requirements</span></span>

|<span data-ttu-id="ac55d-271">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac55d-271">Requirement</span></span>| <span data-ttu-id="ac55d-272">Valor</span><span class="sxs-lookup"><span data-stu-id="ac55d-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac55d-273">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac55d-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac55d-274">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-274">1.1</span></span>|
|[<span data-ttu-id="ac55d-275">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac55d-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac55d-276">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac55d-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac55d-277">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ac55d-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="ac55d-278">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="ac55d-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="ac55d-279">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o host atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="ac55d-279">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="ac55d-280">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac55d-280">Type</span></span>

*   [<span data-ttu-id="ac55d-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ac55d-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="ac55d-282">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac55d-282">Requirements</span></span>

|<span data-ttu-id="ac55d-283">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac55d-283">Requirement</span></span>| <span data-ttu-id="ac55d-284">Valor</span><span class="sxs-lookup"><span data-stu-id="ac55d-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac55d-285">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac55d-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac55d-286">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-286">1.1</span></span>|
|[<span data-ttu-id="ac55d-287">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac55d-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac55d-288">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac55d-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac55d-289">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ac55d-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="ac55d-290">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="ac55d-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="ac55d-291">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="ac55d-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ac55d-292">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="ac55d-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ac55d-293">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac55d-293">Type</span></span>

*   [<span data-ttu-id="ac55d-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ac55d-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="ac55d-295">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac55d-295">Requirements</span></span>

|<span data-ttu-id="ac55d-296">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac55d-296">Requirement</span></span>| <span data-ttu-id="ac55d-297">Valor</span><span class="sxs-lookup"><span data-stu-id="ac55d-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac55d-298">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac55d-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac55d-299">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-299">1.1</span></span>|
|[<span data-ttu-id="ac55d-300">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ac55d-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="ac55d-301">Restrito</span><span class="sxs-lookup"><span data-stu-id="ac55d-301">Restricted</span></span>|
|[<span data-ttu-id="ac55d-302">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac55d-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac55d-303">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac55d-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="ac55d-304">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="ac55d-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="ac55d-305">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="ac55d-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ac55d-306">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac55d-306">Type</span></span>

*   [<span data-ttu-id="ac55d-307">UI</span><span class="sxs-lookup"><span data-stu-id="ac55d-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="ac55d-308">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac55d-308">Requirements</span></span>

|<span data-ttu-id="ac55d-309">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac55d-309">Requirement</span></span>| <span data-ttu-id="ac55d-310">Valor</span><span class="sxs-lookup"><span data-stu-id="ac55d-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac55d-311">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac55d-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac55d-312">1.1</span><span class="sxs-lookup"><span data-stu-id="ac55d-312">1.1</span></span>|
|[<span data-ttu-id="ac55d-313">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac55d-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac55d-314">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac55d-314">Compose or Read</span></span>|
