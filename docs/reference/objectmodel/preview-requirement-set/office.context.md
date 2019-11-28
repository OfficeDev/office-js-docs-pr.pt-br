---
title: Office. Context – conjunto de requisitos de visualização
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 5c34a7a0db5880a94ba5519059a93010a5243978
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629185"
---
# <a name="context"></a><span data-ttu-id="b9cb0-102">context</span><span class="sxs-lookup"><span data-stu-id="b9cb0-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="b9cb0-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="b9cb0-103">[Office](Office.md).context</span></span>

<span data-ttu-id="b9cb0-104">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-104">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="b9cb0-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="b9cb0-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9cb0-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9cb0-106">Requirements</span></span>

|<span data-ttu-id="b9cb0-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9cb0-107">Requirement</span></span>| <span data-ttu-id="b9cb0-108">Valor</span><span class="sxs-lookup"><span data-stu-id="b9cb0-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9cb0-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9cb0-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9cb0-110">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-110">1.0</span></span>|
|[<span data-ttu-id="b9cb0-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9cb0-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9cb0-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b9cb0-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="b9cb0-113">Properties</span></span>

| <span data-ttu-id="b9cb0-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="b9cb0-114">Property</span></span> | <span data-ttu-id="b9cb0-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="b9cb0-115">Modes</span></span> | <span data-ttu-id="b9cb0-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="b9cb0-116">Return type</span></span> | <span data-ttu-id="b9cb0-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-117">Minimum</span></span><br><span data-ttu-id="b9cb0-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="b9cb0-118">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="b9cb0-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="b9cb0-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="b9cb0-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="b9cb0-120">Compose</span></span><br><span data-ttu-id="b9cb0-121">Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-121">Read</span></span> | <span data-ttu-id="b9cb0-122">String</span><span class="sxs-lookup"><span data-stu-id="b9cb0-122">String</span></span> | <span data-ttu-id="b9cb0-123">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-123">1.0</span></span> |
| [<span data-ttu-id="b9cb0-124">la</span><span class="sxs-lookup"><span data-stu-id="b9cb0-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="b9cb0-125">Escrever</span><span class="sxs-lookup"><span data-stu-id="b9cb0-125">Compose</span></span><br><span data-ttu-id="b9cb0-126">Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-126">Read</span></span> | [<span data-ttu-id="b9cb0-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b9cb0-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation) | <span data-ttu-id="b9cb0-128">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-128">1.0</span></span> |
| [<span data-ttu-id="b9cb0-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="b9cb0-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="b9cb0-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="b9cb0-130">Compose</span></span><br><span data-ttu-id="b9cb0-131">Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-131">Read</span></span> | <span data-ttu-id="b9cb0-132">String</span><span class="sxs-lookup"><span data-stu-id="b9cb0-132">String</span></span> | <span data-ttu-id="b9cb0-133">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-133">1.0</span></span> |
| [<span data-ttu-id="b9cb0-134">principal</span><span class="sxs-lookup"><span data-stu-id="b9cb0-134">host</span></span>](#host-hosttype) | <span data-ttu-id="b9cb0-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="b9cb0-135">Compose</span></span><br><span data-ttu-id="b9cb0-136">Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-136">Read</span></span> | [<span data-ttu-id="b9cb0-137">HostType</span><span class="sxs-lookup"><span data-stu-id="b9cb0-137">HostType</span></span>](/javascript/api/office/office.hosttype) | <span data-ttu-id="b9cb0-138">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-138">1.0</span></span> |
| [<span data-ttu-id="b9cb0-139">officeTheme</span><span class="sxs-lookup"><span data-stu-id="b9cb0-139">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="b9cb0-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="b9cb0-140">Compose</span></span><br><span data-ttu-id="b9cb0-141">Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-141">Read</span></span> | [<span data-ttu-id="b9cb0-142">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="b9cb0-142">OfficeTheme</span></span>](/javascript/api/office/office.officetheme) | <span data-ttu-id="b9cb0-143">Visualização</span><span class="sxs-lookup"><span data-stu-id="b9cb0-143">Preview</span></span> |
| [<span data-ttu-id="b9cb0-144">plataforma</span><span class="sxs-lookup"><span data-stu-id="b9cb0-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="b9cb0-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="b9cb0-145">Compose</span></span><br><span data-ttu-id="b9cb0-146">Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-146">Read</span></span> | [<span data-ttu-id="b9cb0-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b9cb0-147">PlatformType</span></span>](/javascript/api/office/office.platformtype) | <span data-ttu-id="b9cb0-148">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-148">1.0</span></span> |
| [<span data-ttu-id="b9cb0-149">atende</span><span class="sxs-lookup"><span data-stu-id="b9cb0-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="b9cb0-150">Escrever</span><span class="sxs-lookup"><span data-stu-id="b9cb0-150">Compose</span></span><br><span data-ttu-id="b9cb0-151">Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-151">Read</span></span> | [<span data-ttu-id="b9cb0-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b9cb0-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport) | <span data-ttu-id="b9cb0-153">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-153">1.0</span></span> |
| [<span data-ttu-id="b9cb0-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="b9cb0-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="b9cb0-155">Escrever</span><span class="sxs-lookup"><span data-stu-id="b9cb0-155">Compose</span></span><br><span data-ttu-id="b9cb0-156">Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-156">Read</span></span> | [<span data-ttu-id="b9cb0-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b9cb0-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings) | <span data-ttu-id="b9cb0-158">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-158">1.0</span></span> |
| [<span data-ttu-id="b9cb0-159">ui</span><span class="sxs-lookup"><span data-stu-id="b9cb0-159">ui</span></span>](#ui-ui) | <span data-ttu-id="b9cb0-160">Escrever</span><span class="sxs-lookup"><span data-stu-id="b9cb0-160">Compose</span></span><br><span data-ttu-id="b9cb0-161">Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-161">Read</span></span> | [<span data-ttu-id="b9cb0-162">UI</span><span class="sxs-lookup"><span data-stu-id="b9cb0-162">UI</span></span>](/javascript/api/office/office.ui) | <span data-ttu-id="b9cb0-163">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-163">1.0</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b9cb0-164">Namespaces</span><span class="sxs-lookup"><span data-stu-id="b9cb0-164">Namespaces</span></span>

<span data-ttu-id="b9cb0-165">[auth](/javascript/api/office/office.auth): fornece suporte para [logon único (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token).</span><span class="sxs-lookup"><span data-stu-id="b9cb0-165">[auth](/javascript/api/office/office.auth): Provides support for [single sign-on (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token).</span></span>

<span data-ttu-id="b9cb0-166">[Mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-166">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

## <a name="property-details"></a><span data-ttu-id="b9cb0-167">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="b9cb0-167">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="b9cb0-168">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b9cb0-168">contentLanguage: String</span></span>

<span data-ttu-id="b9cb0-169">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-169">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="b9cb0-170">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-170">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="b9cb0-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-171">Type</span></span>

*   <span data-ttu-id="b9cb0-172">String</span><span class="sxs-lookup"><span data-stu-id="b9cb0-172">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9cb0-173">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9cb0-173">Requirements</span></span>

|<span data-ttu-id="b9cb0-174">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9cb0-174">Requirement</span></span>| <span data-ttu-id="b9cb0-175">Valor</span><span class="sxs-lookup"><span data-stu-id="b9cb0-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9cb0-176">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9cb0-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9cb0-177">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-177">1.0</span></span>|
|[<span data-ttu-id="b9cb0-178">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9cb0-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9cb0-179">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9cb0-180">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-180">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="b9cb0-181">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="b9cb0-181">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="b9cb0-182">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-182">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b9cb0-183">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-183">Type</span></span>

*   [<span data-ttu-id="b9cb0-184">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b9cb0-184">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="b9cb0-185">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9cb0-185">Requirements</span></span>

|<span data-ttu-id="b9cb0-186">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9cb0-186">Requirement</span></span>| <span data-ttu-id="b9cb0-187">Valor</span><span class="sxs-lookup"><span data-stu-id="b9cb0-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9cb0-188">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9cb0-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9cb0-189">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-189">1.0</span></span>|
|[<span data-ttu-id="b9cb0-190">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9cb0-190">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9cb0-191">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-191">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9cb0-192">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-192">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="b9cb0-193">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b9cb0-193">displayLanguage: String</span></span>

<span data-ttu-id="b9cb0-194">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-194">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="b9cb0-195">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-195">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="b9cb0-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-196">Type</span></span>

*   <span data-ttu-id="b9cb0-197">String</span><span class="sxs-lookup"><span data-stu-id="b9cb0-197">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9cb0-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9cb0-198">Requirements</span></span>

|<span data-ttu-id="b9cb0-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9cb0-199">Requirement</span></span>| <span data-ttu-id="b9cb0-200">Valor</span><span class="sxs-lookup"><span data-stu-id="b9cb0-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9cb0-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9cb0-201">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9cb0-202">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-202">1.0</span></span>|
|[<span data-ttu-id="b9cb0-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9cb0-203">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9cb0-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9cb0-205">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-205">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="b9cb0-206">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="b9cb0-206">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="b9cb0-207">Obtém o host do aplicativo do Office no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-207">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b9cb0-208">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-208">Type</span></span>

*   [<span data-ttu-id="b9cb0-209">HostType</span><span class="sxs-lookup"><span data-stu-id="b9cb0-209">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="b9cb0-210">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9cb0-210">Requirements</span></span>

|<span data-ttu-id="b9cb0-211">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9cb0-211">Requirement</span></span>| <span data-ttu-id="b9cb0-212">Valor</span><span class="sxs-lookup"><span data-stu-id="b9cb0-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9cb0-213">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9cb0-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9cb0-214">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-214">1.0</span></span>|
|[<span data-ttu-id="b9cb0-215">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9cb0-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9cb0-216">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-216">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9cb0-217">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-217">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officethemejavascriptapiofficeofficeofficetheme"></a><span data-ttu-id="b9cb0-218">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="b9cb0-218">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="b9cb0-219">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-219">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="b9cb0-220">Só há suporte para esse membro no Outlook no Windows.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-220">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="b9cb0-221">O uso de cores de temas do Office permite coordenar o esquema de cores do seu suplemento com o tema atual do Office selecionado pelo usuário com a **conta de arquivo > office > Office Theme UI**, que é aplicada em todos os aplicativos host do Office.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-221">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="b9cb0-222">Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-222">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="b9cb0-223">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-223">Type</span></span>

*   [<span data-ttu-id="b9cb0-224">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="b9cb0-224">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="b9cb0-225">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b9cb0-225">Properties:</span></span>

|<span data-ttu-id="b9cb0-226">Nome</span><span class="sxs-lookup"><span data-stu-id="b9cb0-226">Name</span></span>| <span data-ttu-id="b9cb0-227">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-227">Type</span></span>| <span data-ttu-id="b9cb0-228">Descrição</span><span class="sxs-lookup"><span data-stu-id="b9cb0-228">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="b9cb0-229">String</span><span class="sxs-lookup"><span data-stu-id="b9cb0-229">String</span></span>|<span data-ttu-id="b9cb0-230">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-230">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="b9cb0-231">String</span><span class="sxs-lookup"><span data-stu-id="b9cb0-231">String</span></span>|<span data-ttu-id="b9cb0-232">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-232">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="b9cb0-233">String</span><span class="sxs-lookup"><span data-stu-id="b9cb0-233">String</span></span>|<span data-ttu-id="b9cb0-234">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-234">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="b9cb0-235">String</span><span class="sxs-lookup"><span data-stu-id="b9cb0-235">String</span></span>|<span data-ttu-id="b9cb0-236">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-236">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9cb0-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9cb0-237">Requirements</span></span>

|<span data-ttu-id="b9cb0-238">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9cb0-238">Requirement</span></span>| <span data-ttu-id="b9cb0-239">Valor</span><span class="sxs-lookup"><span data-stu-id="b9cb0-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9cb0-240">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9cb0-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9cb0-241">Visualização</span><span class="sxs-lookup"><span data-stu-id="b9cb0-241">Preview</span></span>|
|[<span data-ttu-id="b9cb0-242">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9cb0-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9cb0-243">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-243">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9cb0-244">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-244">Example</span></span>

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

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="b9cb0-245">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="b9cb0-245">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="b9cb0-246">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-246">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b9cb0-247">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-247">Type</span></span>

*   [<span data-ttu-id="b9cb0-248">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b9cb0-248">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="b9cb0-249">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9cb0-249">Requirements</span></span>

|<span data-ttu-id="b9cb0-250">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9cb0-250">Requirement</span></span>| <span data-ttu-id="b9cb0-251">Valor</span><span class="sxs-lookup"><span data-stu-id="b9cb0-251">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9cb0-252">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9cb0-252">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9cb0-253">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-253">1.0</span></span>|
|[<span data-ttu-id="b9cb0-254">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9cb0-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9cb0-255">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-255">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9cb0-256">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-256">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="b9cb0-257">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="b9cb0-257">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="b9cb0-258">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o host atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-258">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b9cb0-259">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-259">Type</span></span>

*   [<span data-ttu-id="b9cb0-260">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b9cb0-260">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="b9cb0-261">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9cb0-261">Requirements</span></span>

|<span data-ttu-id="b9cb0-262">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9cb0-262">Requirement</span></span>| <span data-ttu-id="b9cb0-263">Valor</span><span class="sxs-lookup"><span data-stu-id="b9cb0-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9cb0-264">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9cb0-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9cb0-265">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-265">1.0</span></span>|
|[<span data-ttu-id="b9cb0-266">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9cb0-266">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9cb0-267">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-267">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9cb0-268">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-268">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.8")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="b9cb0-269">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="b9cb0-269">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="b9cb0-270">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-270">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="b9cb0-271">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-271">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="b9cb0-272">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-272">Type</span></span>

*   [<span data-ttu-id="b9cb0-273">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b9cb0-273">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="b9cb0-274">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9cb0-274">Requirements</span></span>

|<span data-ttu-id="b9cb0-275">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9cb0-275">Requirement</span></span>| <span data-ttu-id="b9cb0-276">Valor</span><span class="sxs-lookup"><span data-stu-id="b9cb0-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9cb0-277">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9cb0-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9cb0-278">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-278">1.0</span></span>|
|[<span data-ttu-id="b9cb0-279">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-279">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b9cb0-280">Restrito</span><span class="sxs-lookup"><span data-stu-id="b9cb0-280">Restricted</span></span>|
|[<span data-ttu-id="b9cb0-281">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9cb0-281">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9cb0-282">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-282">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="b9cb0-283">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="b9cb0-283">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="b9cb0-284">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="b9cb0-284">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="b9cb0-285">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9cb0-285">Type</span></span>

*   [<span data-ttu-id="b9cb0-286">UI</span><span class="sxs-lookup"><span data-stu-id="b9cb0-286">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="b9cb0-287">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9cb0-287">Requirements</span></span>

|<span data-ttu-id="b9cb0-288">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9cb0-288">Requirement</span></span>| <span data-ttu-id="b9cb0-289">Valor</span><span class="sxs-lookup"><span data-stu-id="b9cb0-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9cb0-290">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9cb0-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9cb0-291">1.0</span><span class="sxs-lookup"><span data-stu-id="b9cb0-291">1.0</span></span>|
|[<span data-ttu-id="b9cb0-292">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9cb0-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9cb0-293">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9cb0-293">Compose or Read</span></span>|
