---
title: Office.context – conjunto de requisitos 1.6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: d57eae6017d2eb885ebc524f70490787e8b1d221
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111188"
---
# <a name="context"></a><span data-ttu-id="e1660-102">context</span><span class="sxs-lookup"><span data-stu-id="e1660-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="e1660-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="e1660-103">[Office](office.md).context</span></span>

<span data-ttu-id="e1660-104">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="e1660-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="e1660-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.6).</span><span class="sxs-lookup"><span data-stu-id="e1660-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.6).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1660-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1660-106">Requirements</span></span>

|<span data-ttu-id="e1660-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1660-107">Requirement</span></span>| <span data-ttu-id="e1660-108">Valor</span><span class="sxs-lookup"><span data-stu-id="e1660-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1660-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1660-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e1660-110">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-110">1.1</span></span>|
|[<span data-ttu-id="e1660-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1660-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e1660-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1660-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="e1660-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="e1660-113">Properties</span></span>

| <span data-ttu-id="e1660-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="e1660-114">Property</span></span> | <span data-ttu-id="e1660-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="e1660-115">Modes</span></span> | <span data-ttu-id="e1660-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="e1660-116">Return type</span></span> | <span data-ttu-id="e1660-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="e1660-117">Minimum</span></span><br><span data-ttu-id="e1660-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="e1660-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e1660-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="e1660-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="e1660-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1660-120">Compose</span></span><br><span data-ttu-id="e1660-121">Leitura</span><span class="sxs-lookup"><span data-stu-id="e1660-121">Read</span></span> | <span data-ttu-id="e1660-122">String</span><span class="sxs-lookup"><span data-stu-id="e1660-122">String</span></span> | [<span data-ttu-id="e1660-123">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e1660-124">la</span><span class="sxs-lookup"><span data-stu-id="e1660-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="e1660-125">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1660-125">Compose</span></span><br><span data-ttu-id="e1660-126">Leitura</span><span class="sxs-lookup"><span data-stu-id="e1660-126">Read</span></span> | [<span data-ttu-id="e1660-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="e1660-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.6) | [<span data-ttu-id="e1660-128">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e1660-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="e1660-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="e1660-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1660-130">Compose</span></span><br><span data-ttu-id="e1660-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="e1660-131">Read</span></span> | <span data-ttu-id="e1660-132">String</span><span class="sxs-lookup"><span data-stu-id="e1660-132">String</span></span> | [<span data-ttu-id="e1660-133">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e1660-134">principal</span><span class="sxs-lookup"><span data-stu-id="e1660-134">host</span></span>](#host-hosttype) | <span data-ttu-id="e1660-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1660-135">Compose</span></span><br><span data-ttu-id="e1660-136">Leitura</span><span class="sxs-lookup"><span data-stu-id="e1660-136">Read</span></span> | [<span data-ttu-id="e1660-137">HostType</span><span class="sxs-lookup"><span data-stu-id="e1660-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.6) | [<span data-ttu-id="e1660-138">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e1660-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="e1660-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="e1660-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1660-140">Compose</span></span><br><span data-ttu-id="e1660-141">Leitura</span><span class="sxs-lookup"><span data-stu-id="e1660-141">Read</span></span> | [<span data-ttu-id="e1660-142">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="e1660-142">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6) | [<span data-ttu-id="e1660-143">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e1660-144">plataforma</span><span class="sxs-lookup"><span data-stu-id="e1660-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="e1660-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1660-145">Compose</span></span><br><span data-ttu-id="e1660-146">Leitura</span><span class="sxs-lookup"><span data-stu-id="e1660-146">Read</span></span> | [<span data-ttu-id="e1660-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="e1660-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.6) | [<span data-ttu-id="e1660-148">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e1660-149">atende</span><span class="sxs-lookup"><span data-stu-id="e1660-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="e1660-150">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1660-150">Compose</span></span><br><span data-ttu-id="e1660-151">Leitura</span><span class="sxs-lookup"><span data-stu-id="e1660-151">Read</span></span> | [<span data-ttu-id="e1660-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="e1660-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.6) | [<span data-ttu-id="e1660-153">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e1660-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="e1660-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="e1660-155">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1660-155">Compose</span></span><br><span data-ttu-id="e1660-156">Leitura</span><span class="sxs-lookup"><span data-stu-id="e1660-156">Read</span></span> | [<span data-ttu-id="e1660-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e1660-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.6) | [<span data-ttu-id="e1660-158">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e1660-159">ui</span><span class="sxs-lookup"><span data-stu-id="e1660-159">ui</span></span>](#ui-ui) | <span data-ttu-id="e1660-160">Escrever</span><span class="sxs-lookup"><span data-stu-id="e1660-160">Compose</span></span><br><span data-ttu-id="e1660-161">Leitura</span><span class="sxs-lookup"><span data-stu-id="e1660-161">Read</span></span> | [<span data-ttu-id="e1660-162">UI</span><span class="sxs-lookup"><span data-stu-id="e1660-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.6) | [<span data-ttu-id="e1660-163">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="e1660-164">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="e1660-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="e1660-165">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e1660-165">contentLanguage: String</span></span>

<span data-ttu-id="e1660-166">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="e1660-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="e1660-167">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="e1660-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="e1660-168">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1660-168">Type</span></span>

*   <span data-ttu-id="e1660-169">String</span><span class="sxs-lookup"><span data-stu-id="e1660-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1660-170">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1660-170">Requirements</span></span>

|<span data-ttu-id="e1660-171">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1660-171">Requirement</span></span>| <span data-ttu-id="e1660-172">Valor</span><span class="sxs-lookup"><span data-stu-id="e1660-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1660-173">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1660-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e1660-174">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-174">1.1</span></span>|
|[<span data-ttu-id="e1660-175">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1660-175">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e1660-176">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1660-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1660-177">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1660-177">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="e1660-178">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="e1660-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="e1660-179">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="e1660-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="e1660-180">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1660-180">Type</span></span>

*   [<span data-ttu-id="e1660-181">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="e1660-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="e1660-182">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1660-182">Requirements</span></span>

|<span data-ttu-id="e1660-183">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1660-183">Requirement</span></span>| <span data-ttu-id="e1660-184">Valor</span><span class="sxs-lookup"><span data-stu-id="e1660-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1660-185">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1660-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e1660-186">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-186">1.1</span></span>|
|[<span data-ttu-id="e1660-187">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1660-187">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e1660-188">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1660-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1660-189">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1660-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="e1660-190">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e1660-190">displayLanguage: String</span></span>

<span data-ttu-id="e1660-191">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="e1660-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="e1660-192">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="e1660-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="e1660-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1660-193">Type</span></span>

*   <span data-ttu-id="e1660-194">String</span><span class="sxs-lookup"><span data-stu-id="e1660-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e1660-195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1660-195">Requirements</span></span>

|<span data-ttu-id="e1660-196">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1660-196">Requirement</span></span>| <span data-ttu-id="e1660-197">Valor</span><span class="sxs-lookup"><span data-stu-id="e1660-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1660-198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1660-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e1660-199">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-199">1.1</span></span>|
|[<span data-ttu-id="e1660-200">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1660-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e1660-201">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1660-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1660-202">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1660-202">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="e1660-203">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="e1660-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="e1660-204">Obtém o host do aplicativo do Office no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="e1660-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="e1660-205">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1660-205">Type</span></span>

*   [<span data-ttu-id="e1660-206">HostType</span><span class="sxs-lookup"><span data-stu-id="e1660-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="e1660-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1660-207">Requirements</span></span>

|<span data-ttu-id="e1660-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1660-208">Requirement</span></span>| <span data-ttu-id="e1660-209">Valor</span><span class="sxs-lookup"><span data-stu-id="e1660-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1660-210">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1660-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e1660-211">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-211">1.1</span></span>|
|[<span data-ttu-id="e1660-212">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1660-212">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e1660-213">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1660-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1660-214">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1660-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="e1660-215">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="e1660-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="e1660-216">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="e1660-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="e1660-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1660-217">Type</span></span>

*   [<span data-ttu-id="e1660-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="e1660-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="e1660-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1660-219">Requirements</span></span>

|<span data-ttu-id="e1660-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1660-220">Requirement</span></span>| <span data-ttu-id="e1660-221">Valor</span><span class="sxs-lookup"><span data-stu-id="e1660-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1660-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1660-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e1660-223">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-223">1.1</span></span>|
|[<span data-ttu-id="e1660-224">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1660-224">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e1660-225">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1660-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1660-226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1660-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="e1660-227">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="e1660-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="e1660-228">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o host atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="e1660-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="e1660-229">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1660-229">Type</span></span>

*   [<span data-ttu-id="e1660-230">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="e1660-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="e1660-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1660-231">Requirements</span></span>

|<span data-ttu-id="e1660-232">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1660-232">Requirement</span></span>| <span data-ttu-id="e1660-233">Valor</span><span class="sxs-lookup"><span data-stu-id="e1660-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1660-234">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1660-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e1660-235">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-235">1.1</span></span>|
|[<span data-ttu-id="e1660-236">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1660-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e1660-237">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1660-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e1660-238">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e1660-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="e1660-239">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="e1660-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="e1660-240">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="e1660-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="e1660-241">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="e1660-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="e1660-242">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1660-242">Type</span></span>

*   [<span data-ttu-id="e1660-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e1660-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="e1660-244">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1660-244">Requirements</span></span>

|<span data-ttu-id="e1660-245">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1660-245">Requirement</span></span>| <span data-ttu-id="e1660-246">Valor</span><span class="sxs-lookup"><span data-stu-id="e1660-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1660-247">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1660-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e1660-248">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-248">1.1</span></span>|
|[<span data-ttu-id="e1660-249">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="e1660-249">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e1660-250">Restrito</span><span class="sxs-lookup"><span data-stu-id="e1660-250">Restricted</span></span>|
|[<span data-ttu-id="e1660-251">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1660-251">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e1660-252">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1660-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="e1660-253">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="e1660-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="e1660-254">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="e1660-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="e1660-255">Tipo</span><span class="sxs-lookup"><span data-stu-id="e1660-255">Type</span></span>

*   [<span data-ttu-id="e1660-256">UI</span><span class="sxs-lookup"><span data-stu-id="e1660-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="e1660-257">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e1660-257">Requirements</span></span>

|<span data-ttu-id="e1660-258">Requisito</span><span class="sxs-lookup"><span data-stu-id="e1660-258">Requirement</span></span>| <span data-ttu-id="e1660-259">Valor</span><span class="sxs-lookup"><span data-stu-id="e1660-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="e1660-260">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e1660-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e1660-261">1.1</span><span class="sxs-lookup"><span data-stu-id="e1660-261">1.1</span></span>|
|[<span data-ttu-id="e1660-262">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e1660-262">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e1660-263">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e1660-263">Compose or Read</span></span>|
