---
title: Office. Context – conjunto de requisitos 1,3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 79e65c08e5b3ed21d618423f521010d3041d62af
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814903"
---
# <a name="context"></a><span data-ttu-id="133d8-102">context</span><span class="sxs-lookup"><span data-stu-id="133d8-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="133d8-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="133d8-103">[Office](office.md).context</span></span>

<span data-ttu-id="133d8-104">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="133d8-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="133d8-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.3).</span><span class="sxs-lookup"><span data-stu-id="133d8-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3).</span></span>

##### <a name="requirements"></a><span data-ttu-id="133d8-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="133d8-106">Requirements</span></span>

|<span data-ttu-id="133d8-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="133d8-107">Requirement</span></span>| <span data-ttu-id="133d8-108">Valor</span><span class="sxs-lookup"><span data-stu-id="133d8-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="133d8-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="133d8-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="133d8-110">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-110">1.1</span></span>|
|[<span data-ttu-id="133d8-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="133d8-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="133d8-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="133d8-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="133d8-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="133d8-113">Properties</span></span>

| <span data-ttu-id="133d8-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="133d8-114">Property</span></span> | <span data-ttu-id="133d8-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="133d8-115">Modes</span></span> | <span data-ttu-id="133d8-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="133d8-116">Return type</span></span> | <span data-ttu-id="133d8-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="133d8-117">Minimum</span></span><br><span data-ttu-id="133d8-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="133d8-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="133d8-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="133d8-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="133d8-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="133d8-120">Compose</span></span><br><span data-ttu-id="133d8-121">Leitura</span><span class="sxs-lookup"><span data-stu-id="133d8-121">Read</span></span> | <span data-ttu-id="133d8-122">String</span><span class="sxs-lookup"><span data-stu-id="133d8-122">String</span></span> | [<span data-ttu-id="133d8-123">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="133d8-124">la</span><span class="sxs-lookup"><span data-stu-id="133d8-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="133d8-125">Escrever</span><span class="sxs-lookup"><span data-stu-id="133d8-125">Compose</span></span><br><span data-ttu-id="133d8-126">Leitura</span><span class="sxs-lookup"><span data-stu-id="133d8-126">Read</span></span> | [<span data-ttu-id="133d8-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="133d8-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3) | [<span data-ttu-id="133d8-128">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="133d8-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="133d8-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="133d8-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="133d8-130">Compose</span></span><br><span data-ttu-id="133d8-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="133d8-131">Read</span></span> | <span data-ttu-id="133d8-132">String</span><span class="sxs-lookup"><span data-stu-id="133d8-132">String</span></span> | [<span data-ttu-id="133d8-133">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="133d8-134">principal</span><span class="sxs-lookup"><span data-stu-id="133d8-134">host</span></span>](#host-hosttype) | <span data-ttu-id="133d8-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="133d8-135">Compose</span></span><br><span data-ttu-id="133d8-136">Leitura</span><span class="sxs-lookup"><span data-stu-id="133d8-136">Read</span></span> | [<span data-ttu-id="133d8-137">HostType</span><span class="sxs-lookup"><span data-stu-id="133d8-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.3) | [<span data-ttu-id="133d8-138">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="133d8-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="133d8-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="133d8-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="133d8-140">Compose</span></span><br><span data-ttu-id="133d8-141">Leitura</span><span class="sxs-lookup"><span data-stu-id="133d8-141">Read</span></span> | [<span data-ttu-id="133d8-142">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="133d8-142">Mailbox</span></span>](/javascript/api/office/office.mailbox?view=outlook-js-1.3) | [<span data-ttu-id="133d8-143">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="133d8-144">plataforma</span><span class="sxs-lookup"><span data-stu-id="133d8-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="133d8-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="133d8-145">Compose</span></span><br><span data-ttu-id="133d8-146">Leitura</span><span class="sxs-lookup"><span data-stu-id="133d8-146">Read</span></span> | [<span data-ttu-id="133d8-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="133d8-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.3) | [<span data-ttu-id="133d8-148">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="133d8-149">atende</span><span class="sxs-lookup"><span data-stu-id="133d8-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="133d8-150">Escrever</span><span class="sxs-lookup"><span data-stu-id="133d8-150">Compose</span></span><br><span data-ttu-id="133d8-151">Leitura</span><span class="sxs-lookup"><span data-stu-id="133d8-151">Read</span></span> | [<span data-ttu-id="133d8-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="133d8-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3) | [<span data-ttu-id="133d8-153">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="133d8-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="133d8-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="133d8-155">Escrever</span><span class="sxs-lookup"><span data-stu-id="133d8-155">Compose</span></span><br><span data-ttu-id="133d8-156">Leitura</span><span class="sxs-lookup"><span data-stu-id="133d8-156">Read</span></span> | [<span data-ttu-id="133d8-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="133d8-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3) | [<span data-ttu-id="133d8-158">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="133d8-159">ui</span><span class="sxs-lookup"><span data-stu-id="133d8-159">ui</span></span>](#ui-ui) | <span data-ttu-id="133d8-160">Escrever</span><span class="sxs-lookup"><span data-stu-id="133d8-160">Compose</span></span><br><span data-ttu-id="133d8-161">Leitura</span><span class="sxs-lookup"><span data-stu-id="133d8-161">Read</span></span> | [<span data-ttu-id="133d8-162">UI</span><span class="sxs-lookup"><span data-stu-id="133d8-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3) | [<span data-ttu-id="133d8-163">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="133d8-164">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="133d8-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="133d8-165">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="133d8-165">contentLanguage: String</span></span>

<span data-ttu-id="133d8-166">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="133d8-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="133d8-167">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="133d8-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="133d8-168">Tipo</span><span class="sxs-lookup"><span data-stu-id="133d8-168">Type</span></span>

*   <span data-ttu-id="133d8-169">String</span><span class="sxs-lookup"><span data-stu-id="133d8-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="133d8-170">Requisitos</span><span class="sxs-lookup"><span data-stu-id="133d8-170">Requirements</span></span>

|<span data-ttu-id="133d8-171">Requisito</span><span class="sxs-lookup"><span data-stu-id="133d8-171">Requirement</span></span>| <span data-ttu-id="133d8-172">Valor</span><span class="sxs-lookup"><span data-stu-id="133d8-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="133d8-173">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="133d8-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="133d8-174">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-174">1.1</span></span>|
|[<span data-ttu-id="133d8-175">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="133d8-175">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="133d8-176">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="133d8-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="133d8-177">Exemplo</span><span class="sxs-lookup"><span data-stu-id="133d8-177">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="133d8-178">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="133d8-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="133d8-179">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="133d8-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="133d8-180">Tipo</span><span class="sxs-lookup"><span data-stu-id="133d8-180">Type</span></span>

*   [<span data-ttu-id="133d8-181">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="133d8-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="133d8-182">Requisitos</span><span class="sxs-lookup"><span data-stu-id="133d8-182">Requirements</span></span>

|<span data-ttu-id="133d8-183">Requisito</span><span class="sxs-lookup"><span data-stu-id="133d8-183">Requirement</span></span>| <span data-ttu-id="133d8-184">Valor</span><span class="sxs-lookup"><span data-stu-id="133d8-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="133d8-185">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="133d8-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="133d8-186">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-186">1.1</span></span>|
|[<span data-ttu-id="133d8-187">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="133d8-187">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="133d8-188">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="133d8-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="133d8-189">Exemplo</span><span class="sxs-lookup"><span data-stu-id="133d8-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="133d8-190">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="133d8-190">displayLanguage: String</span></span>

<span data-ttu-id="133d8-191">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="133d8-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="133d8-192">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="133d8-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="133d8-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="133d8-193">Type</span></span>

*   <span data-ttu-id="133d8-194">String</span><span class="sxs-lookup"><span data-stu-id="133d8-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="133d8-195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="133d8-195">Requirements</span></span>

|<span data-ttu-id="133d8-196">Requisito</span><span class="sxs-lookup"><span data-stu-id="133d8-196">Requirement</span></span>| <span data-ttu-id="133d8-197">Valor</span><span class="sxs-lookup"><span data-stu-id="133d8-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="133d8-198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="133d8-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="133d8-199">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-199">1.1</span></span>|
|[<span data-ttu-id="133d8-200">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="133d8-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="133d8-201">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="133d8-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="133d8-202">Exemplo</span><span class="sxs-lookup"><span data-stu-id="133d8-202">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="133d8-203">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="133d8-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="133d8-204">Obtém o host do aplicativo do Office no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="133d8-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="133d8-205">Tipo</span><span class="sxs-lookup"><span data-stu-id="133d8-205">Type</span></span>

*   [<span data-ttu-id="133d8-206">HostType</span><span class="sxs-lookup"><span data-stu-id="133d8-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="133d8-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="133d8-207">Requirements</span></span>

|<span data-ttu-id="133d8-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="133d8-208">Requirement</span></span>| <span data-ttu-id="133d8-209">Valor</span><span class="sxs-lookup"><span data-stu-id="133d8-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="133d8-210">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="133d8-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="133d8-211">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-211">1.1</span></span>|
|[<span data-ttu-id="133d8-212">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="133d8-212">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="133d8-213">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="133d8-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="133d8-214">Exemplo</span><span class="sxs-lookup"><span data-stu-id="133d8-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="133d8-215">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="133d8-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="133d8-216">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="133d8-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="133d8-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="133d8-217">Type</span></span>

*   [<span data-ttu-id="133d8-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="133d8-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="133d8-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="133d8-219">Requirements</span></span>

|<span data-ttu-id="133d8-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="133d8-220">Requirement</span></span>| <span data-ttu-id="133d8-221">Valor</span><span class="sxs-lookup"><span data-stu-id="133d8-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="133d8-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="133d8-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="133d8-223">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-223">1.1</span></span>|
|[<span data-ttu-id="133d8-224">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="133d8-224">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="133d8-225">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="133d8-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="133d8-226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="133d8-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="133d8-227">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="133d8-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="133d8-228">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o host atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="133d8-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="133d8-229">Tipo</span><span class="sxs-lookup"><span data-stu-id="133d8-229">Type</span></span>

*   [<span data-ttu-id="133d8-230">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="133d8-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="133d8-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="133d8-231">Requirements</span></span>

|<span data-ttu-id="133d8-232">Requisito</span><span class="sxs-lookup"><span data-stu-id="133d8-232">Requirement</span></span>| <span data-ttu-id="133d8-233">Valor</span><span class="sxs-lookup"><span data-stu-id="133d8-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="133d8-234">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="133d8-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="133d8-235">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-235">1.1</span></span>|
|[<span data-ttu-id="133d8-236">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="133d8-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="133d8-237">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="133d8-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="133d8-238">Exemplo</span><span class="sxs-lookup"><span data-stu-id="133d8-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="133d8-239">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="133d8-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="133d8-240">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="133d8-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="133d8-241">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="133d8-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="133d8-242">Tipo</span><span class="sxs-lookup"><span data-stu-id="133d8-242">Type</span></span>

*   [<span data-ttu-id="133d8-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="133d8-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="133d8-244">Requisitos</span><span class="sxs-lookup"><span data-stu-id="133d8-244">Requirements</span></span>

|<span data-ttu-id="133d8-245">Requisito</span><span class="sxs-lookup"><span data-stu-id="133d8-245">Requirement</span></span>| <span data-ttu-id="133d8-246">Valor</span><span class="sxs-lookup"><span data-stu-id="133d8-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="133d8-247">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="133d8-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="133d8-248">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-248">1.1</span></span>|
|[<span data-ttu-id="133d8-249">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="133d8-249">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="133d8-250">Restrito</span><span class="sxs-lookup"><span data-stu-id="133d8-250">Restricted</span></span>|
|[<span data-ttu-id="133d8-251">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="133d8-251">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="133d8-252">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="133d8-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="133d8-253">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="133d8-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="133d8-254">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="133d8-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="133d8-255">Tipo</span><span class="sxs-lookup"><span data-stu-id="133d8-255">Type</span></span>

*   [<span data-ttu-id="133d8-256">UI</span><span class="sxs-lookup"><span data-stu-id="133d8-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="133d8-257">Requisitos</span><span class="sxs-lookup"><span data-stu-id="133d8-257">Requirements</span></span>

|<span data-ttu-id="133d8-258">Requisito</span><span class="sxs-lookup"><span data-stu-id="133d8-258">Requirement</span></span>| <span data-ttu-id="133d8-259">Valor</span><span class="sxs-lookup"><span data-stu-id="133d8-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="133d8-260">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="133d8-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="133d8-261">1.1</span><span class="sxs-lookup"><span data-stu-id="133d8-261">1.1</span></span>|
|[<span data-ttu-id="133d8-262">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="133d8-262">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="133d8-263">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="133d8-263">Compose or Read</span></span>|
