---
title: Office. Context – conjunto de requisitos 1,2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: f5b1e610814e5387c6b27c17569b29860bd2e4be
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814966"
---
# <a name="context"></a><span data-ttu-id="8aa5a-102">context</span><span class="sxs-lookup"><span data-stu-id="8aa5a-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="8aa5a-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="8aa5a-103">[Office](office.md).context</span></span>

<span data-ttu-id="8aa5a-104">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="8aa5a-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="8aa5a-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.2).</span><span class="sxs-lookup"><span data-stu-id="8aa5a-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.2).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8aa5a-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8aa5a-106">Requirements</span></span>

|<span data-ttu-id="8aa5a-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="8aa5a-107">Requirement</span></span>| <span data-ttu-id="8aa5a-108">Valor</span><span class="sxs-lookup"><span data-stu-id="8aa5a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aa5a-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8aa5a-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8aa5a-110">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-110">1.1</span></span>|
|[<span data-ttu-id="8aa5a-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8aa5a-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aa5a-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8aa5a-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8aa5a-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="8aa5a-113">Properties</span></span>

| <span data-ttu-id="8aa5a-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="8aa5a-114">Property</span></span> | <span data-ttu-id="8aa5a-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="8aa5a-115">Modes</span></span> | <span data-ttu-id="8aa5a-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="8aa5a-116">Return type</span></span> | <span data-ttu-id="8aa5a-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-117">Minimum</span></span><br><span data-ttu-id="8aa5a-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="8aa5a-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8aa5a-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="8aa5a-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="8aa5a-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="8aa5a-120">Compose</span></span><br><span data-ttu-id="8aa5a-121">Leitura</span><span class="sxs-lookup"><span data-stu-id="8aa5a-121">Read</span></span> | <span data-ttu-id="8aa5a-122">String</span><span class="sxs-lookup"><span data-stu-id="8aa5a-122">String</span></span> | [<span data-ttu-id="8aa5a-123">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8aa5a-124">la</span><span class="sxs-lookup"><span data-stu-id="8aa5a-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="8aa5a-125">Escrever</span><span class="sxs-lookup"><span data-stu-id="8aa5a-125">Compose</span></span><br><span data-ttu-id="8aa5a-126">Leitura</span><span class="sxs-lookup"><span data-stu-id="8aa5a-126">Read</span></span> | [<span data-ttu-id="8aa5a-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="8aa5a-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.2) | [<span data-ttu-id="8aa5a-128">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8aa5a-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="8aa5a-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="8aa5a-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="8aa5a-130">Compose</span></span><br><span data-ttu-id="8aa5a-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="8aa5a-131">Read</span></span> | <span data-ttu-id="8aa5a-132">String</span><span class="sxs-lookup"><span data-stu-id="8aa5a-132">String</span></span> | [<span data-ttu-id="8aa5a-133">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8aa5a-134">principal</span><span class="sxs-lookup"><span data-stu-id="8aa5a-134">host</span></span>](#host-hosttype) | <span data-ttu-id="8aa5a-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="8aa5a-135">Compose</span></span><br><span data-ttu-id="8aa5a-136">Leitura</span><span class="sxs-lookup"><span data-stu-id="8aa5a-136">Read</span></span> | [<span data-ttu-id="8aa5a-137">HostType</span><span class="sxs-lookup"><span data-stu-id="8aa5a-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.2) | [<span data-ttu-id="8aa5a-138">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8aa5a-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="8aa5a-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="8aa5a-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="8aa5a-140">Compose</span></span><br><span data-ttu-id="8aa5a-141">Leitura</span><span class="sxs-lookup"><span data-stu-id="8aa5a-141">Read</span></span> | [<span data-ttu-id="8aa5a-142">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="8aa5a-142">Mailbox</span></span>](/javascript/api/office/office.mailbox?view=outlook-js-1.2) | [<span data-ttu-id="8aa5a-143">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8aa5a-144">plataforma</span><span class="sxs-lookup"><span data-stu-id="8aa5a-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="8aa5a-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="8aa5a-145">Compose</span></span><br><span data-ttu-id="8aa5a-146">Leitura</span><span class="sxs-lookup"><span data-stu-id="8aa5a-146">Read</span></span> | [<span data-ttu-id="8aa5a-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="8aa5a-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.2) | [<span data-ttu-id="8aa5a-148">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8aa5a-149">atende</span><span class="sxs-lookup"><span data-stu-id="8aa5a-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="8aa5a-150">Escrever</span><span class="sxs-lookup"><span data-stu-id="8aa5a-150">Compose</span></span><br><span data-ttu-id="8aa5a-151">Leitura</span><span class="sxs-lookup"><span data-stu-id="8aa5a-151">Read</span></span> | [<span data-ttu-id="8aa5a-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="8aa5a-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.2) | [<span data-ttu-id="8aa5a-153">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8aa5a-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="8aa5a-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="8aa5a-155">Escrever</span><span class="sxs-lookup"><span data-stu-id="8aa5a-155">Compose</span></span><br><span data-ttu-id="8aa5a-156">Leitura</span><span class="sxs-lookup"><span data-stu-id="8aa5a-156">Read</span></span> | [<span data-ttu-id="8aa5a-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8aa5a-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.2) | [<span data-ttu-id="8aa5a-158">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8aa5a-159">ui</span><span class="sxs-lookup"><span data-stu-id="8aa5a-159">ui</span></span>](#ui-ui) | <span data-ttu-id="8aa5a-160">Escrever</span><span class="sxs-lookup"><span data-stu-id="8aa5a-160">Compose</span></span><br><span data-ttu-id="8aa5a-161">Leitura</span><span class="sxs-lookup"><span data-stu-id="8aa5a-161">Read</span></span> | [<span data-ttu-id="8aa5a-162">UI</span><span class="sxs-lookup"><span data-stu-id="8aa5a-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.2) | [<span data-ttu-id="8aa5a-163">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="8aa5a-164">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="8aa5a-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="8aa5a-165">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8aa5a-165">contentLanguage: String</span></span>

<span data-ttu-id="8aa5a-166">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="8aa5a-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="8aa5a-167">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="8aa5a-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="8aa5a-168">Tipo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-168">Type</span></span>

*   <span data-ttu-id="8aa5a-169">String</span><span class="sxs-lookup"><span data-stu-id="8aa5a-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8aa5a-170">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8aa5a-170">Requirements</span></span>

|<span data-ttu-id="8aa5a-171">Requisito</span><span class="sxs-lookup"><span data-stu-id="8aa5a-171">Requirement</span></span>| <span data-ttu-id="8aa5a-172">Valor</span><span class="sxs-lookup"><span data-stu-id="8aa5a-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aa5a-173">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8aa5a-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8aa5a-174">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-174">1.1</span></span>|
|[<span data-ttu-id="8aa5a-175">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8aa5a-175">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aa5a-176">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8aa5a-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8aa5a-177">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-177">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="8aa5a-178">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="8aa5a-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="8aa5a-179">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="8aa5a-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="8aa5a-180">Tipo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-180">Type</span></span>

*   [<span data-ttu-id="8aa5a-181">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="8aa5a-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="8aa5a-182">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8aa5a-182">Requirements</span></span>

|<span data-ttu-id="8aa5a-183">Requisito</span><span class="sxs-lookup"><span data-stu-id="8aa5a-183">Requirement</span></span>| <span data-ttu-id="8aa5a-184">Valor</span><span class="sxs-lookup"><span data-stu-id="8aa5a-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aa5a-185">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8aa5a-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8aa5a-186">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-186">1.1</span></span>|
|[<span data-ttu-id="8aa5a-187">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8aa5a-187">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aa5a-188">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8aa5a-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8aa5a-189">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="8aa5a-190">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8aa5a-190">displayLanguage: String</span></span>

<span data-ttu-id="8aa5a-191">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="8aa5a-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="8aa5a-192">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="8aa5a-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="8aa5a-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-193">Type</span></span>

*   <span data-ttu-id="8aa5a-194">String</span><span class="sxs-lookup"><span data-stu-id="8aa5a-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8aa5a-195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8aa5a-195">Requirements</span></span>

|<span data-ttu-id="8aa5a-196">Requisito</span><span class="sxs-lookup"><span data-stu-id="8aa5a-196">Requirement</span></span>| <span data-ttu-id="8aa5a-197">Valor</span><span class="sxs-lookup"><span data-stu-id="8aa5a-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aa5a-198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8aa5a-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8aa5a-199">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-199">1.1</span></span>|
|[<span data-ttu-id="8aa5a-200">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8aa5a-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aa5a-201">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8aa5a-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8aa5a-202">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-202">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="8aa5a-203">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="8aa5a-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="8aa5a-204">Obtém o host do aplicativo do Office no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="8aa5a-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="8aa5a-205">Tipo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-205">Type</span></span>

*   [<span data-ttu-id="8aa5a-206">HostType</span><span class="sxs-lookup"><span data-stu-id="8aa5a-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="8aa5a-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8aa5a-207">Requirements</span></span>

|<span data-ttu-id="8aa5a-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="8aa5a-208">Requirement</span></span>| <span data-ttu-id="8aa5a-209">Valor</span><span class="sxs-lookup"><span data-stu-id="8aa5a-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aa5a-210">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8aa5a-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8aa5a-211">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-211">1.1</span></span>|
|[<span data-ttu-id="8aa5a-212">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8aa5a-212">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aa5a-213">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8aa5a-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8aa5a-214">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="8aa5a-215">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="8aa5a-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="8aa5a-216">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="8aa5a-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="8aa5a-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-217">Type</span></span>

*   [<span data-ttu-id="8aa5a-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="8aa5a-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="8aa5a-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8aa5a-219">Requirements</span></span>

|<span data-ttu-id="8aa5a-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="8aa5a-220">Requirement</span></span>| <span data-ttu-id="8aa5a-221">Valor</span><span class="sxs-lookup"><span data-stu-id="8aa5a-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aa5a-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8aa5a-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8aa5a-223">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-223">1.1</span></span>|
|[<span data-ttu-id="8aa5a-224">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8aa5a-224">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aa5a-225">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8aa5a-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8aa5a-226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="8aa5a-227">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="8aa5a-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="8aa5a-228">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o host atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="8aa5a-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="8aa5a-229">Tipo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-229">Type</span></span>

*   [<span data-ttu-id="8aa5a-230">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="8aa5a-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="8aa5a-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8aa5a-231">Requirements</span></span>

|<span data-ttu-id="8aa5a-232">Requisito</span><span class="sxs-lookup"><span data-stu-id="8aa5a-232">Requirement</span></span>| <span data-ttu-id="8aa5a-233">Valor</span><span class="sxs-lookup"><span data-stu-id="8aa5a-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aa5a-234">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8aa5a-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8aa5a-235">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-235">1.1</span></span>|
|[<span data-ttu-id="8aa5a-236">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8aa5a-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aa5a-237">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8aa5a-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8aa5a-238">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="8aa5a-239">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="8aa5a-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="8aa5a-240">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="8aa5a-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="8aa5a-241">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="8aa5a-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="8aa5a-242">Tipo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-242">Type</span></span>

*   [<span data-ttu-id="8aa5a-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8aa5a-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="8aa5a-244">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8aa5a-244">Requirements</span></span>

|<span data-ttu-id="8aa5a-245">Requisito</span><span class="sxs-lookup"><span data-stu-id="8aa5a-245">Requirement</span></span>| <span data-ttu-id="8aa5a-246">Valor</span><span class="sxs-lookup"><span data-stu-id="8aa5a-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aa5a-247">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8aa5a-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8aa5a-248">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-248">1.1</span></span>|
|[<span data-ttu-id="8aa5a-249">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-249">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8aa5a-250">Restrito</span><span class="sxs-lookup"><span data-stu-id="8aa5a-250">Restricted</span></span>|
|[<span data-ttu-id="8aa5a-251">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8aa5a-251">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aa5a-252">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8aa5a-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="8aa5a-253">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="8aa5a-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="8aa5a-254">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="8aa5a-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="8aa5a-255">Tipo</span><span class="sxs-lookup"><span data-stu-id="8aa5a-255">Type</span></span>

*   [<span data-ttu-id="8aa5a-256">UI</span><span class="sxs-lookup"><span data-stu-id="8aa5a-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="8aa5a-257">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8aa5a-257">Requirements</span></span>

|<span data-ttu-id="8aa5a-258">Requisito</span><span class="sxs-lookup"><span data-stu-id="8aa5a-258">Requirement</span></span>| <span data-ttu-id="8aa5a-259">Valor</span><span class="sxs-lookup"><span data-stu-id="8aa5a-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aa5a-260">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8aa5a-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8aa5a-261">1.1</span><span class="sxs-lookup"><span data-stu-id="8aa5a-261">1.1</span></span>|
|[<span data-ttu-id="8aa5a-262">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8aa5a-262">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aa5a-263">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8aa5a-263">Compose or Read</span></span>|
