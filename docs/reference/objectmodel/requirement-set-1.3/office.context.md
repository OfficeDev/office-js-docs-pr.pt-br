---
title: Office.context - conjunto de requisitos 1.3
description: Office. Membros do objeto context disponível para Outlook os complementos usando o conjunto de requisitos da API de Caixa de Correio 1.3.
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: c5c7f6eaa46bfa5067572878b03e47511310f894
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590390"
---
# <a name="context-mailbox-requirement-set-13"></a><span data-ttu-id="197f7-103">context (Conjunto de requisitos de caixa de correio 1.3)</span><span class="sxs-lookup"><span data-stu-id="197f7-103">context (Mailbox requirement set 1.3)</span></span>

### <a name="officecontext"></a><span data-ttu-id="197f7-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="197f7-104">[Office](office.md).context</span></span>

<span data-ttu-id="197f7-105">Office.context fornece interfaces compartilhadas que são usadas por complementos em todos os Office aplicativos.</span><span class="sxs-lookup"><span data-stu-id="197f7-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="197f7-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.3&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="197f7-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="197f7-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="197f7-107">Requirements</span></span>

|<span data-ttu-id="197f7-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="197f7-108">Requirement</span></span>| <span data-ttu-id="197f7-109">Valor</span><span class="sxs-lookup"><span data-stu-id="197f7-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="197f7-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="197f7-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="197f7-111">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-111">1.1</span></span>|
|[<span data-ttu-id="197f7-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="197f7-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="197f7-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="197f7-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="197f7-114">Properties</span></span>

| <span data-ttu-id="197f7-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="197f7-115">Property</span></span> | <span data-ttu-id="197f7-116">Modos</span><span class="sxs-lookup"><span data-stu-id="197f7-116">Modes</span></span> | <span data-ttu-id="197f7-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="197f7-117">Return type</span></span> | <span data-ttu-id="197f7-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="197f7-118">Minimum</span></span><br><span data-ttu-id="197f7-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="197f7-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="197f7-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="197f7-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="197f7-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="197f7-121">Compose</span></span><br><span data-ttu-id="197f7-122">Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-122">Read</span></span> | <span data-ttu-id="197f7-123">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="197f7-123">String</span></span> | [<span data-ttu-id="197f7-124">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="197f7-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="197f7-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="197f7-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="197f7-126">Compose</span></span><br><span data-ttu-id="197f7-127">Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-127">Read</span></span> | [<span data-ttu-id="197f7-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="197f7-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="197f7-129">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="197f7-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="197f7-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="197f7-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="197f7-131">Compose</span></span><br><span data-ttu-id="197f7-132">Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-132">Read</span></span> | <span data-ttu-id="197f7-133">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="197f7-133">String</span></span> | [<span data-ttu-id="197f7-134">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="197f7-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="197f7-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="197f7-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="197f7-136">Compose</span></span><br><span data-ttu-id="197f7-137">Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-137">Read</span></span> | [<span data-ttu-id="197f7-138">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="197f7-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="197f7-139">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="197f7-140">requirements</span><span class="sxs-lookup"><span data-stu-id="197f7-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="197f7-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="197f7-141">Compose</span></span><br><span data-ttu-id="197f7-142">Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-142">Read</span></span> | [<span data-ttu-id="197f7-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="197f7-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="197f7-144">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="197f7-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="197f7-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="197f7-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="197f7-146">Compose</span></span><br><span data-ttu-id="197f7-147">Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-147">Read</span></span> | [<span data-ttu-id="197f7-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="197f7-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="197f7-149">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="197f7-150">ui</span><span class="sxs-lookup"><span data-stu-id="197f7-150">ui</span></span>](#ui-ui) | <span data-ttu-id="197f7-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="197f7-151">Compose</span></span><br><span data-ttu-id="197f7-152">Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-152">Read</span></span> | [<span data-ttu-id="197f7-153">UI</span><span class="sxs-lookup"><span data-stu-id="197f7-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="197f7-154">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="197f7-155">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="197f7-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="197f7-156">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="197f7-156">contentLanguage: String</span></span>

<span data-ttu-id="197f7-157">Obtém a localidade (idioma) especificada pelo usuário para editar o item.</span><span class="sxs-lookup"><span data-stu-id="197f7-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="197f7-158">O `contentLanguage` valor reflete a **configuração** atual de Idioma de Edição especificada com opções de > de arquivo **> idioma** no aplicativo Office cliente.</span><span class="sxs-lookup"><span data-stu-id="197f7-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="197f7-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="197f7-159">Type</span></span>

*   <span data-ttu-id="197f7-160">String</span><span class="sxs-lookup"><span data-stu-id="197f7-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="197f7-161">Requisitos</span><span class="sxs-lookup"><span data-stu-id="197f7-161">Requirements</span></span>

|<span data-ttu-id="197f7-162">Requisito</span><span class="sxs-lookup"><span data-stu-id="197f7-162">Requirement</span></span>| <span data-ttu-id="197f7-163">Valor</span><span class="sxs-lookup"><span data-stu-id="197f7-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="197f7-164">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="197f7-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="197f7-165">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-165">1.1</span></span>|
|[<span data-ttu-id="197f7-166">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="197f7-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="197f7-167">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="197f7-168">Exemplo</span><span class="sxs-lookup"><span data-stu-id="197f7-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="197f7-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="197f7-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="197f7-170">Obtém informações sobre o ambiente no qual o complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="197f7-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="197f7-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="197f7-171">Type</span></span>

*   [<span data-ttu-id="197f7-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="197f7-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="197f7-173">Requisitos</span><span class="sxs-lookup"><span data-stu-id="197f7-173">Requirements</span></span>

|<span data-ttu-id="197f7-174">Requisito</span><span class="sxs-lookup"><span data-stu-id="197f7-174">Requirement</span></span>| <span data-ttu-id="197f7-175">Valor</span><span class="sxs-lookup"><span data-stu-id="197f7-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="197f7-176">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="197f7-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="197f7-177">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-177">1.1</span></span>|
|[<span data-ttu-id="197f7-178">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="197f7-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="197f7-179">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="197f7-180">Exemplo</span><span class="sxs-lookup"><span data-stu-id="197f7-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="197f7-181">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="197f7-181">displayLanguage: String</span></span>

<span data-ttu-id="197f7-182">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente Office cliente.</span><span class="sxs-lookup"><span data-stu-id="197f7-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="197f7-183">O valor reflete a configuração atual de Idioma de Exibição especificada com Opções > > Idioma no aplicativo Office `displayLanguage` cliente.  </span><span class="sxs-lookup"><span data-stu-id="197f7-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="197f7-184">Tipo</span><span class="sxs-lookup"><span data-stu-id="197f7-184">Type</span></span>

*   <span data-ttu-id="197f7-185">String</span><span class="sxs-lookup"><span data-stu-id="197f7-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="197f7-186">Requisitos</span><span class="sxs-lookup"><span data-stu-id="197f7-186">Requirements</span></span>

|<span data-ttu-id="197f7-187">Requisito</span><span class="sxs-lookup"><span data-stu-id="197f7-187">Requirement</span></span>| <span data-ttu-id="197f7-188">Valor</span><span class="sxs-lookup"><span data-stu-id="197f7-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="197f7-189">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="197f7-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="197f7-190">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-190">1.1</span></span>|
|[<span data-ttu-id="197f7-191">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="197f7-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="197f7-192">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="197f7-193">Exemplo</span><span class="sxs-lookup"><span data-stu-id="197f7-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="197f7-194">requirements: [RequirementsSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="197f7-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="197f7-195">Fornece um método para determinar quais conjuntos de requisitos são suportados no aplicativo e na plataforma atual.</span><span class="sxs-lookup"><span data-stu-id="197f7-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="197f7-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="197f7-196">Type</span></span>

*   [<span data-ttu-id="197f7-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="197f7-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="197f7-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="197f7-198">Requirements</span></span>

|<span data-ttu-id="197f7-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="197f7-199">Requirement</span></span>| <span data-ttu-id="197f7-200">Valor</span><span class="sxs-lookup"><span data-stu-id="197f7-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="197f7-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="197f7-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="197f7-202">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-202">1.1</span></span>|
|[<span data-ttu-id="197f7-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="197f7-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="197f7-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="197f7-205">Exemplo</span><span class="sxs-lookup"><span data-stu-id="197f7-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="197f7-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="197f7-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="197f7-207">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="197f7-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="197f7-208">O objeto permite que você armazene e acesse dados para um complemento de email armazenado na caixa de correio de um usuário, de modo que está disponível para esse complemento quando ele está sendo executado de qualquer cliente Outlook usado para acessar essa caixa de `RoamingSettings` correio.</span><span class="sxs-lookup"><span data-stu-id="197f7-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="197f7-209">Tipo</span><span class="sxs-lookup"><span data-stu-id="197f7-209">Type</span></span>

*   [<span data-ttu-id="197f7-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="197f7-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="197f7-211">Requisitos</span><span class="sxs-lookup"><span data-stu-id="197f7-211">Requirements</span></span>

|<span data-ttu-id="197f7-212">Requisito</span><span class="sxs-lookup"><span data-stu-id="197f7-212">Requirement</span></span>| <span data-ttu-id="197f7-213">Valor</span><span class="sxs-lookup"><span data-stu-id="197f7-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="197f7-214">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="197f7-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="197f7-215">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-215">1.1</span></span>|
|[<span data-ttu-id="197f7-216">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="197f7-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="197f7-217">Restrito</span><span class="sxs-lookup"><span data-stu-id="197f7-217">Restricted</span></span>|
|[<span data-ttu-id="197f7-218">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="197f7-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="197f7-219">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="197f7-220">interface do usuário: [interface do usuário](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="197f7-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="197f7-221">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus Office de usuário.</span><span class="sxs-lookup"><span data-stu-id="197f7-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="197f7-222">Tipo</span><span class="sxs-lookup"><span data-stu-id="197f7-222">Type</span></span>

*   [<span data-ttu-id="197f7-223">UI</span><span class="sxs-lookup"><span data-stu-id="197f7-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="197f7-224">Requisitos</span><span class="sxs-lookup"><span data-stu-id="197f7-224">Requirements</span></span>

|<span data-ttu-id="197f7-225">Requisito</span><span class="sxs-lookup"><span data-stu-id="197f7-225">Requirement</span></span>| <span data-ttu-id="197f7-226">Valor</span><span class="sxs-lookup"><span data-stu-id="197f7-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="197f7-227">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="197f7-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="197f7-228">1.1</span><span class="sxs-lookup"><span data-stu-id="197f7-228">1.1</span></span>|
|[<span data-ttu-id="197f7-229">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="197f7-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="197f7-230">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="197f7-230">Compose or Read</span></span>|
