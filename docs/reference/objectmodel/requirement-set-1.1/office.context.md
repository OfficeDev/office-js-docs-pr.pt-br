---
title: Office.context - conjunto de requisitos 1.1
description: Office. Membros do objeto Context disponíveis para Outlook de entrada usando o conjunto de requisitos da API de Caixa de Correio 1.1.
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 41273bfc5362a9d5572e38b8e80b81041f5aa312
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590873"
---
# <a name="context-mailbox-requirement-set-11"></a><span data-ttu-id="8fcac-103">context (Conjunto de requisitos de caixa de correio 1.1)</span><span class="sxs-lookup"><span data-stu-id="8fcac-103">context (Mailbox requirement set 1.1)</span></span>

### <a name="officecontext"></a><span data-ttu-id="8fcac-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="8fcac-104">[Office](office.md).context</span></span>

<span data-ttu-id="8fcac-105">Office.context fornece interfaces compartilhadas que são usadas por complementos em todos os Office aplicativos.</span><span class="sxs-lookup"><span data-stu-id="8fcac-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="8fcac-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="8fcac-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fcac-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fcac-107">Requirements</span></span>

|<span data-ttu-id="8fcac-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fcac-108">Requirement</span></span>| <span data-ttu-id="8fcac-109">Valor</span><span class="sxs-lookup"><span data-stu-id="8fcac-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fcac-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fcac-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8fcac-111">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-111">1.1</span></span>|
|[<span data-ttu-id="8fcac-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fcac-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8fcac-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="8fcac-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="8fcac-114">Properties</span></span>

| <span data-ttu-id="8fcac-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="8fcac-115">Property</span></span> | <span data-ttu-id="8fcac-116">Modos</span><span class="sxs-lookup"><span data-stu-id="8fcac-116">Modes</span></span> | <span data-ttu-id="8fcac-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="8fcac-117">Return type</span></span> | <span data-ttu-id="8fcac-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="8fcac-118">Minimum</span></span><br><span data-ttu-id="8fcac-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="8fcac-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8fcac-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="8fcac-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="8fcac-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fcac-121">Compose</span></span><br><span data-ttu-id="8fcac-122">Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-122">Read</span></span> | <span data-ttu-id="8fcac-123">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8fcac-123">String</span></span> | [<span data-ttu-id="8fcac-124">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8fcac-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="8fcac-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="8fcac-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fcac-126">Compose</span></span><br><span data-ttu-id="8fcac-127">Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-127">Read</span></span> | [<span data-ttu-id="8fcac-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="8fcac-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="8fcac-129">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8fcac-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="8fcac-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="8fcac-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fcac-131">Compose</span></span><br><span data-ttu-id="8fcac-132">Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-132">Read</span></span> | <span data-ttu-id="8fcac-133">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8fcac-133">String</span></span> | [<span data-ttu-id="8fcac-134">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8fcac-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="8fcac-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="8fcac-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fcac-136">Compose</span></span><br><span data-ttu-id="8fcac-137">Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-137">Read</span></span> | [<span data-ttu-id="8fcac-138">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="8fcac-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="8fcac-139">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8fcac-140">requirements</span><span class="sxs-lookup"><span data-stu-id="8fcac-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="8fcac-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fcac-141">Compose</span></span><br><span data-ttu-id="8fcac-142">Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-142">Read</span></span> | [<span data-ttu-id="8fcac-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="8fcac-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="8fcac-144">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8fcac-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="8fcac-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="8fcac-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fcac-146">Compose</span></span><br><span data-ttu-id="8fcac-147">Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-147">Read</span></span> | [<span data-ttu-id="8fcac-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8fcac-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="8fcac-149">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8fcac-150">ui</span><span class="sxs-lookup"><span data-stu-id="8fcac-150">ui</span></span>](#ui-ui) | <span data-ttu-id="8fcac-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="8fcac-151">Compose</span></span><br><span data-ttu-id="8fcac-152">Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-152">Read</span></span> | [<span data-ttu-id="8fcac-153">UI</span><span class="sxs-lookup"><span data-stu-id="8fcac-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="8fcac-154">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="8fcac-155">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="8fcac-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="8fcac-156">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="8fcac-156">contentLanguage: String</span></span>

<span data-ttu-id="8fcac-157">Obtém a localidade (idioma) especificada pelo usuário para editar o item.</span><span class="sxs-lookup"><span data-stu-id="8fcac-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="8fcac-158">O `contentLanguage` valor reflete a **configuração** atual de Idioma de Edição especificada com opções de > de arquivo **> idioma** no aplicativo Office cliente.</span><span class="sxs-lookup"><span data-stu-id="8fcac-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="8fcac-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fcac-159">Type</span></span>

*   <span data-ttu-id="8fcac-160">String</span><span class="sxs-lookup"><span data-stu-id="8fcac-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fcac-161">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fcac-161">Requirements</span></span>

|<span data-ttu-id="8fcac-162">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fcac-162">Requirement</span></span>| <span data-ttu-id="8fcac-163">Valor</span><span class="sxs-lookup"><span data-stu-id="8fcac-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fcac-164">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fcac-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8fcac-165">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-165">1.1</span></span>|
|[<span data-ttu-id="8fcac-166">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fcac-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8fcac-167">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fcac-168">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fcac-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="8fcac-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="8fcac-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="8fcac-170">Obtém informações sobre o ambiente no qual o complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="8fcac-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="8fcac-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fcac-171">Type</span></span>

*   [<span data-ttu-id="8fcac-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="8fcac-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="8fcac-173">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fcac-173">Requirements</span></span>

|<span data-ttu-id="8fcac-174">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fcac-174">Requirement</span></span>| <span data-ttu-id="8fcac-175">Valor</span><span class="sxs-lookup"><span data-stu-id="8fcac-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fcac-176">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fcac-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8fcac-177">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-177">1.1</span></span>|
|[<span data-ttu-id="8fcac-178">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fcac-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8fcac-179">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fcac-180">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fcac-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="8fcac-181">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="8fcac-181">displayLanguage: String</span></span>

<span data-ttu-id="8fcac-182">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente Office cliente.</span><span class="sxs-lookup"><span data-stu-id="8fcac-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="8fcac-183">O valor reflete a configuração atual de Idioma de Exibição especificada com Opções > > Idioma no aplicativo Office `displayLanguage` cliente.  </span><span class="sxs-lookup"><span data-stu-id="8fcac-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="8fcac-184">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fcac-184">Type</span></span>

*   <span data-ttu-id="8fcac-185">String</span><span class="sxs-lookup"><span data-stu-id="8fcac-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fcac-186">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fcac-186">Requirements</span></span>

|<span data-ttu-id="8fcac-187">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fcac-187">Requirement</span></span>| <span data-ttu-id="8fcac-188">Valor</span><span class="sxs-lookup"><span data-stu-id="8fcac-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fcac-189">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fcac-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8fcac-190">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-190">1.1</span></span>|
|[<span data-ttu-id="8fcac-191">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fcac-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8fcac-192">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fcac-193">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fcac-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="8fcac-194">requirements: [RequirementsSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="8fcac-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="8fcac-195">Fornece um método para determinar quais conjuntos de requisitos são suportados no aplicativo e na plataforma atual.</span><span class="sxs-lookup"><span data-stu-id="8fcac-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="8fcac-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fcac-196">Type</span></span>

*   [<span data-ttu-id="8fcac-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="8fcac-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="8fcac-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fcac-198">Requirements</span></span>

|<span data-ttu-id="8fcac-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fcac-199">Requirement</span></span>| <span data-ttu-id="8fcac-200">Valor</span><span class="sxs-lookup"><span data-stu-id="8fcac-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fcac-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fcac-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8fcac-202">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-202">1.1</span></span>|
|[<span data-ttu-id="8fcac-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fcac-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8fcac-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8fcac-205">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8fcac-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="8fcac-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="8fcac-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="8fcac-207">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="8fcac-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="8fcac-208">O objeto permite que você armazene e acesse dados para um complemento de email armazenado na caixa de correio de um usuário, de modo que está disponível para esse complemento quando ele está sendo executado de qualquer cliente Outlook usado para acessar essa caixa de `RoamingSettings` correio.</span><span class="sxs-lookup"><span data-stu-id="8fcac-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="8fcac-209">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fcac-209">Type</span></span>

*   [<span data-ttu-id="8fcac-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8fcac-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="8fcac-211">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fcac-211">Requirements</span></span>

|<span data-ttu-id="8fcac-212">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fcac-212">Requirement</span></span>| <span data-ttu-id="8fcac-213">Valor</span><span class="sxs-lookup"><span data-stu-id="8fcac-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fcac-214">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fcac-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8fcac-215">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-215">1.1</span></span>|
|[<span data-ttu-id="8fcac-216">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8fcac-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="8fcac-217">Restrito</span><span class="sxs-lookup"><span data-stu-id="8fcac-217">Restricted</span></span>|
|[<span data-ttu-id="8fcac-218">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fcac-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8fcac-219">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="8fcac-220">interface do usuário: [interface do usuário](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="8fcac-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="8fcac-221">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus Office de usuário.</span><span class="sxs-lookup"><span data-stu-id="8fcac-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="8fcac-222">Tipo</span><span class="sxs-lookup"><span data-stu-id="8fcac-222">Type</span></span>

*   [<span data-ttu-id="8fcac-223">UI</span><span class="sxs-lookup"><span data-stu-id="8fcac-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="8fcac-224">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8fcac-224">Requirements</span></span>

|<span data-ttu-id="8fcac-225">Requisito</span><span class="sxs-lookup"><span data-stu-id="8fcac-225">Requirement</span></span>| <span data-ttu-id="8fcac-226">Valor</span><span class="sxs-lookup"><span data-stu-id="8fcac-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fcac-227">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8fcac-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8fcac-228">1.1</span><span class="sxs-lookup"><span data-stu-id="8fcac-228">1.1</span></span>|
|[<span data-ttu-id="8fcac-229">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8fcac-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8fcac-230">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="8fcac-230">Compose or Read</span></span>|
