---
title: Office. Context – conjunto de requisitos 1,4
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,4.
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 0ec84c9d0695871fa3be265c37ce1e682cdfb6af
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570769"
---
# <a name="context-mailbox-requirement-set-14"></a><span data-ttu-id="3fa7f-103">contexto (conjunto de requisitos de caixa de correio 1,4)</span><span class="sxs-lookup"><span data-stu-id="3fa7f-103">context (Mailbox requirement set 1.4)</span></span>

### <a name="officecontext"></a><span data-ttu-id="3fa7f-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="3fa7f-104">[Office](office.md).context</span></span>

<span data-ttu-id="3fa7f-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="3fa7f-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="3fa7f-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="3fa7f-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3fa7f-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3fa7f-107">Requirements</span></span>

|<span data-ttu-id="3fa7f-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="3fa7f-108">Requirement</span></span>| <span data-ttu-id="3fa7f-109">Valor</span><span class="sxs-lookup"><span data-stu-id="3fa7f-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fa7f-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3fa7f-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fa7f-111">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-111">1.1</span></span>|
|[<span data-ttu-id="3fa7f-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3fa7f-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fa7f-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="3fa7f-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="3fa7f-114">Properties</span></span>

| <span data-ttu-id="3fa7f-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="3fa7f-115">Property</span></span> | <span data-ttu-id="3fa7f-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="3fa7f-116">Modes</span></span> | <span data-ttu-id="3fa7f-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="3fa7f-117">Return type</span></span> | <span data-ttu-id="3fa7f-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="3fa7f-118">Minimum</span></span><br><span data-ttu-id="3fa7f-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="3fa7f-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3fa7f-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="3fa7f-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="3fa7f-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="3fa7f-121">Compose</span></span><br><span data-ttu-id="3fa7f-122">Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-122">Read</span></span> | <span data-ttu-id="3fa7f-123">String</span><span class="sxs-lookup"><span data-stu-id="3fa7f-123">String</span></span> | [<span data-ttu-id="3fa7f-124">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fa7f-125">la</span><span class="sxs-lookup"><span data-stu-id="3fa7f-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="3fa7f-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="3fa7f-126">Compose</span></span><br><span data-ttu-id="3fa7f-127">Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-127">Read</span></span> | [<span data-ttu-id="3fa7f-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="3fa7f-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="3fa7f-129">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fa7f-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="3fa7f-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="3fa7f-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="3fa7f-131">Compose</span></span><br><span data-ttu-id="3fa7f-132">Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-132">Read</span></span> | <span data-ttu-id="3fa7f-133">String</span><span class="sxs-lookup"><span data-stu-id="3fa7f-133">String</span></span> | [<span data-ttu-id="3fa7f-134">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fa7f-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="3fa7f-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="3fa7f-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="3fa7f-136">Compose</span></span><br><span data-ttu-id="3fa7f-137">Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-137">Read</span></span> | [<span data-ttu-id="3fa7f-138">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="3fa7f-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="3fa7f-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fa7f-140">atende</span><span class="sxs-lookup"><span data-stu-id="3fa7f-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="3fa7f-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="3fa7f-141">Compose</span></span><br><span data-ttu-id="3fa7f-142">Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-142">Read</span></span> | [<span data-ttu-id="3fa7f-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="3fa7f-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="3fa7f-144">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fa7f-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="3fa7f-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="3fa7f-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="3fa7f-146">Compose</span></span><br><span data-ttu-id="3fa7f-147">Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-147">Read</span></span> | [<span data-ttu-id="3fa7f-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="3fa7f-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="3fa7f-149">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fa7f-150">ui</span><span class="sxs-lookup"><span data-stu-id="3fa7f-150">ui</span></span>](#ui-ui) | <span data-ttu-id="3fa7f-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="3fa7f-151">Compose</span></span><br><span data-ttu-id="3fa7f-152">Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-152">Read</span></span> | [<span data-ttu-id="3fa7f-153">UI</span><span class="sxs-lookup"><span data-stu-id="3fa7f-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="3fa7f-154">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="3fa7f-155">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="3fa7f-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="3fa7f-156">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3fa7f-156">contentLanguage: String</span></span>

<span data-ttu-id="3fa7f-157">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="3fa7f-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="3fa7f-158">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="3fa7f-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="3fa7f-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="3fa7f-159">Type</span></span>

*   <span data-ttu-id="3fa7f-160">String</span><span class="sxs-lookup"><span data-stu-id="3fa7f-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3fa7f-161">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3fa7f-161">Requirements</span></span>

|<span data-ttu-id="3fa7f-162">Requisito</span><span class="sxs-lookup"><span data-stu-id="3fa7f-162">Requirement</span></span>| <span data-ttu-id="3fa7f-163">Valor</span><span class="sxs-lookup"><span data-stu-id="3fa7f-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fa7f-164">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3fa7f-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fa7f-165">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-165">1.1</span></span>|
|[<span data-ttu-id="3fa7f-166">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3fa7f-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fa7f-167">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3fa7f-168">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3fa7f-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="3fa7f-169">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="3fa7f-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="3fa7f-170">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="3fa7f-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="3fa7f-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="3fa7f-171">Type</span></span>

*   [<span data-ttu-id="3fa7f-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="3fa7f-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="3fa7f-173">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3fa7f-173">Requirements</span></span>

|<span data-ttu-id="3fa7f-174">Requisito</span><span class="sxs-lookup"><span data-stu-id="3fa7f-174">Requirement</span></span>| <span data-ttu-id="3fa7f-175">Valor</span><span class="sxs-lookup"><span data-stu-id="3fa7f-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fa7f-176">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3fa7f-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fa7f-177">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-177">1.1</span></span>|
|[<span data-ttu-id="3fa7f-178">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3fa7f-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fa7f-179">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3fa7f-180">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3fa7f-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="3fa7f-181">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3fa7f-181">displayLanguage: String</span></span>

<span data-ttu-id="3fa7f-182">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="3fa7f-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="3fa7f-183">O `displayLanguage` valor reflete a configuração de **idioma de exibição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="3fa7f-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="3fa7f-184">Tipo</span><span class="sxs-lookup"><span data-stu-id="3fa7f-184">Type</span></span>

*   <span data-ttu-id="3fa7f-185">String</span><span class="sxs-lookup"><span data-stu-id="3fa7f-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3fa7f-186">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3fa7f-186">Requirements</span></span>

|<span data-ttu-id="3fa7f-187">Requisito</span><span class="sxs-lookup"><span data-stu-id="3fa7f-187">Requirement</span></span>| <span data-ttu-id="3fa7f-188">Valor</span><span class="sxs-lookup"><span data-stu-id="3fa7f-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fa7f-189">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3fa7f-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fa7f-190">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-190">1.1</span></span>|
|[<span data-ttu-id="3fa7f-191">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3fa7f-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fa7f-192">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3fa7f-193">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3fa7f-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="3fa7f-194">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="3fa7f-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="3fa7f-195">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o aplicativo atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="3fa7f-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="3fa7f-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="3fa7f-196">Type</span></span>

*   [<span data-ttu-id="3fa7f-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="3fa7f-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="3fa7f-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3fa7f-198">Requirements</span></span>

|<span data-ttu-id="3fa7f-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="3fa7f-199">Requirement</span></span>| <span data-ttu-id="3fa7f-200">Valor</span><span class="sxs-lookup"><span data-stu-id="3fa7f-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fa7f-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3fa7f-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fa7f-202">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-202">1.1</span></span>|
|[<span data-ttu-id="3fa7f-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3fa7f-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fa7f-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3fa7f-205">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3fa7f-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="3fa7f-206">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="3fa7f-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="3fa7f-207">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="3fa7f-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="3fa7f-208">O `RoamingSettings` objeto permite armazenar e acessar dados de um suplemento de email armazenado na caixa de correio de um usuário, para que esteja disponível para esse suplemento quando ele estiver sendo executado a partir de qualquer cliente do Outlook usado para acessar a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="3fa7f-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="3fa7f-209">Tipo</span><span class="sxs-lookup"><span data-stu-id="3fa7f-209">Type</span></span>

*   [<span data-ttu-id="3fa7f-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="3fa7f-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="3fa7f-211">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3fa7f-211">Requirements</span></span>

|<span data-ttu-id="3fa7f-212">Requisito</span><span class="sxs-lookup"><span data-stu-id="3fa7f-212">Requirement</span></span>| <span data-ttu-id="3fa7f-213">Valor</span><span class="sxs-lookup"><span data-stu-id="3fa7f-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fa7f-214">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3fa7f-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fa7f-215">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-215">1.1</span></span>|
|[<span data-ttu-id="3fa7f-216">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3fa7f-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="3fa7f-217">Restrito</span><span class="sxs-lookup"><span data-stu-id="3fa7f-217">Restricted</span></span>|
|[<span data-ttu-id="3fa7f-218">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3fa7f-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fa7f-219">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="3fa7f-220">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="3fa7f-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="3fa7f-221">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="3fa7f-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="3fa7f-222">Tipo</span><span class="sxs-lookup"><span data-stu-id="3fa7f-222">Type</span></span>

*   [<span data-ttu-id="3fa7f-223">UI</span><span class="sxs-lookup"><span data-stu-id="3fa7f-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="3fa7f-224">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3fa7f-224">Requirements</span></span>

|<span data-ttu-id="3fa7f-225">Requisito</span><span class="sxs-lookup"><span data-stu-id="3fa7f-225">Requirement</span></span>| <span data-ttu-id="3fa7f-226">Valor</span><span class="sxs-lookup"><span data-stu-id="3fa7f-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fa7f-227">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3fa7f-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fa7f-228">1.1</span><span class="sxs-lookup"><span data-stu-id="3fa7f-228">1.1</span></span>|
|[<span data-ttu-id="3fa7f-229">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3fa7f-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fa7f-230">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3fa7f-230">Compose or Read</span></span>|
