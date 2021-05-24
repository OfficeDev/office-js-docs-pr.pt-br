---
title: Office.context - conjunto de requisitos 1.8
description: Office. Membros do objeto Context disponíveis para Outlook de entrada usando o conjunto de requisitos da API de Caixa de Correio 1.8.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 99573d9984c571c99461e90e8bdccdca35fe30b7
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590964"
---
# <a name="context-mailbox-requirement-set-18"></a><span data-ttu-id="93061-103">context (Conjunto de requisitos de caixa de correio 1.8)</span><span class="sxs-lookup"><span data-stu-id="93061-103">context (Mailbox requirement set 1.8)</span></span>

### <a name="officecontext"></a><span data-ttu-id="93061-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="93061-104">[Office](office.md).context</span></span>

<span data-ttu-id="93061-105">Office.context fornece interfaces compartilhadas que são usadas por complementos em todos os Office aplicativos.</span><span class="sxs-lookup"><span data-stu-id="93061-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="93061-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="93061-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="93061-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="93061-107">Requirements</span></span>

|<span data-ttu-id="93061-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="93061-108">Requirement</span></span>| <span data-ttu-id="93061-109">Valor</span><span class="sxs-lookup"><span data-stu-id="93061-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="93061-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="93061-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93061-111">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-111">1.1</span></span>|
|[<span data-ttu-id="93061-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="93061-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93061-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="93061-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="93061-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="93061-114">Properties</span></span>

| <span data-ttu-id="93061-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="93061-115">Property</span></span> | <span data-ttu-id="93061-116">Modos</span><span class="sxs-lookup"><span data-stu-id="93061-116">Modes</span></span> | <span data-ttu-id="93061-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="93061-117">Return type</span></span> | <span data-ttu-id="93061-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="93061-118">Minimum</span></span><br><span data-ttu-id="93061-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="93061-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="93061-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="93061-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="93061-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="93061-121">Compose</span></span><br><span data-ttu-id="93061-122">Ler</span><span class="sxs-lookup"><span data-stu-id="93061-122">Read</span></span> | <span data-ttu-id="93061-123">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="93061-123">String</span></span> | [<span data-ttu-id="93061-124">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="93061-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="93061-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="93061-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="93061-126">Compose</span></span><br><span data-ttu-id="93061-127">Ler</span><span class="sxs-lookup"><span data-stu-id="93061-127">Read</span></span> | [<span data-ttu-id="93061-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="93061-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="93061-129">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="93061-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="93061-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="93061-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="93061-131">Compose</span></span><br><span data-ttu-id="93061-132">Ler</span><span class="sxs-lookup"><span data-stu-id="93061-132">Read</span></span> | <span data-ttu-id="93061-133">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="93061-133">String</span></span> | [<span data-ttu-id="93061-134">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="93061-135">host</span><span class="sxs-lookup"><span data-stu-id="93061-135">host</span></span>](#host-hosttype) | <span data-ttu-id="93061-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="93061-136">Compose</span></span><br><span data-ttu-id="93061-137">Ler</span><span class="sxs-lookup"><span data-stu-id="93061-137">Read</span></span> | [<span data-ttu-id="93061-138">HostType</span><span class="sxs-lookup"><span data-stu-id="93061-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="93061-139">1.5</span><span class="sxs-lookup"><span data-stu-id="93061-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="93061-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="93061-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="93061-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="93061-141">Compose</span></span><br><span data-ttu-id="93061-142">Ler</span><span class="sxs-lookup"><span data-stu-id="93061-142">Read</span></span> | [<span data-ttu-id="93061-143">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="93061-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="93061-144">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="93061-145">platform</span><span class="sxs-lookup"><span data-stu-id="93061-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="93061-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="93061-146">Compose</span></span><br><span data-ttu-id="93061-147">Ler</span><span class="sxs-lookup"><span data-stu-id="93061-147">Read</span></span> | [<span data-ttu-id="93061-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="93061-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="93061-149">1.5</span><span class="sxs-lookup"><span data-stu-id="93061-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="93061-150">requirements</span><span class="sxs-lookup"><span data-stu-id="93061-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="93061-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="93061-151">Compose</span></span><br><span data-ttu-id="93061-152">Ler</span><span class="sxs-lookup"><span data-stu-id="93061-152">Read</span></span> | [<span data-ttu-id="93061-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="93061-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="93061-154">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="93061-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="93061-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="93061-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="93061-156">Compose</span></span><br><span data-ttu-id="93061-157">Ler</span><span class="sxs-lookup"><span data-stu-id="93061-157">Read</span></span> | [<span data-ttu-id="93061-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="93061-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="93061-159">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="93061-160">ui</span><span class="sxs-lookup"><span data-stu-id="93061-160">ui</span></span>](#ui-ui) | <span data-ttu-id="93061-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="93061-161">Compose</span></span><br><span data-ttu-id="93061-162">Ler</span><span class="sxs-lookup"><span data-stu-id="93061-162">Read</span></span> | [<span data-ttu-id="93061-163">UI</span><span class="sxs-lookup"><span data-stu-id="93061-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="93061-164">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="93061-165">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="93061-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="93061-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="93061-166">contentLanguage: String</span></span>

<span data-ttu-id="93061-167">Obtém a localidade (idioma) especificada pelo usuário para editar o item.</span><span class="sxs-lookup"><span data-stu-id="93061-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="93061-168">O `contentLanguage` valor reflete a **configuração** atual de Idioma de Edição especificada com opções de > de arquivo **> idioma** no aplicativo Office cliente.</span><span class="sxs-lookup"><span data-stu-id="93061-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="93061-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="93061-169">Type</span></span>

*   <span data-ttu-id="93061-170">String</span><span class="sxs-lookup"><span data-stu-id="93061-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="93061-171">Requisitos</span><span class="sxs-lookup"><span data-stu-id="93061-171">Requirements</span></span>

|<span data-ttu-id="93061-172">Requisito</span><span class="sxs-lookup"><span data-stu-id="93061-172">Requirement</span></span>| <span data-ttu-id="93061-173">Valor</span><span class="sxs-lookup"><span data-stu-id="93061-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="93061-174">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="93061-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93061-175">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-175">1.1</span></span>|
|[<span data-ttu-id="93061-176">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="93061-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93061-177">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="93061-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="93061-178">Exemplo</span><span class="sxs-lookup"><span data-stu-id="93061-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="93061-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="93061-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="93061-180">Obtém informações sobre o ambiente no qual o complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="93061-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="93061-181">Tipo</span><span class="sxs-lookup"><span data-stu-id="93061-181">Type</span></span>

*   [<span data-ttu-id="93061-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="93061-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="93061-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="93061-183">Requirements</span></span>

|<span data-ttu-id="93061-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="93061-184">Requirement</span></span>| <span data-ttu-id="93061-185">Valor</span><span class="sxs-lookup"><span data-stu-id="93061-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="93061-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="93061-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93061-187">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-187">1.1</span></span>|
|[<span data-ttu-id="93061-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="93061-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93061-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="93061-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="93061-190">Exemplo</span><span class="sxs-lookup"><span data-stu-id="93061-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="93061-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="93061-191">displayLanguage: String</span></span>

<span data-ttu-id="93061-192">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente Office cliente.</span><span class="sxs-lookup"><span data-stu-id="93061-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="93061-193">O valor reflete a configuração atual de Idioma de Exibição especificada com Opções > > Idioma no aplicativo Office `displayLanguage` cliente.  </span><span class="sxs-lookup"><span data-stu-id="93061-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="93061-194">Tipo</span><span class="sxs-lookup"><span data-stu-id="93061-194">Type</span></span>

*   <span data-ttu-id="93061-195">String</span><span class="sxs-lookup"><span data-stu-id="93061-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="93061-196">Requisitos</span><span class="sxs-lookup"><span data-stu-id="93061-196">Requirements</span></span>

|<span data-ttu-id="93061-197">Requisito</span><span class="sxs-lookup"><span data-stu-id="93061-197">Requirement</span></span>| <span data-ttu-id="93061-198">Valor</span><span class="sxs-lookup"><span data-stu-id="93061-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="93061-199">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="93061-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93061-200">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-200">1.1</span></span>|
|[<span data-ttu-id="93061-201">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="93061-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93061-202">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="93061-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="93061-203">Exemplo</span><span class="sxs-lookup"><span data-stu-id="93061-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="93061-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="93061-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="93061-205">Obtém o Office aplicativo que está hospedando o complemento.</span><span class="sxs-lookup"><span data-stu-id="93061-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="93061-206">Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter o host.</span><span class="sxs-lookup"><span data-stu-id="93061-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="93061-207">Tipo</span><span class="sxs-lookup"><span data-stu-id="93061-207">Type</span></span>

*   [<span data-ttu-id="93061-208">HostType</span><span class="sxs-lookup"><span data-stu-id="93061-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="93061-209">Requisitos</span><span class="sxs-lookup"><span data-stu-id="93061-209">Requirements</span></span>

|<span data-ttu-id="93061-210">Requisito</span><span class="sxs-lookup"><span data-stu-id="93061-210">Requirement</span></span>| <span data-ttu-id="93061-211">Valor</span><span class="sxs-lookup"><span data-stu-id="93061-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="93061-212">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="93061-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93061-213">1,5</span><span class="sxs-lookup"><span data-stu-id="93061-213">1.5</span></span>|
|[<span data-ttu-id="93061-214">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="93061-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93061-215">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="93061-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="93061-216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="93061-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="93061-217">plataforma: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="93061-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="93061-218">Fornece a plataforma na qual o complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="93061-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="93061-219">Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter a plataforma.</span><span class="sxs-lookup"><span data-stu-id="93061-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="93061-220">Tipo</span><span class="sxs-lookup"><span data-stu-id="93061-220">Type</span></span>

*   [<span data-ttu-id="93061-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="93061-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="93061-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="93061-222">Requirements</span></span>

|<span data-ttu-id="93061-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="93061-223">Requirement</span></span>| <span data-ttu-id="93061-224">Valor</span><span class="sxs-lookup"><span data-stu-id="93061-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="93061-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="93061-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93061-226">1,5</span><span class="sxs-lookup"><span data-stu-id="93061-226">1.5</span></span>|
|[<span data-ttu-id="93061-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="93061-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93061-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="93061-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="93061-229">Exemplo</span><span class="sxs-lookup"><span data-stu-id="93061-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="93061-230">requirements: [RequirementsSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="93061-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="93061-231">Fornece um método para determinar quais conjuntos de requisitos são suportados no aplicativo e na plataforma atual.</span><span class="sxs-lookup"><span data-stu-id="93061-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="93061-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="93061-232">Type</span></span>

*   [<span data-ttu-id="93061-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="93061-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="93061-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="93061-234">Requirements</span></span>

|<span data-ttu-id="93061-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="93061-235">Requirement</span></span>| <span data-ttu-id="93061-236">Valor</span><span class="sxs-lookup"><span data-stu-id="93061-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="93061-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="93061-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93061-238">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-238">1.1</span></span>|
|[<span data-ttu-id="93061-239">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="93061-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93061-240">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="93061-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="93061-241">Exemplo</span><span class="sxs-lookup"><span data-stu-id="93061-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="93061-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="93061-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="93061-243">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="93061-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="93061-244">O objeto permite que você armazene e acesse dados para um complemento de email armazenado na caixa de correio de um usuário, de modo que está disponível para esse complemento quando ele está sendo executado de qualquer cliente Outlook usado para acessar essa caixa de `RoamingSettings` correio.</span><span class="sxs-lookup"><span data-stu-id="93061-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="93061-245">Tipo</span><span class="sxs-lookup"><span data-stu-id="93061-245">Type</span></span>

*   [<span data-ttu-id="93061-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="93061-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="93061-247">Requisitos</span><span class="sxs-lookup"><span data-stu-id="93061-247">Requirements</span></span>

|<span data-ttu-id="93061-248">Requisito</span><span class="sxs-lookup"><span data-stu-id="93061-248">Requirement</span></span>| <span data-ttu-id="93061-249">Valor</span><span class="sxs-lookup"><span data-stu-id="93061-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="93061-250">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="93061-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93061-251">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-251">1.1</span></span>|
|[<span data-ttu-id="93061-252">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="93061-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="93061-253">Restrito</span><span class="sxs-lookup"><span data-stu-id="93061-253">Restricted</span></span>|
|[<span data-ttu-id="93061-254">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="93061-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93061-255">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="93061-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="93061-256">interface do usuário: [interface do usuário](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="93061-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="93061-257">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus Office de usuário.</span><span class="sxs-lookup"><span data-stu-id="93061-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="93061-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="93061-258">Type</span></span>

*   [<span data-ttu-id="93061-259">UI</span><span class="sxs-lookup"><span data-stu-id="93061-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="93061-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="93061-260">Requirements</span></span>

|<span data-ttu-id="93061-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="93061-261">Requirement</span></span>| <span data-ttu-id="93061-262">Valor</span><span class="sxs-lookup"><span data-stu-id="93061-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="93061-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="93061-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93061-264">1.1</span><span class="sxs-lookup"><span data-stu-id="93061-264">1.1</span></span>|
|[<span data-ttu-id="93061-265">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="93061-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93061-266">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="93061-266">Compose or Read</span></span>|
