---
title: Office. Context – conjunto de requisitos 1,4
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,4.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 139ca19d059b303976117cd259fa90c58e5eeb13
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293734"
---
# <a name="context-mailbox-requirement-set-14"></a><span data-ttu-id="ebdf2-103">contexto (conjunto de requisitos de caixa de correio 1,4)</span><span class="sxs-lookup"><span data-stu-id="ebdf2-103">context (Mailbox requirement set 1.4)</span></span>

### <a name="officecontext"></a><span data-ttu-id="ebdf2-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="ebdf2-104">[Office](office.md).context</span></span>

<span data-ttu-id="ebdf2-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="ebdf2-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="ebdf2-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.4).</span><span class="sxs-lookup"><span data-stu-id="ebdf2-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ebdf2-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ebdf2-107">Requirements</span></span>

|<span data-ttu-id="ebdf2-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="ebdf2-108">Requirement</span></span>| <span data-ttu-id="ebdf2-109">Valor</span><span class="sxs-lookup"><span data-stu-id="ebdf2-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebdf2-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ebdf2-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebdf2-111">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-111">1.1</span></span>|
|[<span data-ttu-id="ebdf2-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ebdf2-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ebdf2-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ebdf2-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="ebdf2-114">Properties</span></span>

| <span data-ttu-id="ebdf2-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="ebdf2-115">Property</span></span> | <span data-ttu-id="ebdf2-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="ebdf2-116">Modes</span></span> | <span data-ttu-id="ebdf2-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="ebdf2-117">Return type</span></span> | <span data-ttu-id="ebdf2-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="ebdf2-118">Minimum</span></span><br><span data-ttu-id="ebdf2-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="ebdf2-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ebdf2-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="ebdf2-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="ebdf2-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="ebdf2-121">Compose</span></span><br><span data-ttu-id="ebdf2-122">Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-122">Read</span></span> | <span data-ttu-id="ebdf2-123">String</span><span class="sxs-lookup"><span data-stu-id="ebdf2-123">String</span></span> | [<span data-ttu-id="ebdf2-124">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ebdf2-125">la</span><span class="sxs-lookup"><span data-stu-id="ebdf2-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="ebdf2-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="ebdf2-126">Compose</span></span><br><span data-ttu-id="ebdf2-127">Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-127">Read</span></span> | [<span data-ttu-id="ebdf2-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ebdf2-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4) | [<span data-ttu-id="ebdf2-129">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ebdf2-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="ebdf2-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="ebdf2-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="ebdf2-131">Compose</span></span><br><span data-ttu-id="ebdf2-132">Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-132">Read</span></span> | <span data-ttu-id="ebdf2-133">String</span><span class="sxs-lookup"><span data-stu-id="ebdf2-133">String</span></span> | [<span data-ttu-id="ebdf2-134">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ebdf2-135">principal</span><span class="sxs-lookup"><span data-stu-id="ebdf2-135">host</span></span>](#host-hosttype) | <span data-ttu-id="ebdf2-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="ebdf2-136">Compose</span></span><br><span data-ttu-id="ebdf2-137">Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-137">Read</span></span> | [<span data-ttu-id="ebdf2-138">HostType</span><span class="sxs-lookup"><span data-stu-id="ebdf2-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.4) | [<span data-ttu-id="ebdf2-139">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ebdf2-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="ebdf2-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="ebdf2-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="ebdf2-141">Compose</span></span><br><span data-ttu-id="ebdf2-142">Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-142">Read</span></span> | [<span data-ttu-id="ebdf2-143">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="ebdf2-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4) | [<span data-ttu-id="ebdf2-144">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ebdf2-145">plataforma</span><span class="sxs-lookup"><span data-stu-id="ebdf2-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="ebdf2-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="ebdf2-146">Compose</span></span><br><span data-ttu-id="ebdf2-147">Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-147">Read</span></span> | [<span data-ttu-id="ebdf2-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ebdf2-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.4) | [<span data-ttu-id="ebdf2-149">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ebdf2-150">atende</span><span class="sxs-lookup"><span data-stu-id="ebdf2-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="ebdf2-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="ebdf2-151">Compose</span></span><br><span data-ttu-id="ebdf2-152">Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-152">Read</span></span> | [<span data-ttu-id="ebdf2-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ebdf2-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4) | [<span data-ttu-id="ebdf2-154">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ebdf2-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="ebdf2-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="ebdf2-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="ebdf2-156">Compose</span></span><br><span data-ttu-id="ebdf2-157">Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-157">Read</span></span> | [<span data-ttu-id="ebdf2-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ebdf2-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4) | [<span data-ttu-id="ebdf2-159">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ebdf2-160">ui</span><span class="sxs-lookup"><span data-stu-id="ebdf2-160">ui</span></span>](#ui-ui) | <span data-ttu-id="ebdf2-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="ebdf2-161">Compose</span></span><br><span data-ttu-id="ebdf2-162">Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-162">Read</span></span> | [<span data-ttu-id="ebdf2-163">UI</span><span class="sxs-lookup"><span data-stu-id="ebdf2-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4) | [<span data-ttu-id="ebdf2-164">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="ebdf2-165">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="ebdf2-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="ebdf2-166">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ebdf2-166">contentLanguage: String</span></span>

<span data-ttu-id="ebdf2-167">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="ebdf2-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="ebdf2-168">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="ebdf2-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ebdf2-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-169">Type</span></span>

*   <span data-ttu-id="ebdf2-170">String</span><span class="sxs-lookup"><span data-stu-id="ebdf2-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ebdf2-171">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ebdf2-171">Requirements</span></span>

|<span data-ttu-id="ebdf2-172">Requisito</span><span class="sxs-lookup"><span data-stu-id="ebdf2-172">Requirement</span></span>| <span data-ttu-id="ebdf2-173">Valor</span><span class="sxs-lookup"><span data-stu-id="ebdf2-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebdf2-174">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ebdf2-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebdf2-175">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-175">1.1</span></span>|
|[<span data-ttu-id="ebdf2-176">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ebdf2-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ebdf2-177">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ebdf2-178">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="ebdf2-179">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="ebdf2-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="ebdf2-180">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="ebdf2-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ebdf2-181">Tipo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-181">Type</span></span>

*   [<span data-ttu-id="ebdf2-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ebdf2-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="ebdf2-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ebdf2-183">Requirements</span></span>

|<span data-ttu-id="ebdf2-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="ebdf2-184">Requirement</span></span>| <span data-ttu-id="ebdf2-185">Valor</span><span class="sxs-lookup"><span data-stu-id="ebdf2-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebdf2-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ebdf2-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebdf2-187">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-187">1.1</span></span>|
|[<span data-ttu-id="ebdf2-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ebdf2-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ebdf2-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ebdf2-190">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="ebdf2-191">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ebdf2-191">displayLanguage: String</span></span>

<span data-ttu-id="ebdf2-192">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="ebdf2-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="ebdf2-193">O `displayLanguage` valor reflete a configuração de **idioma de exibição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="ebdf2-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ebdf2-194">Tipo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-194">Type</span></span>

*   <span data-ttu-id="ebdf2-195">String</span><span class="sxs-lookup"><span data-stu-id="ebdf2-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ebdf2-196">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ebdf2-196">Requirements</span></span>

|<span data-ttu-id="ebdf2-197">Requisito</span><span class="sxs-lookup"><span data-stu-id="ebdf2-197">Requirement</span></span>| <span data-ttu-id="ebdf2-198">Valor</span><span class="sxs-lookup"><span data-stu-id="ebdf2-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebdf2-199">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ebdf2-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebdf2-200">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-200">1.1</span></span>|
|[<span data-ttu-id="ebdf2-201">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ebdf2-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ebdf2-202">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ebdf2-203">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="ebdf2-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="ebdf2-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="ebdf2-205">Obtém o aplicativo do Office que está hospedando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="ebdf2-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="ebdf2-206">Tipo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-206">Type</span></span>

*   [<span data-ttu-id="ebdf2-207">HostType</span><span class="sxs-lookup"><span data-stu-id="ebdf2-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="ebdf2-208">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ebdf2-208">Requirements</span></span>

|<span data-ttu-id="ebdf2-209">Requisito</span><span class="sxs-lookup"><span data-stu-id="ebdf2-209">Requirement</span></span>| <span data-ttu-id="ebdf2-210">Valor</span><span class="sxs-lookup"><span data-stu-id="ebdf2-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebdf2-211">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ebdf2-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebdf2-212">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-212">1.1</span></span>|
|[<span data-ttu-id="ebdf2-213">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ebdf2-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ebdf2-214">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ebdf2-215">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="ebdf2-216">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="ebdf2-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="ebdf2-217">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="ebdf2-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ebdf2-218">Tipo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-218">Type</span></span>

*   [<span data-ttu-id="ebdf2-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ebdf2-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="ebdf2-220">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ebdf2-220">Requirements</span></span>

|<span data-ttu-id="ebdf2-221">Requisito</span><span class="sxs-lookup"><span data-stu-id="ebdf2-221">Requirement</span></span>| <span data-ttu-id="ebdf2-222">Valor</span><span class="sxs-lookup"><span data-stu-id="ebdf2-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebdf2-223">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ebdf2-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebdf2-224">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-224">1.1</span></span>|
|[<span data-ttu-id="ebdf2-225">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ebdf2-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ebdf2-226">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ebdf2-227">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="ebdf2-228">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="ebdf2-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="ebdf2-229">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o aplicativo atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="ebdf2-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="ebdf2-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-230">Type</span></span>

*   [<span data-ttu-id="ebdf2-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ebdf2-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="ebdf2-232">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ebdf2-232">Requirements</span></span>

|<span data-ttu-id="ebdf2-233">Requisito</span><span class="sxs-lookup"><span data-stu-id="ebdf2-233">Requirement</span></span>| <span data-ttu-id="ebdf2-234">Valor</span><span class="sxs-lookup"><span data-stu-id="ebdf2-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebdf2-235">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ebdf2-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebdf2-236">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-236">1.1</span></span>|
|[<span data-ttu-id="ebdf2-237">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ebdf2-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ebdf2-238">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ebdf2-239">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="ebdf2-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="ebdf2-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="ebdf2-241">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="ebdf2-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ebdf2-242">O `RoamingSettings` objeto permite armazenar e acessar dados de um suplemento de email armazenado na caixa de correio de um usuário, para que esteja disponível para esse suplemento quando ele estiver sendo executado a partir de qualquer cliente do Outlook usado para acessar a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="ebdf2-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ebdf2-243">Tipo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-243">Type</span></span>

*   [<span data-ttu-id="ebdf2-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ebdf2-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="ebdf2-245">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ebdf2-245">Requirements</span></span>

|<span data-ttu-id="ebdf2-246">Requisito</span><span class="sxs-lookup"><span data-stu-id="ebdf2-246">Requirement</span></span>| <span data-ttu-id="ebdf2-247">Valor</span><span class="sxs-lookup"><span data-stu-id="ebdf2-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebdf2-248">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ebdf2-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebdf2-249">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-249">1.1</span></span>|
|[<span data-ttu-id="ebdf2-250">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="ebdf2-251">Restrito</span><span class="sxs-lookup"><span data-stu-id="ebdf2-251">Restricted</span></span>|
|[<span data-ttu-id="ebdf2-252">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ebdf2-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ebdf2-253">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="ebdf2-254">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="ebdf2-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="ebdf2-255">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="ebdf2-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ebdf2-256">Tipo</span><span class="sxs-lookup"><span data-stu-id="ebdf2-256">Type</span></span>

*   [<span data-ttu-id="ebdf2-257">UI</span><span class="sxs-lookup"><span data-stu-id="ebdf2-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="ebdf2-258">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ebdf2-258">Requirements</span></span>

|<span data-ttu-id="ebdf2-259">Requisito</span><span class="sxs-lookup"><span data-stu-id="ebdf2-259">Requirement</span></span>| <span data-ttu-id="ebdf2-260">Valor</span><span class="sxs-lookup"><span data-stu-id="ebdf2-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="ebdf2-261">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ebdf2-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ebdf2-262">1.1</span><span class="sxs-lookup"><span data-stu-id="ebdf2-262">1.1</span></span>|
|[<span data-ttu-id="ebdf2-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ebdf2-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ebdf2-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ebdf2-264">Compose or Read</span></span>|
