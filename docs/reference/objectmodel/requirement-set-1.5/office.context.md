---
title: Office. Context – conjunto de requisitos 1,5
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,5.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 4587609e08d48e5c9e0cc74ec2c9050af0c3877a
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293748"
---
# <a name="context-mailbox-requirement-set-15"></a><span data-ttu-id="dd405-103">contexto (conjunto de requisitos de caixa de correio 1,5)</span><span class="sxs-lookup"><span data-stu-id="dd405-103">context (Mailbox requirement set 1.5)</span></span>

### <a name="officecontext"></a><span data-ttu-id="dd405-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="dd405-104">[Office](office.md).context</span></span>

<span data-ttu-id="dd405-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="dd405-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="dd405-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.5).</span><span class="sxs-lookup"><span data-stu-id="dd405-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.5).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd405-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dd405-107">Requirements</span></span>

|<span data-ttu-id="dd405-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="dd405-108">Requirement</span></span>| <span data-ttu-id="dd405-109">Valor</span><span class="sxs-lookup"><span data-stu-id="dd405-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd405-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dd405-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd405-111">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-111">1.1</span></span>|
|[<span data-ttu-id="dd405-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dd405-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd405-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="dd405-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="dd405-114">Properties</span></span>

| <span data-ttu-id="dd405-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="dd405-115">Property</span></span> | <span data-ttu-id="dd405-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="dd405-116">Modes</span></span> | <span data-ttu-id="dd405-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="dd405-117">Return type</span></span> | <span data-ttu-id="dd405-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="dd405-118">Minimum</span></span><br><span data-ttu-id="dd405-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="dd405-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dd405-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="dd405-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="dd405-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="dd405-121">Compose</span></span><br><span data-ttu-id="dd405-122">Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-122">Read</span></span> | <span data-ttu-id="dd405-123">String</span><span class="sxs-lookup"><span data-stu-id="dd405-123">String</span></span> | [<span data-ttu-id="dd405-124">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dd405-125">la</span><span class="sxs-lookup"><span data-stu-id="dd405-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="dd405-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="dd405-126">Compose</span></span><br><span data-ttu-id="dd405-127">Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-127">Read</span></span> | [<span data-ttu-id="dd405-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="dd405-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.5) | [<span data-ttu-id="dd405-129">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dd405-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="dd405-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="dd405-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="dd405-131">Compose</span></span><br><span data-ttu-id="dd405-132">Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-132">Read</span></span> | <span data-ttu-id="dd405-133">String</span><span class="sxs-lookup"><span data-stu-id="dd405-133">String</span></span> | [<span data-ttu-id="dd405-134">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dd405-135">principal</span><span class="sxs-lookup"><span data-stu-id="dd405-135">host</span></span>](#host-hosttype) | <span data-ttu-id="dd405-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="dd405-136">Compose</span></span><br><span data-ttu-id="dd405-137">Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-137">Read</span></span> | [<span data-ttu-id="dd405-138">HostType</span><span class="sxs-lookup"><span data-stu-id="dd405-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.5) | [<span data-ttu-id="dd405-139">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dd405-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="dd405-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="dd405-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="dd405-141">Compose</span></span><br><span data-ttu-id="dd405-142">Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-142">Read</span></span> | [<span data-ttu-id="dd405-143">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="dd405-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5) | [<span data-ttu-id="dd405-144">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dd405-145">plataforma</span><span class="sxs-lookup"><span data-stu-id="dd405-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="dd405-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="dd405-146">Compose</span></span><br><span data-ttu-id="dd405-147">Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-147">Read</span></span> | [<span data-ttu-id="dd405-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="dd405-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.5) | [<span data-ttu-id="dd405-149">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dd405-150">atende</span><span class="sxs-lookup"><span data-stu-id="dd405-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="dd405-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="dd405-151">Compose</span></span><br><span data-ttu-id="dd405-152">Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-152">Read</span></span> | [<span data-ttu-id="dd405-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="dd405-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.5) | [<span data-ttu-id="dd405-154">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dd405-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="dd405-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="dd405-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="dd405-156">Compose</span></span><br><span data-ttu-id="dd405-157">Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-157">Read</span></span> | [<span data-ttu-id="dd405-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="dd405-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.5) | [<span data-ttu-id="dd405-159">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dd405-160">ui</span><span class="sxs-lookup"><span data-stu-id="dd405-160">ui</span></span>](#ui-ui) | <span data-ttu-id="dd405-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="dd405-161">Compose</span></span><br><span data-ttu-id="dd405-162">Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-162">Read</span></span> | [<span data-ttu-id="dd405-163">UI</span><span class="sxs-lookup"><span data-stu-id="dd405-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.5) | [<span data-ttu-id="dd405-164">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="dd405-165">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="dd405-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="dd405-166">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dd405-166">contentLanguage: String</span></span>

<span data-ttu-id="dd405-167">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="dd405-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="dd405-168">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="dd405-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="dd405-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="dd405-169">Type</span></span>

*   <span data-ttu-id="dd405-170">String</span><span class="sxs-lookup"><span data-stu-id="dd405-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd405-171">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dd405-171">Requirements</span></span>

|<span data-ttu-id="dd405-172">Requisito</span><span class="sxs-lookup"><span data-stu-id="dd405-172">Requirement</span></span>| <span data-ttu-id="dd405-173">Valor</span><span class="sxs-lookup"><span data-stu-id="dd405-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd405-174">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dd405-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd405-175">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-175">1.1</span></span>|
|[<span data-ttu-id="dd405-176">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dd405-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd405-177">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd405-178">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dd405-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="dd405-179">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="dd405-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="dd405-180">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="dd405-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="dd405-181">Tipo</span><span class="sxs-lookup"><span data-stu-id="dd405-181">Type</span></span>

*   [<span data-ttu-id="dd405-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="dd405-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="dd405-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dd405-183">Requirements</span></span>

|<span data-ttu-id="dd405-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="dd405-184">Requirement</span></span>| <span data-ttu-id="dd405-185">Valor</span><span class="sxs-lookup"><span data-stu-id="dd405-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd405-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dd405-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd405-187">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-187">1.1</span></span>|
|[<span data-ttu-id="dd405-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dd405-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd405-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd405-190">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dd405-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="dd405-191">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dd405-191">displayLanguage: String</span></span>

<span data-ttu-id="dd405-192">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="dd405-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="dd405-193">O `displayLanguage` valor reflete a configuração de **idioma de exibição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="dd405-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="dd405-194">Tipo</span><span class="sxs-lookup"><span data-stu-id="dd405-194">Type</span></span>

*   <span data-ttu-id="dd405-195">String</span><span class="sxs-lookup"><span data-stu-id="dd405-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd405-196">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dd405-196">Requirements</span></span>

|<span data-ttu-id="dd405-197">Requisito</span><span class="sxs-lookup"><span data-stu-id="dd405-197">Requirement</span></span>| <span data-ttu-id="dd405-198">Valor</span><span class="sxs-lookup"><span data-stu-id="dd405-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd405-199">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dd405-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd405-200">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-200">1.1</span></span>|
|[<span data-ttu-id="dd405-201">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dd405-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd405-202">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd405-203">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dd405-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="dd405-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="dd405-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="dd405-205">Obtém o aplicativo do Office que está hospedando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="dd405-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="dd405-206">Tipo</span><span class="sxs-lookup"><span data-stu-id="dd405-206">Type</span></span>

*   [<span data-ttu-id="dd405-207">HostType</span><span class="sxs-lookup"><span data-stu-id="dd405-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="dd405-208">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dd405-208">Requirements</span></span>

|<span data-ttu-id="dd405-209">Requisito</span><span class="sxs-lookup"><span data-stu-id="dd405-209">Requirement</span></span>| <span data-ttu-id="dd405-210">Valor</span><span class="sxs-lookup"><span data-stu-id="dd405-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd405-211">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dd405-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd405-212">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-212">1.1</span></span>|
|[<span data-ttu-id="dd405-213">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dd405-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd405-214">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd405-215">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dd405-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="dd405-216">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="dd405-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="dd405-217">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="dd405-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="dd405-218">Tipo</span><span class="sxs-lookup"><span data-stu-id="dd405-218">Type</span></span>

*   [<span data-ttu-id="dd405-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="dd405-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="dd405-220">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dd405-220">Requirements</span></span>

|<span data-ttu-id="dd405-221">Requisito</span><span class="sxs-lookup"><span data-stu-id="dd405-221">Requirement</span></span>| <span data-ttu-id="dd405-222">Valor</span><span class="sxs-lookup"><span data-stu-id="dd405-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd405-223">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dd405-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd405-224">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-224">1.1</span></span>|
|[<span data-ttu-id="dd405-225">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dd405-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd405-226">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd405-227">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dd405-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="dd405-228">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="dd405-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="dd405-229">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o aplicativo atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="dd405-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="dd405-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="dd405-230">Type</span></span>

*   [<span data-ttu-id="dd405-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="dd405-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="dd405-232">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dd405-232">Requirements</span></span>

|<span data-ttu-id="dd405-233">Requisito</span><span class="sxs-lookup"><span data-stu-id="dd405-233">Requirement</span></span>| <span data-ttu-id="dd405-234">Valor</span><span class="sxs-lookup"><span data-stu-id="dd405-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd405-235">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dd405-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd405-236">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-236">1.1</span></span>|
|[<span data-ttu-id="dd405-237">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dd405-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd405-238">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd405-239">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dd405-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="dd405-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="dd405-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="dd405-241">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="dd405-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="dd405-242">O `RoamingSettings` objeto permite armazenar e acessar dados de um suplemento de email armazenado na caixa de correio de um usuário, para que esteja disponível para esse suplemento quando ele estiver sendo executado a partir de qualquer cliente do Outlook usado para acessar a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="dd405-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="dd405-243">Tipo</span><span class="sxs-lookup"><span data-stu-id="dd405-243">Type</span></span>

*   [<span data-ttu-id="dd405-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="dd405-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="dd405-245">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dd405-245">Requirements</span></span>

|<span data-ttu-id="dd405-246">Requisito</span><span class="sxs-lookup"><span data-stu-id="dd405-246">Requirement</span></span>| <span data-ttu-id="dd405-247">Valor</span><span class="sxs-lookup"><span data-stu-id="dd405-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd405-248">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dd405-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd405-249">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-249">1.1</span></span>|
|[<span data-ttu-id="dd405-250">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="dd405-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="dd405-251">Restrito</span><span class="sxs-lookup"><span data-stu-id="dd405-251">Restricted</span></span>|
|[<span data-ttu-id="dd405-252">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dd405-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd405-253">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="dd405-254">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="dd405-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="dd405-255">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="dd405-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="dd405-256">Tipo</span><span class="sxs-lookup"><span data-stu-id="dd405-256">Type</span></span>

*   [<span data-ttu-id="dd405-257">UI</span><span class="sxs-lookup"><span data-stu-id="dd405-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="dd405-258">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dd405-258">Requirements</span></span>

|<span data-ttu-id="dd405-259">Requisito</span><span class="sxs-lookup"><span data-stu-id="dd405-259">Requirement</span></span>| <span data-ttu-id="dd405-260">Valor</span><span class="sxs-lookup"><span data-stu-id="dd405-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd405-261">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dd405-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd405-262">1.1</span><span class="sxs-lookup"><span data-stu-id="dd405-262">1.1</span></span>|
|[<span data-ttu-id="dd405-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dd405-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd405-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="dd405-264">Compose or Read</span></span>|
