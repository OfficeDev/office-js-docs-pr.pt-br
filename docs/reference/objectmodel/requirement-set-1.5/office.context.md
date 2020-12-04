---
title: Office. Context – conjunto de requisitos 1,5
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,5.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 966c2065268d973ac8476fda839d2a6cdf038f4e
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570734"
---
# <a name="context-mailbox-requirement-set-15"></a><span data-ttu-id="785c1-103">contexto (conjunto de requisitos de caixa de correio 1,5)</span><span class="sxs-lookup"><span data-stu-id="785c1-103">context (Mailbox requirement set 1.5)</span></span>

### <a name="officecontext"></a><span data-ttu-id="785c1-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="785c1-104">[Office](office.md).context</span></span>

<span data-ttu-id="785c1-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="785c1-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="785c1-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="785c1-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="785c1-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="785c1-107">Requirements</span></span>

|<span data-ttu-id="785c1-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="785c1-108">Requirement</span></span>| <span data-ttu-id="785c1-109">Valor</span><span class="sxs-lookup"><span data-stu-id="785c1-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="785c1-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="785c1-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="785c1-111">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-111">1.1</span></span>|
|[<span data-ttu-id="785c1-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="785c1-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="785c1-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="785c1-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="785c1-114">Properties</span></span>

| <span data-ttu-id="785c1-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="785c1-115">Property</span></span> | <span data-ttu-id="785c1-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="785c1-116">Modes</span></span> | <span data-ttu-id="785c1-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="785c1-117">Return type</span></span> | <span data-ttu-id="785c1-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="785c1-118">Minimum</span></span><br><span data-ttu-id="785c1-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="785c1-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="785c1-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="785c1-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="785c1-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="785c1-121">Compose</span></span><br><span data-ttu-id="785c1-122">Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-122">Read</span></span> | <span data-ttu-id="785c1-123">String</span><span class="sxs-lookup"><span data-stu-id="785c1-123">String</span></span> | [<span data-ttu-id="785c1-124">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="785c1-125">la</span><span class="sxs-lookup"><span data-stu-id="785c1-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="785c1-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="785c1-126">Compose</span></span><br><span data-ttu-id="785c1-127">Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-127">Read</span></span> | [<span data-ttu-id="785c1-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="785c1-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="785c1-129">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="785c1-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="785c1-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="785c1-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="785c1-131">Compose</span></span><br><span data-ttu-id="785c1-132">Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-132">Read</span></span> | <span data-ttu-id="785c1-133">String</span><span class="sxs-lookup"><span data-stu-id="785c1-133">String</span></span> | [<span data-ttu-id="785c1-134">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="785c1-135">principal</span><span class="sxs-lookup"><span data-stu-id="785c1-135">host</span></span>](#host-hosttype) | <span data-ttu-id="785c1-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="785c1-136">Compose</span></span><br><span data-ttu-id="785c1-137">Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-137">Read</span></span> | [<span data-ttu-id="785c1-138">HostType</span><span class="sxs-lookup"><span data-stu-id="785c1-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="785c1-139">1,5</span><span class="sxs-lookup"><span data-stu-id="785c1-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="785c1-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="785c1-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="785c1-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="785c1-141">Compose</span></span><br><span data-ttu-id="785c1-142">Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-142">Read</span></span> | [<span data-ttu-id="785c1-143">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="785c1-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="785c1-144">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="785c1-145">plataforma</span><span class="sxs-lookup"><span data-stu-id="785c1-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="785c1-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="785c1-146">Compose</span></span><br><span data-ttu-id="785c1-147">Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-147">Read</span></span> | [<span data-ttu-id="785c1-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="785c1-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="785c1-149">1,5</span><span class="sxs-lookup"><span data-stu-id="785c1-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="785c1-150">atende</span><span class="sxs-lookup"><span data-stu-id="785c1-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="785c1-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="785c1-151">Compose</span></span><br><span data-ttu-id="785c1-152">Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-152">Read</span></span> | [<span data-ttu-id="785c1-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="785c1-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="785c1-154">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="785c1-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="785c1-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="785c1-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="785c1-156">Compose</span></span><br><span data-ttu-id="785c1-157">Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-157">Read</span></span> | [<span data-ttu-id="785c1-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="785c1-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="785c1-159">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="785c1-160">ui</span><span class="sxs-lookup"><span data-stu-id="785c1-160">ui</span></span>](#ui-ui) | <span data-ttu-id="785c1-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="785c1-161">Compose</span></span><br><span data-ttu-id="785c1-162">Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-162">Read</span></span> | [<span data-ttu-id="785c1-163">UI</span><span class="sxs-lookup"><span data-stu-id="785c1-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="785c1-164">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="785c1-165">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="785c1-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="785c1-166">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="785c1-166">contentLanguage: String</span></span>

<span data-ttu-id="785c1-167">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="785c1-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="785c1-168">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="785c1-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="785c1-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="785c1-169">Type</span></span>

*   <span data-ttu-id="785c1-170">String</span><span class="sxs-lookup"><span data-stu-id="785c1-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="785c1-171">Requisitos</span><span class="sxs-lookup"><span data-stu-id="785c1-171">Requirements</span></span>

|<span data-ttu-id="785c1-172">Requisito</span><span class="sxs-lookup"><span data-stu-id="785c1-172">Requirement</span></span>| <span data-ttu-id="785c1-173">Valor</span><span class="sxs-lookup"><span data-stu-id="785c1-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="785c1-174">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="785c1-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="785c1-175">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-175">1.1</span></span>|
|[<span data-ttu-id="785c1-176">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="785c1-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="785c1-177">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="785c1-178">Exemplo</span><span class="sxs-lookup"><span data-stu-id="785c1-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="785c1-179">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="785c1-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="785c1-180">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="785c1-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="785c1-181">Tipo</span><span class="sxs-lookup"><span data-stu-id="785c1-181">Type</span></span>

*   [<span data-ttu-id="785c1-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="785c1-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="785c1-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="785c1-183">Requirements</span></span>

|<span data-ttu-id="785c1-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="785c1-184">Requirement</span></span>| <span data-ttu-id="785c1-185">Valor</span><span class="sxs-lookup"><span data-stu-id="785c1-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="785c1-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="785c1-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="785c1-187">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-187">1.1</span></span>|
|[<span data-ttu-id="785c1-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="785c1-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="785c1-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="785c1-190">Exemplo</span><span class="sxs-lookup"><span data-stu-id="785c1-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="785c1-191">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="785c1-191">displayLanguage: String</span></span>

<span data-ttu-id="785c1-192">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="785c1-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="785c1-193">O `displayLanguage` valor reflete a configuração de **idioma de exibição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="785c1-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="785c1-194">Tipo</span><span class="sxs-lookup"><span data-stu-id="785c1-194">Type</span></span>

*   <span data-ttu-id="785c1-195">String</span><span class="sxs-lookup"><span data-stu-id="785c1-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="785c1-196">Requisitos</span><span class="sxs-lookup"><span data-stu-id="785c1-196">Requirements</span></span>

|<span data-ttu-id="785c1-197">Requisito</span><span class="sxs-lookup"><span data-stu-id="785c1-197">Requirement</span></span>| <span data-ttu-id="785c1-198">Valor</span><span class="sxs-lookup"><span data-stu-id="785c1-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="785c1-199">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="785c1-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="785c1-200">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-200">1.1</span></span>|
|[<span data-ttu-id="785c1-201">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="785c1-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="785c1-202">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="785c1-203">Exemplo</span><span class="sxs-lookup"><span data-stu-id="785c1-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="785c1-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="785c1-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="785c1-205">Obtém o aplicativo do Office que está hospedando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="785c1-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="785c1-206">Como alternativa, você pode usar a propriedade [Office. Context. Diagnostics](#diagnostics-contextinformation) para obter o host.</span><span class="sxs-lookup"><span data-stu-id="785c1-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="785c1-207">Tipo</span><span class="sxs-lookup"><span data-stu-id="785c1-207">Type</span></span>

*   [<span data-ttu-id="785c1-208">HostType</span><span class="sxs-lookup"><span data-stu-id="785c1-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="785c1-209">Requisitos</span><span class="sxs-lookup"><span data-stu-id="785c1-209">Requirements</span></span>

|<span data-ttu-id="785c1-210">Requisito</span><span class="sxs-lookup"><span data-stu-id="785c1-210">Requirement</span></span>| <span data-ttu-id="785c1-211">Valor</span><span class="sxs-lookup"><span data-stu-id="785c1-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="785c1-212">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="785c1-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="785c1-213">1,5</span><span class="sxs-lookup"><span data-stu-id="785c1-213">1.5</span></span>|
|[<span data-ttu-id="785c1-214">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="785c1-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="785c1-215">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="785c1-216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="785c1-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="785c1-217">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="785c1-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="785c1-218">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="785c1-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="785c1-219">Como alternativa, você pode usar a propriedade [Office. Context. Diagnostics](#diagnostics-contextinformation) para obter a plataforma.</span><span class="sxs-lookup"><span data-stu-id="785c1-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="785c1-220">Tipo</span><span class="sxs-lookup"><span data-stu-id="785c1-220">Type</span></span>

*   [<span data-ttu-id="785c1-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="785c1-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="785c1-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="785c1-222">Requirements</span></span>

|<span data-ttu-id="785c1-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="785c1-223">Requirement</span></span>| <span data-ttu-id="785c1-224">Valor</span><span class="sxs-lookup"><span data-stu-id="785c1-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="785c1-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="785c1-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="785c1-226">1,5</span><span class="sxs-lookup"><span data-stu-id="785c1-226">1.5</span></span>|
|[<span data-ttu-id="785c1-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="785c1-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="785c1-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="785c1-229">Exemplo</span><span class="sxs-lookup"><span data-stu-id="785c1-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="785c1-230">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="785c1-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="785c1-231">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o aplicativo atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="785c1-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="785c1-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="785c1-232">Type</span></span>

*   [<span data-ttu-id="785c1-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="785c1-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="785c1-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="785c1-234">Requirements</span></span>

|<span data-ttu-id="785c1-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="785c1-235">Requirement</span></span>| <span data-ttu-id="785c1-236">Valor</span><span class="sxs-lookup"><span data-stu-id="785c1-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="785c1-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="785c1-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="785c1-238">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-238">1.1</span></span>|
|[<span data-ttu-id="785c1-239">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="785c1-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="785c1-240">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="785c1-241">Exemplo</span><span class="sxs-lookup"><span data-stu-id="785c1-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="785c1-242">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="785c1-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="785c1-243">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="785c1-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="785c1-244">O `RoamingSettings` objeto permite armazenar e acessar dados de um suplemento de email armazenado na caixa de correio de um usuário, para que esteja disponível para esse suplemento quando ele estiver sendo executado a partir de qualquer cliente do Outlook usado para acessar a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="785c1-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="785c1-245">Tipo</span><span class="sxs-lookup"><span data-stu-id="785c1-245">Type</span></span>

*   [<span data-ttu-id="785c1-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="785c1-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="785c1-247">Requisitos</span><span class="sxs-lookup"><span data-stu-id="785c1-247">Requirements</span></span>

|<span data-ttu-id="785c1-248">Requisito</span><span class="sxs-lookup"><span data-stu-id="785c1-248">Requirement</span></span>| <span data-ttu-id="785c1-249">Valor</span><span class="sxs-lookup"><span data-stu-id="785c1-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="785c1-250">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="785c1-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="785c1-251">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-251">1.1</span></span>|
|[<span data-ttu-id="785c1-252">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="785c1-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="785c1-253">Restrito</span><span class="sxs-lookup"><span data-stu-id="785c1-253">Restricted</span></span>|
|[<span data-ttu-id="785c1-254">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="785c1-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="785c1-255">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="785c1-256">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="785c1-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="785c1-257">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="785c1-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="785c1-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="785c1-258">Type</span></span>

*   [<span data-ttu-id="785c1-259">UI</span><span class="sxs-lookup"><span data-stu-id="785c1-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="785c1-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="785c1-260">Requirements</span></span>

|<span data-ttu-id="785c1-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="785c1-261">Requirement</span></span>| <span data-ttu-id="785c1-262">Valor</span><span class="sxs-lookup"><span data-stu-id="785c1-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="785c1-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="785c1-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="785c1-264">1.1</span><span class="sxs-lookup"><span data-stu-id="785c1-264">1.1</span></span>|
|[<span data-ttu-id="785c1-265">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="785c1-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="785c1-266">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="785c1-266">Compose or Read</span></span>|
