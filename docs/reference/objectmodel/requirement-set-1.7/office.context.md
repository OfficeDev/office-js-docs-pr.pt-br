---
title: Office. Context – conjunto de requisitos 1,7
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,7.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 4a1ca6b4975ffba2c2bd400267fbe7db63f88244
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570727"
---
# <a name="context-mailbox-requirement-set-17"></a><span data-ttu-id="9628c-103">contexto (conjunto de requisitos de caixa de correio 1,7)</span><span class="sxs-lookup"><span data-stu-id="9628c-103">context (Mailbox requirement set 1.7)</span></span>

### <a name="officecontext"></a><span data-ttu-id="9628c-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="9628c-104">[Office](office.md).context</span></span>

<span data-ttu-id="9628c-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="9628c-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="9628c-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="9628c-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9628c-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9628c-107">Requirements</span></span>

|<span data-ttu-id="9628c-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="9628c-108">Requirement</span></span>| <span data-ttu-id="9628c-109">Valor</span><span class="sxs-lookup"><span data-stu-id="9628c-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="9628c-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9628c-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9628c-111">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-111">1.1</span></span>|
|[<span data-ttu-id="9628c-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9628c-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9628c-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="9628c-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="9628c-114">Properties</span></span>

| <span data-ttu-id="9628c-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="9628c-115">Property</span></span> | <span data-ttu-id="9628c-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="9628c-116">Modes</span></span> | <span data-ttu-id="9628c-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="9628c-117">Return type</span></span> | <span data-ttu-id="9628c-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="9628c-118">Minimum</span></span><br><span data-ttu-id="9628c-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="9628c-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9628c-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="9628c-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="9628c-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="9628c-121">Compose</span></span><br><span data-ttu-id="9628c-122">Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-122">Read</span></span> | <span data-ttu-id="9628c-123">String</span><span class="sxs-lookup"><span data-stu-id="9628c-123">String</span></span> | [<span data-ttu-id="9628c-124">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9628c-125">la</span><span class="sxs-lookup"><span data-stu-id="9628c-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="9628c-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="9628c-126">Compose</span></span><br><span data-ttu-id="9628c-127">Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-127">Read</span></span> | [<span data-ttu-id="9628c-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="9628c-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="9628c-129">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9628c-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="9628c-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="9628c-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="9628c-131">Compose</span></span><br><span data-ttu-id="9628c-132">Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-132">Read</span></span> | <span data-ttu-id="9628c-133">String</span><span class="sxs-lookup"><span data-stu-id="9628c-133">String</span></span> | [<span data-ttu-id="9628c-134">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9628c-135">principal</span><span class="sxs-lookup"><span data-stu-id="9628c-135">host</span></span>](#host-hosttype) | <span data-ttu-id="9628c-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="9628c-136">Compose</span></span><br><span data-ttu-id="9628c-137">Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-137">Read</span></span> | [<span data-ttu-id="9628c-138">HostType</span><span class="sxs-lookup"><span data-stu-id="9628c-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="9628c-139">1,5</span><span class="sxs-lookup"><span data-stu-id="9628c-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="9628c-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="9628c-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="9628c-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="9628c-141">Compose</span></span><br><span data-ttu-id="9628c-142">Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-142">Read</span></span> | [<span data-ttu-id="9628c-143">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="9628c-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="9628c-144">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9628c-145">plataforma</span><span class="sxs-lookup"><span data-stu-id="9628c-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="9628c-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="9628c-146">Compose</span></span><br><span data-ttu-id="9628c-147">Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-147">Read</span></span> | [<span data-ttu-id="9628c-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9628c-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="9628c-149">1,5</span><span class="sxs-lookup"><span data-stu-id="9628c-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="9628c-150">atende</span><span class="sxs-lookup"><span data-stu-id="9628c-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="9628c-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="9628c-151">Compose</span></span><br><span data-ttu-id="9628c-152">Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-152">Read</span></span> | [<span data-ttu-id="9628c-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="9628c-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="9628c-154">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9628c-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="9628c-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="9628c-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="9628c-156">Compose</span></span><br><span data-ttu-id="9628c-157">Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-157">Read</span></span> | [<span data-ttu-id="9628c-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9628c-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="9628c-159">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9628c-160">ui</span><span class="sxs-lookup"><span data-stu-id="9628c-160">ui</span></span>](#ui-ui) | <span data-ttu-id="9628c-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="9628c-161">Compose</span></span><br><span data-ttu-id="9628c-162">Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-162">Read</span></span> | [<span data-ttu-id="9628c-163">UI</span><span class="sxs-lookup"><span data-stu-id="9628c-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="9628c-164">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="9628c-165">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="9628c-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="9628c-166">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9628c-166">contentLanguage: String</span></span>

<span data-ttu-id="9628c-167">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="9628c-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="9628c-168">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="9628c-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="9628c-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="9628c-169">Type</span></span>

*   <span data-ttu-id="9628c-170">String</span><span class="sxs-lookup"><span data-stu-id="9628c-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9628c-171">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9628c-171">Requirements</span></span>

|<span data-ttu-id="9628c-172">Requisito</span><span class="sxs-lookup"><span data-stu-id="9628c-172">Requirement</span></span>| <span data-ttu-id="9628c-173">Valor</span><span class="sxs-lookup"><span data-stu-id="9628c-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="9628c-174">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9628c-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9628c-175">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-175">1.1</span></span>|
|[<span data-ttu-id="9628c-176">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9628c-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9628c-177">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9628c-178">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9628c-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="9628c-179">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="9628c-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="9628c-180">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="9628c-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="9628c-181">Tipo</span><span class="sxs-lookup"><span data-stu-id="9628c-181">Type</span></span>

*   [<span data-ttu-id="9628c-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="9628c-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="9628c-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9628c-183">Requirements</span></span>

|<span data-ttu-id="9628c-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="9628c-184">Requirement</span></span>| <span data-ttu-id="9628c-185">Valor</span><span class="sxs-lookup"><span data-stu-id="9628c-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="9628c-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9628c-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9628c-187">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-187">1.1</span></span>|
|[<span data-ttu-id="9628c-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9628c-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9628c-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9628c-190">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9628c-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="9628c-191">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9628c-191">displayLanguage: String</span></span>

<span data-ttu-id="9628c-192">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="9628c-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="9628c-193">O `displayLanguage` valor reflete a configuração de **idioma de exibição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="9628c-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="9628c-194">Tipo</span><span class="sxs-lookup"><span data-stu-id="9628c-194">Type</span></span>

*   <span data-ttu-id="9628c-195">String</span><span class="sxs-lookup"><span data-stu-id="9628c-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9628c-196">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9628c-196">Requirements</span></span>

|<span data-ttu-id="9628c-197">Requisito</span><span class="sxs-lookup"><span data-stu-id="9628c-197">Requirement</span></span>| <span data-ttu-id="9628c-198">Valor</span><span class="sxs-lookup"><span data-stu-id="9628c-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="9628c-199">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9628c-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9628c-200">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-200">1.1</span></span>|
|[<span data-ttu-id="9628c-201">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9628c-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9628c-202">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9628c-203">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9628c-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="9628c-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="9628c-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="9628c-205">Obtém o aplicativo do Office que está hospedando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="9628c-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="9628c-206">Como alternativa, você pode usar a propriedade [Office. Context. Diagnostics](#diagnostics-contextinformation) para obter o host.</span><span class="sxs-lookup"><span data-stu-id="9628c-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="9628c-207">Tipo</span><span class="sxs-lookup"><span data-stu-id="9628c-207">Type</span></span>

*   [<span data-ttu-id="9628c-208">HostType</span><span class="sxs-lookup"><span data-stu-id="9628c-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="9628c-209">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9628c-209">Requirements</span></span>

|<span data-ttu-id="9628c-210">Requisito</span><span class="sxs-lookup"><span data-stu-id="9628c-210">Requirement</span></span>| <span data-ttu-id="9628c-211">Valor</span><span class="sxs-lookup"><span data-stu-id="9628c-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="9628c-212">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9628c-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9628c-213">1,5</span><span class="sxs-lookup"><span data-stu-id="9628c-213">1.5</span></span>|
|[<span data-ttu-id="9628c-214">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9628c-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9628c-215">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9628c-216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9628c-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="9628c-217">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="9628c-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="9628c-218">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="9628c-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="9628c-219">Como alternativa, você pode usar a propriedade [Office. Context. Diagnostics](#diagnostics-contextinformation) para obter a plataforma.</span><span class="sxs-lookup"><span data-stu-id="9628c-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="9628c-220">Tipo</span><span class="sxs-lookup"><span data-stu-id="9628c-220">Type</span></span>

*   [<span data-ttu-id="9628c-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9628c-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="9628c-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9628c-222">Requirements</span></span>

|<span data-ttu-id="9628c-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="9628c-223">Requirement</span></span>| <span data-ttu-id="9628c-224">Valor</span><span class="sxs-lookup"><span data-stu-id="9628c-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="9628c-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9628c-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9628c-226">1,5</span><span class="sxs-lookup"><span data-stu-id="9628c-226">1.5</span></span>|
|[<span data-ttu-id="9628c-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9628c-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9628c-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9628c-229">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9628c-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="9628c-230">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="9628c-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="9628c-231">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o aplicativo atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="9628c-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="9628c-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="9628c-232">Type</span></span>

*   [<span data-ttu-id="9628c-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="9628c-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="9628c-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9628c-234">Requirements</span></span>

|<span data-ttu-id="9628c-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="9628c-235">Requirement</span></span>| <span data-ttu-id="9628c-236">Valor</span><span class="sxs-lookup"><span data-stu-id="9628c-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="9628c-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9628c-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9628c-238">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-238">1.1</span></span>|
|[<span data-ttu-id="9628c-239">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9628c-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9628c-240">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9628c-241">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9628c-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="9628c-242">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="9628c-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="9628c-243">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="9628c-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="9628c-244">O `RoamingSettings` objeto permite armazenar e acessar dados de um suplemento de email armazenado na caixa de correio de um usuário, para que esteja disponível para esse suplemento quando ele estiver sendo executado a partir de qualquer cliente do Outlook usado para acessar a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="9628c-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="9628c-245">Tipo</span><span class="sxs-lookup"><span data-stu-id="9628c-245">Type</span></span>

*   [<span data-ttu-id="9628c-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9628c-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="9628c-247">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9628c-247">Requirements</span></span>

|<span data-ttu-id="9628c-248">Requisito</span><span class="sxs-lookup"><span data-stu-id="9628c-248">Requirement</span></span>| <span data-ttu-id="9628c-249">Valor</span><span class="sxs-lookup"><span data-stu-id="9628c-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="9628c-250">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9628c-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9628c-251">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-251">1.1</span></span>|
|[<span data-ttu-id="9628c-252">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9628c-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="9628c-253">Restrito</span><span class="sxs-lookup"><span data-stu-id="9628c-253">Restricted</span></span>|
|[<span data-ttu-id="9628c-254">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9628c-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9628c-255">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="9628c-256">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="9628c-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="9628c-257">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="9628c-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="9628c-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="9628c-258">Type</span></span>

*   [<span data-ttu-id="9628c-259">UI</span><span class="sxs-lookup"><span data-stu-id="9628c-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="9628c-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9628c-260">Requirements</span></span>

|<span data-ttu-id="9628c-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="9628c-261">Requirement</span></span>| <span data-ttu-id="9628c-262">Valor</span><span class="sxs-lookup"><span data-stu-id="9628c-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="9628c-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9628c-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9628c-264">1.1</span><span class="sxs-lookup"><span data-stu-id="9628c-264">1.1</span></span>|
|[<span data-ttu-id="9628c-265">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9628c-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9628c-266">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="9628c-266">Compose or Read</span></span>|
