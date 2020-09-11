---
title: Office. Context – conjunto de requisitos 1,1
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,1.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 2474c5f31dcd996785f61f06528ffcf3a92b05c2
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430559"
---
# <a name="context-mailbox-requirement-set-11"></a><span data-ttu-id="a4ad3-103">contexto (conjunto de requisitos de caixa de correio 1,1)</span><span class="sxs-lookup"><span data-stu-id="a4ad3-103">context (Mailbox requirement set 1.1)</span></span>

### <a name="officecontext"></a><span data-ttu-id="a4ad3-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="a4ad3-104">[Office](office.md).context</span></span>

<span data-ttu-id="a4ad3-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="a4ad3-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="a4ad3-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="a4ad3-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4ad3-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a4ad3-107">Requirements</span></span>

|<span data-ttu-id="a4ad3-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="a4ad3-108">Requirement</span></span>| <span data-ttu-id="a4ad3-109">Valor</span><span class="sxs-lookup"><span data-stu-id="a4ad3-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4ad3-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a4ad3-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a4ad3-111">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-111">1.1</span></span>|
|[<span data-ttu-id="a4ad3-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a4ad3-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a4ad3-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="a4ad3-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="a4ad3-114">Properties</span></span>

| <span data-ttu-id="a4ad3-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="a4ad3-115">Property</span></span> | <span data-ttu-id="a4ad3-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="a4ad3-116">Modes</span></span> | <span data-ttu-id="a4ad3-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="a4ad3-117">Return type</span></span> | <span data-ttu-id="a4ad3-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="a4ad3-118">Minimum</span></span><br><span data-ttu-id="a4ad3-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="a4ad3-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a4ad3-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="a4ad3-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="a4ad3-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="a4ad3-121">Compose</span></span><br><span data-ttu-id="a4ad3-122">Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-122">Read</span></span> | <span data-ttu-id="a4ad3-123">String</span><span class="sxs-lookup"><span data-stu-id="a4ad3-123">String</span></span> | [<span data-ttu-id="a4ad3-124">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a4ad3-125">la</span><span class="sxs-lookup"><span data-stu-id="a4ad3-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="a4ad3-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="a4ad3-126">Compose</span></span><br><span data-ttu-id="a4ad3-127">Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-127">Read</span></span> | [<span data-ttu-id="a4ad3-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="a4ad3-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="a4ad3-129">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a4ad3-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="a4ad3-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="a4ad3-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="a4ad3-131">Compose</span></span><br><span data-ttu-id="a4ad3-132">Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-132">Read</span></span> | <span data-ttu-id="a4ad3-133">String</span><span class="sxs-lookup"><span data-stu-id="a4ad3-133">String</span></span> | [<span data-ttu-id="a4ad3-134">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a4ad3-135">principal</span><span class="sxs-lookup"><span data-stu-id="a4ad3-135">host</span></span>](#host-hosttype) | <span data-ttu-id="a4ad3-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="a4ad3-136">Compose</span></span><br><span data-ttu-id="a4ad3-137">Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-137">Read</span></span> | [<span data-ttu-id="a4ad3-138">HostType</span><span class="sxs-lookup"><span data-stu-id="a4ad3-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="a4ad3-139">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a4ad3-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="a4ad3-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="a4ad3-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="a4ad3-141">Compose</span></span><br><span data-ttu-id="a4ad3-142">Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-142">Read</span></span> | [<span data-ttu-id="a4ad3-143">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="a4ad3-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="a4ad3-144">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a4ad3-145">plataforma</span><span class="sxs-lookup"><span data-stu-id="a4ad3-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="a4ad3-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="a4ad3-146">Compose</span></span><br><span data-ttu-id="a4ad3-147">Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-147">Read</span></span> | [<span data-ttu-id="a4ad3-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="a4ad3-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="a4ad3-149">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a4ad3-150">atende</span><span class="sxs-lookup"><span data-stu-id="a4ad3-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="a4ad3-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="a4ad3-151">Compose</span></span><br><span data-ttu-id="a4ad3-152">Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-152">Read</span></span> | [<span data-ttu-id="a4ad3-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="a4ad3-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="a4ad3-154">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a4ad3-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="a4ad3-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="a4ad3-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="a4ad3-156">Compose</span></span><br><span data-ttu-id="a4ad3-157">Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-157">Read</span></span> | [<span data-ttu-id="a4ad3-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a4ad3-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="a4ad3-159">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a4ad3-160">ui</span><span class="sxs-lookup"><span data-stu-id="a4ad3-160">ui</span></span>](#ui-ui) | <span data-ttu-id="a4ad3-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="a4ad3-161">Compose</span></span><br><span data-ttu-id="a4ad3-162">Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-162">Read</span></span> | [<span data-ttu-id="a4ad3-163">UI</span><span class="sxs-lookup"><span data-stu-id="a4ad3-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="a4ad3-164">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="a4ad3-165">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="a4ad3-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="a4ad3-166">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a4ad3-166">contentLanguage: String</span></span>

<span data-ttu-id="a4ad3-167">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="a4ad3-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="a4ad3-168">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="a4ad3-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="a4ad3-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-169">Type</span></span>

*   <span data-ttu-id="a4ad3-170">String</span><span class="sxs-lookup"><span data-stu-id="a4ad3-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4ad3-171">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a4ad3-171">Requirements</span></span>

|<span data-ttu-id="a4ad3-172">Requisito</span><span class="sxs-lookup"><span data-stu-id="a4ad3-172">Requirement</span></span>| <span data-ttu-id="a4ad3-173">Valor</span><span class="sxs-lookup"><span data-stu-id="a4ad3-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4ad3-174">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a4ad3-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a4ad3-175">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-175">1.1</span></span>|
|[<span data-ttu-id="a4ad3-176">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a4ad3-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a4ad3-177">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4ad3-178">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="a4ad3-179">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="a4ad3-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="a4ad3-180">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="a4ad3-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="a4ad3-181">Tipo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-181">Type</span></span>

*   [<span data-ttu-id="a4ad3-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="a4ad3-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="a4ad3-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a4ad3-183">Requirements</span></span>

|<span data-ttu-id="a4ad3-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="a4ad3-184">Requirement</span></span>| <span data-ttu-id="a4ad3-185">Valor</span><span class="sxs-lookup"><span data-stu-id="a4ad3-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4ad3-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a4ad3-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a4ad3-187">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-187">1.1</span></span>|
|[<span data-ttu-id="a4ad3-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a4ad3-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a4ad3-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4ad3-190">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="a4ad3-191">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a4ad3-191">displayLanguage: String</span></span>

<span data-ttu-id="a4ad3-192">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="a4ad3-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="a4ad3-193">O `displayLanguage` valor reflete a configuração de **idioma de exibição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="a4ad3-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="a4ad3-194">Tipo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-194">Type</span></span>

*   <span data-ttu-id="a4ad3-195">String</span><span class="sxs-lookup"><span data-stu-id="a4ad3-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4ad3-196">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a4ad3-196">Requirements</span></span>

|<span data-ttu-id="a4ad3-197">Requisito</span><span class="sxs-lookup"><span data-stu-id="a4ad3-197">Requirement</span></span>| <span data-ttu-id="a4ad3-198">Valor</span><span class="sxs-lookup"><span data-stu-id="a4ad3-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4ad3-199">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a4ad3-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a4ad3-200">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-200">1.1</span></span>|
|[<span data-ttu-id="a4ad3-201">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a4ad3-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a4ad3-202">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4ad3-203">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="a4ad3-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="a4ad3-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="a4ad3-205">Obtém o aplicativo do Office que está hospedando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="a4ad3-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="a4ad3-206">Tipo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-206">Type</span></span>

*   [<span data-ttu-id="a4ad3-207">HostType</span><span class="sxs-lookup"><span data-stu-id="a4ad3-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="a4ad3-208">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a4ad3-208">Requirements</span></span>

|<span data-ttu-id="a4ad3-209">Requisito</span><span class="sxs-lookup"><span data-stu-id="a4ad3-209">Requirement</span></span>| <span data-ttu-id="a4ad3-210">Valor</span><span class="sxs-lookup"><span data-stu-id="a4ad3-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4ad3-211">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a4ad3-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a4ad3-212">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-212">1.1</span></span>|
|[<span data-ttu-id="a4ad3-213">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a4ad3-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a4ad3-214">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4ad3-215">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="a4ad3-216">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="a4ad3-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="a4ad3-217">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="a4ad3-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="a4ad3-218">Tipo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-218">Type</span></span>

*   [<span data-ttu-id="a4ad3-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="a4ad3-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="a4ad3-220">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a4ad3-220">Requirements</span></span>

|<span data-ttu-id="a4ad3-221">Requisito</span><span class="sxs-lookup"><span data-stu-id="a4ad3-221">Requirement</span></span>| <span data-ttu-id="a4ad3-222">Valor</span><span class="sxs-lookup"><span data-stu-id="a4ad3-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4ad3-223">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a4ad3-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a4ad3-224">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-224">1.1</span></span>|
|[<span data-ttu-id="a4ad3-225">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a4ad3-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a4ad3-226">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4ad3-227">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="a4ad3-228">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="a4ad3-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="a4ad3-229">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o aplicativo atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="a4ad3-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="a4ad3-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-230">Type</span></span>

*   [<span data-ttu-id="a4ad3-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="a4ad3-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="a4ad3-232">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a4ad3-232">Requirements</span></span>

|<span data-ttu-id="a4ad3-233">Requisito</span><span class="sxs-lookup"><span data-stu-id="a4ad3-233">Requirement</span></span>| <span data-ttu-id="a4ad3-234">Valor</span><span class="sxs-lookup"><span data-stu-id="a4ad3-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4ad3-235">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a4ad3-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a4ad3-236">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-236">1.1</span></span>|
|[<span data-ttu-id="a4ad3-237">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a4ad3-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a4ad3-238">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4ad3-239">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="a4ad3-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="a4ad3-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="a4ad3-241">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="a4ad3-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="a4ad3-242">O `RoamingSettings` objeto permite armazenar e acessar dados de um suplemento de email armazenado na caixa de correio de um usuário, para que esteja disponível para esse suplemento quando ele estiver sendo executado a partir de qualquer cliente do Outlook usado para acessar a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="a4ad3-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="a4ad3-243">Tipo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-243">Type</span></span>

*   [<span data-ttu-id="a4ad3-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a4ad3-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="a4ad3-245">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a4ad3-245">Requirements</span></span>

|<span data-ttu-id="a4ad3-246">Requisito</span><span class="sxs-lookup"><span data-stu-id="a4ad3-246">Requirement</span></span>| <span data-ttu-id="a4ad3-247">Valor</span><span class="sxs-lookup"><span data-stu-id="a4ad3-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4ad3-248">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a4ad3-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a4ad3-249">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-249">1.1</span></span>|
|[<span data-ttu-id="a4ad3-250">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="a4ad3-251">Restrito</span><span class="sxs-lookup"><span data-stu-id="a4ad3-251">Restricted</span></span>|
|[<span data-ttu-id="a4ad3-252">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a4ad3-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a4ad3-253">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="a4ad3-254">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="a4ad3-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="a4ad3-255">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="a4ad3-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="a4ad3-256">Tipo</span><span class="sxs-lookup"><span data-stu-id="a4ad3-256">Type</span></span>

*   [<span data-ttu-id="a4ad3-257">UI</span><span class="sxs-lookup"><span data-stu-id="a4ad3-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="a4ad3-258">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a4ad3-258">Requirements</span></span>

|<span data-ttu-id="a4ad3-259">Requisito</span><span class="sxs-lookup"><span data-stu-id="a4ad3-259">Requirement</span></span>| <span data-ttu-id="a4ad3-260">Valor</span><span class="sxs-lookup"><span data-stu-id="a4ad3-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4ad3-261">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a4ad3-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a4ad3-262">1.1</span><span class="sxs-lookup"><span data-stu-id="a4ad3-262">1.1</span></span>|
|[<span data-ttu-id="a4ad3-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a4ad3-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a4ad3-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a4ad3-264">Compose or Read</span></span>|
