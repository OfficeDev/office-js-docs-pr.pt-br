---
title: Office. Context – conjunto de requisitos 1,3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: adea7414da7242ba3d2f9d57210c934d61e146df
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163763"
---
# <a name="context"></a><span data-ttu-id="3dbce-102">context</span><span class="sxs-lookup"><span data-stu-id="3dbce-102">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="3dbce-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="3dbce-103">[Office](office.md).context</span></span>

<span data-ttu-id="3dbce-104">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="3dbce-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="3dbce-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.3).</span><span class="sxs-lookup"><span data-stu-id="3dbce-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dbce-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dbce-106">Requirements</span></span>

|<span data-ttu-id="3dbce-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dbce-107">Requirement</span></span>| <span data-ttu-id="3dbce-108">Valor</span><span class="sxs-lookup"><span data-stu-id="3dbce-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dbce-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dbce-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dbce-110">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-110">1.1</span></span>|
|[<span data-ttu-id="3dbce-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dbce-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dbce-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="3dbce-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="3dbce-113">Properties</span></span>

| <span data-ttu-id="3dbce-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="3dbce-114">Property</span></span> | <span data-ttu-id="3dbce-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="3dbce-115">Modes</span></span> | <span data-ttu-id="3dbce-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="3dbce-116">Return type</span></span> | <span data-ttu-id="3dbce-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="3dbce-117">Minimum</span></span><br><span data-ttu-id="3dbce-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="3dbce-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3dbce-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="3dbce-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="3dbce-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="3dbce-120">Compose</span></span><br><span data-ttu-id="3dbce-121">Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-121">Read</span></span> | <span data-ttu-id="3dbce-122">String</span><span class="sxs-lookup"><span data-stu-id="3dbce-122">String</span></span> | [<span data-ttu-id="3dbce-123">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dbce-124">la</span><span class="sxs-lookup"><span data-stu-id="3dbce-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="3dbce-125">Escrever</span><span class="sxs-lookup"><span data-stu-id="3dbce-125">Compose</span></span><br><span data-ttu-id="3dbce-126">Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-126">Read</span></span> | [<span data-ttu-id="3dbce-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="3dbce-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3) | [<span data-ttu-id="3dbce-128">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dbce-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="3dbce-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="3dbce-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="3dbce-130">Compose</span></span><br><span data-ttu-id="3dbce-131">Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-131">Read</span></span> | <span data-ttu-id="3dbce-132">String</span><span class="sxs-lookup"><span data-stu-id="3dbce-132">String</span></span> | [<span data-ttu-id="3dbce-133">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dbce-134">principal</span><span class="sxs-lookup"><span data-stu-id="3dbce-134">host</span></span>](#host-hosttype) | <span data-ttu-id="3dbce-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="3dbce-135">Compose</span></span><br><span data-ttu-id="3dbce-136">Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-136">Read</span></span> | [<span data-ttu-id="3dbce-137">HostType</span><span class="sxs-lookup"><span data-stu-id="3dbce-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.3) | [<span data-ttu-id="3dbce-138">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dbce-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="3dbce-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="3dbce-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="3dbce-140">Compose</span></span><br><span data-ttu-id="3dbce-141">Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-141">Read</span></span> | [<span data-ttu-id="3dbce-142">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="3dbce-142">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3) | [<span data-ttu-id="3dbce-143">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dbce-144">plataforma</span><span class="sxs-lookup"><span data-stu-id="3dbce-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="3dbce-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="3dbce-145">Compose</span></span><br><span data-ttu-id="3dbce-146">Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-146">Read</span></span> | [<span data-ttu-id="3dbce-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="3dbce-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.3) | [<span data-ttu-id="3dbce-148">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dbce-149">atende</span><span class="sxs-lookup"><span data-stu-id="3dbce-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="3dbce-150">Escrever</span><span class="sxs-lookup"><span data-stu-id="3dbce-150">Compose</span></span><br><span data-ttu-id="3dbce-151">Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-151">Read</span></span> | [<span data-ttu-id="3dbce-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="3dbce-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3) | [<span data-ttu-id="3dbce-153">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dbce-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="3dbce-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="3dbce-155">Escrever</span><span class="sxs-lookup"><span data-stu-id="3dbce-155">Compose</span></span><br><span data-ttu-id="3dbce-156">Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-156">Read</span></span> | [<span data-ttu-id="3dbce-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="3dbce-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3) | [<span data-ttu-id="3dbce-158">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dbce-159">ui</span><span class="sxs-lookup"><span data-stu-id="3dbce-159">ui</span></span>](#ui-ui) | <span data-ttu-id="3dbce-160">Escrever</span><span class="sxs-lookup"><span data-stu-id="3dbce-160">Compose</span></span><br><span data-ttu-id="3dbce-161">Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-161">Read</span></span> | [<span data-ttu-id="3dbce-162">UI</span><span class="sxs-lookup"><span data-stu-id="3dbce-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3) | [<span data-ttu-id="3dbce-163">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="3dbce-164">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="3dbce-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="3dbce-165">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3dbce-165">contentLanguage: String</span></span>

<span data-ttu-id="3dbce-166">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="3dbce-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="3dbce-167">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="3dbce-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="3dbce-168">Tipo</span><span class="sxs-lookup"><span data-stu-id="3dbce-168">Type</span></span>

*   <span data-ttu-id="3dbce-169">String</span><span class="sxs-lookup"><span data-stu-id="3dbce-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dbce-170">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dbce-170">Requirements</span></span>

|<span data-ttu-id="3dbce-171">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dbce-171">Requirement</span></span>| <span data-ttu-id="3dbce-172">Valor</span><span class="sxs-lookup"><span data-stu-id="3dbce-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dbce-173">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dbce-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dbce-174">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-174">1.1</span></span>|
|[<span data-ttu-id="3dbce-175">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dbce-175">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dbce-176">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dbce-177">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3dbce-177">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="3dbce-178">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="3dbce-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="3dbce-179">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="3dbce-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="3dbce-180">Tipo</span><span class="sxs-lookup"><span data-stu-id="3dbce-180">Type</span></span>

*   [<span data-ttu-id="3dbce-181">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="3dbce-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="3dbce-182">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dbce-182">Requirements</span></span>

|<span data-ttu-id="3dbce-183">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dbce-183">Requirement</span></span>| <span data-ttu-id="3dbce-184">Valor</span><span class="sxs-lookup"><span data-stu-id="3dbce-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dbce-185">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dbce-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dbce-186">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-186">1.1</span></span>|
|[<span data-ttu-id="3dbce-187">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dbce-187">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dbce-188">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dbce-189">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3dbce-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="3dbce-190">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3dbce-190">displayLanguage: String</span></span>

<span data-ttu-id="3dbce-191">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="3dbce-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="3dbce-192">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="3dbce-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="3dbce-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="3dbce-193">Type</span></span>

*   <span data-ttu-id="3dbce-194">String</span><span class="sxs-lookup"><span data-stu-id="3dbce-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dbce-195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dbce-195">Requirements</span></span>

|<span data-ttu-id="3dbce-196">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dbce-196">Requirement</span></span>| <span data-ttu-id="3dbce-197">Valor</span><span class="sxs-lookup"><span data-stu-id="3dbce-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dbce-198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dbce-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dbce-199">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-199">1.1</span></span>|
|[<span data-ttu-id="3dbce-200">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dbce-200">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dbce-201">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dbce-202">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3dbce-202">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="3dbce-203">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="3dbce-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="3dbce-204">Obtém o host do aplicativo do Office no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="3dbce-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="3dbce-205">Tipo</span><span class="sxs-lookup"><span data-stu-id="3dbce-205">Type</span></span>

*   [<span data-ttu-id="3dbce-206">HostType</span><span class="sxs-lookup"><span data-stu-id="3dbce-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="3dbce-207">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dbce-207">Requirements</span></span>

|<span data-ttu-id="3dbce-208">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dbce-208">Requirement</span></span>| <span data-ttu-id="3dbce-209">Valor</span><span class="sxs-lookup"><span data-stu-id="3dbce-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dbce-210">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dbce-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dbce-211">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-211">1.1</span></span>|
|[<span data-ttu-id="3dbce-212">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dbce-212">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dbce-213">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dbce-214">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3dbce-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="3dbce-215">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="3dbce-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="3dbce-216">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="3dbce-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="3dbce-217">Tipo</span><span class="sxs-lookup"><span data-stu-id="3dbce-217">Type</span></span>

*   [<span data-ttu-id="3dbce-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="3dbce-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="3dbce-219">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dbce-219">Requirements</span></span>

|<span data-ttu-id="3dbce-220">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dbce-220">Requirement</span></span>| <span data-ttu-id="3dbce-221">Valor</span><span class="sxs-lookup"><span data-stu-id="3dbce-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dbce-222">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dbce-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dbce-223">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-223">1.1</span></span>|
|[<span data-ttu-id="3dbce-224">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dbce-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dbce-225">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dbce-226">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3dbce-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="3dbce-227">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="3dbce-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="3dbce-228">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o host atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="3dbce-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="3dbce-229">Tipo</span><span class="sxs-lookup"><span data-stu-id="3dbce-229">Type</span></span>

*   [<span data-ttu-id="3dbce-230">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="3dbce-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="3dbce-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dbce-231">Requirements</span></span>

|<span data-ttu-id="3dbce-232">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dbce-232">Requirement</span></span>| <span data-ttu-id="3dbce-233">Valor</span><span class="sxs-lookup"><span data-stu-id="3dbce-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dbce-234">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dbce-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dbce-235">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-235">1.1</span></span>|
|[<span data-ttu-id="3dbce-236">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dbce-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dbce-237">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dbce-238">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3dbce-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="3dbce-239">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="3dbce-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="3dbce-240">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="3dbce-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="3dbce-241">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="3dbce-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="3dbce-242">Tipo</span><span class="sxs-lookup"><span data-stu-id="3dbce-242">Type</span></span>

*   [<span data-ttu-id="3dbce-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="3dbce-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="3dbce-244">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dbce-244">Requirements</span></span>

|<span data-ttu-id="3dbce-245">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dbce-245">Requirement</span></span>| <span data-ttu-id="3dbce-246">Valor</span><span class="sxs-lookup"><span data-stu-id="3dbce-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dbce-247">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dbce-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dbce-248">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-248">1.1</span></span>|
|[<span data-ttu-id="3dbce-249">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3dbce-249">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="3dbce-250">Restrito</span><span class="sxs-lookup"><span data-stu-id="3dbce-250">Restricted</span></span>|
|[<span data-ttu-id="3dbce-251">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dbce-251">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dbce-252">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="3dbce-253">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="3dbce-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="3dbce-254">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="3dbce-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="3dbce-255">Tipo</span><span class="sxs-lookup"><span data-stu-id="3dbce-255">Type</span></span>

*   [<span data-ttu-id="3dbce-256">UI</span><span class="sxs-lookup"><span data-stu-id="3dbce-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="3dbce-257">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3dbce-257">Requirements</span></span>

|<span data-ttu-id="3dbce-258">Requisito</span><span class="sxs-lookup"><span data-stu-id="3dbce-258">Requirement</span></span>| <span data-ttu-id="3dbce-259">Valor</span><span class="sxs-lookup"><span data-stu-id="3dbce-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dbce-260">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3dbce-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dbce-261">1.1</span><span class="sxs-lookup"><span data-stu-id="3dbce-261">1.1</span></span>|
|[<span data-ttu-id="3dbce-262">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3dbce-262">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dbce-263">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3dbce-263">Compose or Read</span></span>|
