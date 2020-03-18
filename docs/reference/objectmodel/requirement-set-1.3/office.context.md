---
title: Office. Context – conjunto de requisitos 1,3
description: O modelo de objeto para o objeto Context do Outlook na API de suplementos do Outlook (versão da API da caixa de correio 1,3).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: dc524c5657380845744e241a600fbf66f1eed5ce
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720117"
---
# <a name="context"></a><span data-ttu-id="a67d5-103">context</span><span class="sxs-lookup"><span data-stu-id="a67d5-103">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="a67d5-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="a67d5-104">[Office](office.md).context</span></span>

<span data-ttu-id="a67d5-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="a67d5-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="a67d5-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.3).</span><span class="sxs-lookup"><span data-stu-id="a67d5-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a67d5-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a67d5-107">Requirements</span></span>

|<span data-ttu-id="a67d5-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="a67d5-108">Requirement</span></span>| <span data-ttu-id="a67d5-109">Valor</span><span class="sxs-lookup"><span data-stu-id="a67d5-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="a67d5-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a67d5-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a67d5-111">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-111">1.1</span></span>|
|[<span data-ttu-id="a67d5-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a67d5-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a67d5-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="a67d5-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="a67d5-114">Properties</span></span>

| <span data-ttu-id="a67d5-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="a67d5-115">Property</span></span> | <span data-ttu-id="a67d5-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="a67d5-116">Modes</span></span> | <span data-ttu-id="a67d5-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="a67d5-117">Return type</span></span> | <span data-ttu-id="a67d5-118">Mínimo</span><span class="sxs-lookup"><span data-stu-id="a67d5-118">Minimum</span></span><br><span data-ttu-id="a67d5-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="a67d5-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a67d5-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="a67d5-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="a67d5-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="a67d5-121">Compose</span></span><br><span data-ttu-id="a67d5-122">Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-122">Read</span></span> | <span data-ttu-id="a67d5-123">String</span><span class="sxs-lookup"><span data-stu-id="a67d5-123">String</span></span> | [<span data-ttu-id="a67d5-124">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a67d5-125">la</span><span class="sxs-lookup"><span data-stu-id="a67d5-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="a67d5-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="a67d5-126">Compose</span></span><br><span data-ttu-id="a67d5-127">Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-127">Read</span></span> | [<span data-ttu-id="a67d5-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="a67d5-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3) | [<span data-ttu-id="a67d5-129">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a67d5-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="a67d5-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="a67d5-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="a67d5-131">Compose</span></span><br><span data-ttu-id="a67d5-132">Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-132">Read</span></span> | <span data-ttu-id="a67d5-133">String</span><span class="sxs-lookup"><span data-stu-id="a67d5-133">String</span></span> | [<span data-ttu-id="a67d5-134">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a67d5-135">principal</span><span class="sxs-lookup"><span data-stu-id="a67d5-135">host</span></span>](#host-hosttype) | <span data-ttu-id="a67d5-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="a67d5-136">Compose</span></span><br><span data-ttu-id="a67d5-137">Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-137">Read</span></span> | [<span data-ttu-id="a67d5-138">HostType</span><span class="sxs-lookup"><span data-stu-id="a67d5-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.3) | [<span data-ttu-id="a67d5-139">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a67d5-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="a67d5-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="a67d5-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="a67d5-141">Compose</span></span><br><span data-ttu-id="a67d5-142">Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-142">Read</span></span> | [<span data-ttu-id="a67d5-143">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="a67d5-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3) | [<span data-ttu-id="a67d5-144">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a67d5-145">plataforma</span><span class="sxs-lookup"><span data-stu-id="a67d5-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="a67d5-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="a67d5-146">Compose</span></span><br><span data-ttu-id="a67d5-147">Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-147">Read</span></span> | [<span data-ttu-id="a67d5-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="a67d5-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.3) | [<span data-ttu-id="a67d5-149">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a67d5-150">atende</span><span class="sxs-lookup"><span data-stu-id="a67d5-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="a67d5-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="a67d5-151">Compose</span></span><br><span data-ttu-id="a67d5-152">Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-152">Read</span></span> | [<span data-ttu-id="a67d5-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="a67d5-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3) | [<span data-ttu-id="a67d5-154">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a67d5-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="a67d5-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="a67d5-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="a67d5-156">Compose</span></span><br><span data-ttu-id="a67d5-157">Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-157">Read</span></span> | [<span data-ttu-id="a67d5-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a67d5-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3) | [<span data-ttu-id="a67d5-159">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a67d5-160">ui</span><span class="sxs-lookup"><span data-stu-id="a67d5-160">ui</span></span>](#ui-ui) | <span data-ttu-id="a67d5-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="a67d5-161">Compose</span></span><br><span data-ttu-id="a67d5-162">Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-162">Read</span></span> | [<span data-ttu-id="a67d5-163">UI</span><span class="sxs-lookup"><span data-stu-id="a67d5-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3) | [<span data-ttu-id="a67d5-164">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="a67d5-165">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="a67d5-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="a67d5-166">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a67d5-166">contentLanguage: String</span></span>

<span data-ttu-id="a67d5-167">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="a67d5-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="a67d5-168">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="a67d5-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="a67d5-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="a67d5-169">Type</span></span>

*   <span data-ttu-id="a67d5-170">String</span><span class="sxs-lookup"><span data-stu-id="a67d5-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a67d5-171">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a67d5-171">Requirements</span></span>

|<span data-ttu-id="a67d5-172">Requisito</span><span class="sxs-lookup"><span data-stu-id="a67d5-172">Requirement</span></span>| <span data-ttu-id="a67d5-173">Valor</span><span class="sxs-lookup"><span data-stu-id="a67d5-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="a67d5-174">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a67d5-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a67d5-175">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-175">1.1</span></span>|
|[<span data-ttu-id="a67d5-176">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a67d5-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a67d5-177">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a67d5-178">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a67d5-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="a67d5-179">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="a67d5-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="a67d5-180">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="a67d5-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="a67d5-181">Tipo</span><span class="sxs-lookup"><span data-stu-id="a67d5-181">Type</span></span>

*   [<span data-ttu-id="a67d5-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="a67d5-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="a67d5-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a67d5-183">Requirements</span></span>

|<span data-ttu-id="a67d5-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="a67d5-184">Requirement</span></span>| <span data-ttu-id="a67d5-185">Valor</span><span class="sxs-lookup"><span data-stu-id="a67d5-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="a67d5-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a67d5-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a67d5-187">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-187">1.1</span></span>|
|[<span data-ttu-id="a67d5-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a67d5-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a67d5-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a67d5-190">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a67d5-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="a67d5-191">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a67d5-191">displayLanguage: String</span></span>

<span data-ttu-id="a67d5-192">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="a67d5-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="a67d5-193">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="a67d5-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="a67d5-194">Tipo</span><span class="sxs-lookup"><span data-stu-id="a67d5-194">Type</span></span>

*   <span data-ttu-id="a67d5-195">String</span><span class="sxs-lookup"><span data-stu-id="a67d5-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a67d5-196">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a67d5-196">Requirements</span></span>

|<span data-ttu-id="a67d5-197">Requisito</span><span class="sxs-lookup"><span data-stu-id="a67d5-197">Requirement</span></span>| <span data-ttu-id="a67d5-198">Valor</span><span class="sxs-lookup"><span data-stu-id="a67d5-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="a67d5-199">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a67d5-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a67d5-200">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-200">1.1</span></span>|
|[<span data-ttu-id="a67d5-201">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a67d5-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a67d5-202">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a67d5-203">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a67d5-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="a67d5-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="a67d5-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="a67d5-205">Obtém o host do aplicativo do Office no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="a67d5-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="a67d5-206">Tipo</span><span class="sxs-lookup"><span data-stu-id="a67d5-206">Type</span></span>

*   [<span data-ttu-id="a67d5-207">HostType</span><span class="sxs-lookup"><span data-stu-id="a67d5-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="a67d5-208">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a67d5-208">Requirements</span></span>

|<span data-ttu-id="a67d5-209">Requisito</span><span class="sxs-lookup"><span data-stu-id="a67d5-209">Requirement</span></span>| <span data-ttu-id="a67d5-210">Valor</span><span class="sxs-lookup"><span data-stu-id="a67d5-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="a67d5-211">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a67d5-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a67d5-212">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-212">1.1</span></span>|
|[<span data-ttu-id="a67d5-213">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a67d5-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a67d5-214">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a67d5-215">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a67d5-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="a67d5-216">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="a67d5-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="a67d5-217">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="a67d5-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="a67d5-218">Tipo</span><span class="sxs-lookup"><span data-stu-id="a67d5-218">Type</span></span>

*   [<span data-ttu-id="a67d5-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="a67d5-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="a67d5-220">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a67d5-220">Requirements</span></span>

|<span data-ttu-id="a67d5-221">Requisito</span><span class="sxs-lookup"><span data-stu-id="a67d5-221">Requirement</span></span>| <span data-ttu-id="a67d5-222">Valor</span><span class="sxs-lookup"><span data-stu-id="a67d5-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="a67d5-223">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a67d5-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a67d5-224">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-224">1.1</span></span>|
|[<span data-ttu-id="a67d5-225">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a67d5-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a67d5-226">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a67d5-227">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a67d5-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="a67d5-228">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="a67d5-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="a67d5-229">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o host atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="a67d5-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="a67d5-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="a67d5-230">Type</span></span>

*   [<span data-ttu-id="a67d5-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="a67d5-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="a67d5-232">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a67d5-232">Requirements</span></span>

|<span data-ttu-id="a67d5-233">Requisito</span><span class="sxs-lookup"><span data-stu-id="a67d5-233">Requirement</span></span>| <span data-ttu-id="a67d5-234">Valor</span><span class="sxs-lookup"><span data-stu-id="a67d5-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="a67d5-235">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a67d5-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a67d5-236">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-236">1.1</span></span>|
|[<span data-ttu-id="a67d5-237">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a67d5-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a67d5-238">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a67d5-239">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a67d5-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="a67d5-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="a67d5-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="a67d5-241">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="a67d5-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="a67d5-242">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="a67d5-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="a67d5-243">Tipo</span><span class="sxs-lookup"><span data-stu-id="a67d5-243">Type</span></span>

*   [<span data-ttu-id="a67d5-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a67d5-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="a67d5-245">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a67d5-245">Requirements</span></span>

|<span data-ttu-id="a67d5-246">Requisito</span><span class="sxs-lookup"><span data-stu-id="a67d5-246">Requirement</span></span>| <span data-ttu-id="a67d5-247">Valor</span><span class="sxs-lookup"><span data-stu-id="a67d5-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="a67d5-248">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a67d5-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a67d5-249">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-249">1.1</span></span>|
|[<span data-ttu-id="a67d5-250">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a67d5-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="a67d5-251">Restrito</span><span class="sxs-lookup"><span data-stu-id="a67d5-251">Restricted</span></span>|
|[<span data-ttu-id="a67d5-252">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a67d5-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a67d5-253">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="a67d5-254">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="a67d5-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="a67d5-255">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="a67d5-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="a67d5-256">Tipo</span><span class="sxs-lookup"><span data-stu-id="a67d5-256">Type</span></span>

*   [<span data-ttu-id="a67d5-257">UI</span><span class="sxs-lookup"><span data-stu-id="a67d5-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="a67d5-258">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a67d5-258">Requirements</span></span>

|<span data-ttu-id="a67d5-259">Requisito</span><span class="sxs-lookup"><span data-stu-id="a67d5-259">Requirement</span></span>| <span data-ttu-id="a67d5-260">Valor</span><span class="sxs-lookup"><span data-stu-id="a67d5-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="a67d5-261">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a67d5-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a67d5-262">1.1</span><span class="sxs-lookup"><span data-stu-id="a67d5-262">1.1</span></span>|
|[<span data-ttu-id="a67d5-263">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a67d5-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a67d5-264">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a67d5-264">Compose or Read</span></span>|
