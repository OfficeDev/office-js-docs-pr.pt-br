---
title: Office.context – conjunto de requisitos 1.6
description: Office. Membros do objeto Context disponíveis para Outlook de entrada usando o conjunto de requisitos da API de Caixa de Correio 1.6.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: d4c65cea9b581665e0dc7b38a8e0bf10d6b544f9
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590999"
---
# <a name="context-mailbox-requirement-set-16"></a><span data-ttu-id="c5ab7-103">context (Conjunto de requisitos de caixa de correio 1.6)</span><span class="sxs-lookup"><span data-stu-id="c5ab7-103">context (Mailbox requirement set 1.6)</span></span>

### <a name="officecontext"></a><span data-ttu-id="c5ab7-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="c5ab7-104">[Office](office.md).context</span></span>

<span data-ttu-id="c5ab7-105">Office.context fornece interfaces compartilhadas que são usadas por complementos em todos os Office aplicativos.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="c5ab7-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="c5ab7-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5ab7-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5ab7-107">Requirements</span></span>

|<span data-ttu-id="c5ab7-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5ab7-108">Requirement</span></span>| <span data-ttu-id="c5ab7-109">Valor</span><span class="sxs-lookup"><span data-stu-id="c5ab7-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5ab7-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5ab7-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5ab7-111">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-111">1.1</span></span>|
|[<span data-ttu-id="c5ab7-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5ab7-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5ab7-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="c5ab7-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="c5ab7-114">Properties</span></span>

| <span data-ttu-id="c5ab7-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="c5ab7-115">Property</span></span> | <span data-ttu-id="c5ab7-116">Modos</span><span class="sxs-lookup"><span data-stu-id="c5ab7-116">Modes</span></span> | <span data-ttu-id="c5ab7-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="c5ab7-117">Return type</span></span> | <span data-ttu-id="c5ab7-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="c5ab7-118">Minimum</span></span><br><span data-ttu-id="c5ab7-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="c5ab7-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c5ab7-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="c5ab7-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="c5ab7-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="c5ab7-121">Compose</span></span><br><span data-ttu-id="c5ab7-122">Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-122">Read</span></span> | <span data-ttu-id="c5ab7-123">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c5ab7-123">String</span></span> | [<span data-ttu-id="c5ab7-124">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5ab7-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="c5ab7-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="c5ab7-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="c5ab7-126">Compose</span></span><br><span data-ttu-id="c5ab7-127">Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-127">Read</span></span> | [<span data-ttu-id="c5ab7-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="c5ab7-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c5ab7-129">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5ab7-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="c5ab7-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="c5ab7-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="c5ab7-131">Compose</span></span><br><span data-ttu-id="c5ab7-132">Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-132">Read</span></span> | <span data-ttu-id="c5ab7-133">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c5ab7-133">String</span></span> | [<span data-ttu-id="c5ab7-134">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5ab7-135">host</span><span class="sxs-lookup"><span data-stu-id="c5ab7-135">host</span></span>](#host-hosttype) | <span data-ttu-id="c5ab7-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="c5ab7-136">Compose</span></span><br><span data-ttu-id="c5ab7-137">Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-137">Read</span></span> | [<span data-ttu-id="c5ab7-138">HostType</span><span class="sxs-lookup"><span data-stu-id="c5ab7-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c5ab7-139">1.5</span><span class="sxs-lookup"><span data-stu-id="c5ab7-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="c5ab7-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="c5ab7-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="c5ab7-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="c5ab7-141">Compose</span></span><br><span data-ttu-id="c5ab7-142">Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-142">Read</span></span> | [<span data-ttu-id="c5ab7-143">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="c5ab7-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c5ab7-144">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5ab7-145">platform</span><span class="sxs-lookup"><span data-stu-id="c5ab7-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="c5ab7-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="c5ab7-146">Compose</span></span><br><span data-ttu-id="c5ab7-147">Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-147">Read</span></span> | [<span data-ttu-id="c5ab7-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="c5ab7-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c5ab7-149">1.5</span><span class="sxs-lookup"><span data-stu-id="c5ab7-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="c5ab7-150">requirements</span><span class="sxs-lookup"><span data-stu-id="c5ab7-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="c5ab7-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="c5ab7-151">Compose</span></span><br><span data-ttu-id="c5ab7-152">Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-152">Read</span></span> | [<span data-ttu-id="c5ab7-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="c5ab7-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c5ab7-154">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5ab7-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="c5ab7-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="c5ab7-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="c5ab7-156">Compose</span></span><br><span data-ttu-id="c5ab7-157">Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-157">Read</span></span> | [<span data-ttu-id="c5ab7-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c5ab7-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c5ab7-159">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5ab7-160">ui</span><span class="sxs-lookup"><span data-stu-id="c5ab7-160">ui</span></span>](#ui-ui) | <span data-ttu-id="c5ab7-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="c5ab7-161">Compose</span></span><br><span data-ttu-id="c5ab7-162">Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-162">Read</span></span> | [<span data-ttu-id="c5ab7-163">UI</span><span class="sxs-lookup"><span data-stu-id="c5ab7-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c5ab7-164">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="c5ab7-165">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="c5ab7-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="c5ab7-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="c5ab7-166">contentLanguage: String</span></span>

<span data-ttu-id="c5ab7-167">Obtém a localidade (idioma) especificada pelo usuário para editar o item.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="c5ab7-168">O `contentLanguage` valor reflete a **configuração** atual de Idioma de Edição especificada com opções de > de arquivo **> idioma** no aplicativo Office cliente.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="c5ab7-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-169">Type</span></span>

*   <span data-ttu-id="c5ab7-170">String</span><span class="sxs-lookup"><span data-stu-id="c5ab7-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5ab7-171">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5ab7-171">Requirements</span></span>

|<span data-ttu-id="c5ab7-172">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5ab7-172">Requirement</span></span>| <span data-ttu-id="c5ab7-173">Valor</span><span class="sxs-lookup"><span data-stu-id="c5ab7-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5ab7-174">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5ab7-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5ab7-175">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-175">1.1</span></span>|
|[<span data-ttu-id="c5ab7-176">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5ab7-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5ab7-177">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5ab7-178">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="c5ab7-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="c5ab7-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="c5ab7-180">Obtém informações sobre o ambiente no qual o complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c5ab7-181">Tipo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-181">Type</span></span>

*   [<span data-ttu-id="c5ab7-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="c5ab7-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="c5ab7-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5ab7-183">Requirements</span></span>

|<span data-ttu-id="c5ab7-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5ab7-184">Requirement</span></span>| <span data-ttu-id="c5ab7-185">Valor</span><span class="sxs-lookup"><span data-stu-id="c5ab7-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5ab7-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5ab7-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5ab7-187">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-187">1.1</span></span>|
|[<span data-ttu-id="c5ab7-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5ab7-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5ab7-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5ab7-190">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="c5ab7-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="c5ab7-191">displayLanguage: String</span></span>

<span data-ttu-id="c5ab7-192">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente Office cliente.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="c5ab7-193">O valor reflete a configuração atual de Idioma de Exibição especificada com Opções > > Idioma no aplicativo Office `displayLanguage` cliente.  </span><span class="sxs-lookup"><span data-stu-id="c5ab7-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="c5ab7-194">Tipo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-194">Type</span></span>

*   <span data-ttu-id="c5ab7-195">String</span><span class="sxs-lookup"><span data-stu-id="c5ab7-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5ab7-196">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5ab7-196">Requirements</span></span>

|<span data-ttu-id="c5ab7-197">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5ab7-197">Requirement</span></span>| <span data-ttu-id="c5ab7-198">Valor</span><span class="sxs-lookup"><span data-stu-id="c5ab7-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5ab7-199">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5ab7-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5ab7-200">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-200">1.1</span></span>|
|[<span data-ttu-id="c5ab7-201">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5ab7-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5ab7-202">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5ab7-203">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="c5ab7-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="c5ab7-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="c5ab7-205">Obtém o Office aplicativo que está hospedando o complemento.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="c5ab7-206">Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter o host.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="c5ab7-207">Tipo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-207">Type</span></span>

*   [<span data-ttu-id="c5ab7-208">HostType</span><span class="sxs-lookup"><span data-stu-id="c5ab7-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="c5ab7-209">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5ab7-209">Requirements</span></span>

|<span data-ttu-id="c5ab7-210">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5ab7-210">Requirement</span></span>| <span data-ttu-id="c5ab7-211">Valor</span><span class="sxs-lookup"><span data-stu-id="c5ab7-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5ab7-212">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5ab7-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5ab7-213">1,5</span><span class="sxs-lookup"><span data-stu-id="c5ab7-213">1.5</span></span>|
|[<span data-ttu-id="c5ab7-214">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5ab7-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5ab7-215">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5ab7-216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="c5ab7-217">plataforma: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="c5ab7-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="c5ab7-218">Fornece a plataforma na qual o complemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="c5ab7-219">Como alternativa, você pode usar a [propriedade Office.context.diagnostics](#diagnostics-contextinformation) para obter a plataforma.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="c5ab7-220">Tipo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-220">Type</span></span>

*   [<span data-ttu-id="c5ab7-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="c5ab7-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="c5ab7-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5ab7-222">Requirements</span></span>

|<span data-ttu-id="c5ab7-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5ab7-223">Requirement</span></span>| <span data-ttu-id="c5ab7-224">Valor</span><span class="sxs-lookup"><span data-stu-id="c5ab7-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5ab7-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5ab7-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5ab7-226">1,5</span><span class="sxs-lookup"><span data-stu-id="c5ab7-226">1.5</span></span>|
|[<span data-ttu-id="c5ab7-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5ab7-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5ab7-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5ab7-229">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="c5ab7-230">requirements: [RequirementsSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="c5ab7-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="c5ab7-231">Fornece um método para determinar quais conjuntos de requisitos são suportados no aplicativo e na plataforma atual.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="c5ab7-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-232">Type</span></span>

*   [<span data-ttu-id="c5ab7-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="c5ab7-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="c5ab7-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5ab7-234">Requirements</span></span>

|<span data-ttu-id="c5ab7-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5ab7-235">Requirement</span></span>| <span data-ttu-id="c5ab7-236">Valor</span><span class="sxs-lookup"><span data-stu-id="c5ab7-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5ab7-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5ab7-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5ab7-238">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-238">1.1</span></span>|
|[<span data-ttu-id="c5ab7-239">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5ab7-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5ab7-240">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5ab7-241">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="c5ab7-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="c5ab7-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="c5ab7-243">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="c5ab7-244">O objeto permite que você armazene e acesse dados para um complemento de email armazenado na caixa de correio de um usuário, de modo que está disponível para esse complemento quando ele está sendo executado de qualquer cliente Outlook usado para acessar essa caixa de `RoamingSettings` correio.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c5ab7-245">Tipo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-245">Type</span></span>

*   [<span data-ttu-id="c5ab7-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c5ab7-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="c5ab7-247">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5ab7-247">Requirements</span></span>

|<span data-ttu-id="c5ab7-248">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5ab7-248">Requirement</span></span>| <span data-ttu-id="c5ab7-249">Valor</span><span class="sxs-lookup"><span data-stu-id="c5ab7-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5ab7-250">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5ab7-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5ab7-251">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-251">1.1</span></span>|
|[<span data-ttu-id="c5ab7-252">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="c5ab7-253">Restrito</span><span class="sxs-lookup"><span data-stu-id="c5ab7-253">Restricted</span></span>|
|[<span data-ttu-id="c5ab7-254">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5ab7-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5ab7-255">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="c5ab7-256">interface do usuário: [interface do usuário](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="c5ab7-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="c5ab7-257">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus Office de usuário.</span><span class="sxs-lookup"><span data-stu-id="c5ab7-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c5ab7-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="c5ab7-258">Type</span></span>

*   [<span data-ttu-id="c5ab7-259">UI</span><span class="sxs-lookup"><span data-stu-id="c5ab7-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="c5ab7-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5ab7-260">Requirements</span></span>

|<span data-ttu-id="c5ab7-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5ab7-261">Requirement</span></span>| <span data-ttu-id="c5ab7-262">Valor</span><span class="sxs-lookup"><span data-stu-id="c5ab7-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5ab7-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5ab7-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5ab7-264">1.1</span><span class="sxs-lookup"><span data-stu-id="c5ab7-264">1.1</span></span>|
|[<span data-ttu-id="c5ab7-265">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5ab7-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5ab7-266">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5ab7-266">Compose or Read</span></span>|
