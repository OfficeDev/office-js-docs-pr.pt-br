---
title: Office.context – conjunto de requisitos 1.6
description: Membros do objeto Office. Context disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,6.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 55e3761aea94d902903c53a9b3be687d94b42e12
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570755"
---
# <a name="context-mailbox-requirement-set-16"></a><span data-ttu-id="c0128-103">contexto (conjunto de requisitos de caixa de correio 1,6)</span><span class="sxs-lookup"><span data-stu-id="c0128-103">context (Mailbox requirement set 1.6)</span></span>

### <a name="officecontext"></a><span data-ttu-id="c0128-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="c0128-104">[Office](office.md).context</span></span>

<span data-ttu-id="c0128-105">O Office. Context fornece interfaces compartilhadas usadas por suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="c0128-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="c0128-106">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="c0128-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0128-107">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c0128-107">Requirements</span></span>

|<span data-ttu-id="c0128-108">Requisito</span><span class="sxs-lookup"><span data-stu-id="c0128-108">Requirement</span></span>| <span data-ttu-id="c0128-109">Valor</span><span class="sxs-lookup"><span data-stu-id="c0128-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0128-110">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c0128-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c0128-111">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-111">1.1</span></span>|
|[<span data-ttu-id="c0128-112">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c0128-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c0128-113">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="c0128-114">Propriedades</span><span class="sxs-lookup"><span data-stu-id="c0128-114">Properties</span></span>

| <span data-ttu-id="c0128-115">Propriedade</span><span class="sxs-lookup"><span data-stu-id="c0128-115">Property</span></span> | <span data-ttu-id="c0128-116">Modelos</span><span class="sxs-lookup"><span data-stu-id="c0128-116">Modes</span></span> | <span data-ttu-id="c0128-117">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="c0128-117">Return type</span></span> | <span data-ttu-id="c0128-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="c0128-118">Minimum</span></span><br><span data-ttu-id="c0128-119">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="c0128-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c0128-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="c0128-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="c0128-121">Escrever</span><span class="sxs-lookup"><span data-stu-id="c0128-121">Compose</span></span><br><span data-ttu-id="c0128-122">Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-122">Read</span></span> | <span data-ttu-id="c0128-123">String</span><span class="sxs-lookup"><span data-stu-id="c0128-123">String</span></span> | [<span data-ttu-id="c0128-124">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c0128-125">la</span><span class="sxs-lookup"><span data-stu-id="c0128-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="c0128-126">Escrever</span><span class="sxs-lookup"><span data-stu-id="c0128-126">Compose</span></span><br><span data-ttu-id="c0128-127">Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-127">Read</span></span> | [<span data-ttu-id="c0128-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="c0128-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c0128-129">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c0128-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="c0128-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="c0128-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="c0128-131">Compose</span></span><br><span data-ttu-id="c0128-132">Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-132">Read</span></span> | <span data-ttu-id="c0128-133">String</span><span class="sxs-lookup"><span data-stu-id="c0128-133">String</span></span> | [<span data-ttu-id="c0128-134">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c0128-135">principal</span><span class="sxs-lookup"><span data-stu-id="c0128-135">host</span></span>](#host-hosttype) | <span data-ttu-id="c0128-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="c0128-136">Compose</span></span><br><span data-ttu-id="c0128-137">Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-137">Read</span></span> | [<span data-ttu-id="c0128-138">HostType</span><span class="sxs-lookup"><span data-stu-id="c0128-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c0128-139">1,5</span><span class="sxs-lookup"><span data-stu-id="c0128-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="c0128-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="c0128-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="c0128-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="c0128-141">Compose</span></span><br><span data-ttu-id="c0128-142">Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-142">Read</span></span> | [<span data-ttu-id="c0128-143">Caixa de Correio</span><span class="sxs-lookup"><span data-stu-id="c0128-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c0128-144">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c0128-145">plataforma</span><span class="sxs-lookup"><span data-stu-id="c0128-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="c0128-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="c0128-146">Compose</span></span><br><span data-ttu-id="c0128-147">Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-147">Read</span></span> | [<span data-ttu-id="c0128-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="c0128-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c0128-149">1,5</span><span class="sxs-lookup"><span data-stu-id="c0128-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="c0128-150">atende</span><span class="sxs-lookup"><span data-stu-id="c0128-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="c0128-151">Escrever</span><span class="sxs-lookup"><span data-stu-id="c0128-151">Compose</span></span><br><span data-ttu-id="c0128-152">Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-152">Read</span></span> | [<span data-ttu-id="c0128-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="c0128-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c0128-154">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c0128-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="c0128-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="c0128-156">Escrever</span><span class="sxs-lookup"><span data-stu-id="c0128-156">Compose</span></span><br><span data-ttu-id="c0128-157">Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-157">Read</span></span> | [<span data-ttu-id="c0128-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c0128-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c0128-159">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c0128-160">ui</span><span class="sxs-lookup"><span data-stu-id="c0128-160">ui</span></span>](#ui-ui) | <span data-ttu-id="c0128-161">Escrever</span><span class="sxs-lookup"><span data-stu-id="c0128-161">Compose</span></span><br><span data-ttu-id="c0128-162">Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-162">Read</span></span> | [<span data-ttu-id="c0128-163">UI</span><span class="sxs-lookup"><span data-stu-id="c0128-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="c0128-164">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="c0128-165">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="c0128-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="c0128-166">contentLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c0128-166">contentLanguage: String</span></span>

<span data-ttu-id="c0128-167">Obtém a localidade (idioma) especificada pelo usuário para edição do item.</span><span class="sxs-lookup"><span data-stu-id="c0128-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="c0128-168">O `contentLanguage` valor reflete a configuração de **idioma de edição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="c0128-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="c0128-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="c0128-169">Type</span></span>

*   <span data-ttu-id="c0128-170">String</span><span class="sxs-lookup"><span data-stu-id="c0128-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0128-171">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c0128-171">Requirements</span></span>

|<span data-ttu-id="c0128-172">Requisito</span><span class="sxs-lookup"><span data-stu-id="c0128-172">Requirement</span></span>| <span data-ttu-id="c0128-173">Valor</span><span class="sxs-lookup"><span data-stu-id="c0128-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0128-174">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c0128-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c0128-175">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-175">1.1</span></span>|
|[<span data-ttu-id="c0128-176">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c0128-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c0128-177">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c0128-178">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c0128-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="c0128-179">diagnóstico: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="c0128-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="c0128-180">Obtém informações sobre o ambiente no qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="c0128-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c0128-181">Tipo</span><span class="sxs-lookup"><span data-stu-id="c0128-181">Type</span></span>

*   [<span data-ttu-id="c0128-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="c0128-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="c0128-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c0128-183">Requirements</span></span>

|<span data-ttu-id="c0128-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="c0128-184">Requirement</span></span>| <span data-ttu-id="c0128-185">Valor</span><span class="sxs-lookup"><span data-stu-id="c0128-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0128-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c0128-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c0128-187">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-187">1.1</span></span>|
|[<span data-ttu-id="c0128-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c0128-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c0128-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c0128-190">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c0128-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="c0128-191">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c0128-191">displayLanguage: String</span></span>

<span data-ttu-id="c0128-192">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="c0128-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="c0128-193">O `displayLanguage` valor reflete a configuração de **idioma de exibição** atual especificada com opções de **arquivo > > idioma** no aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="c0128-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="c0128-194">Tipo</span><span class="sxs-lookup"><span data-stu-id="c0128-194">Type</span></span>

*   <span data-ttu-id="c0128-195">String</span><span class="sxs-lookup"><span data-stu-id="c0128-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0128-196">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c0128-196">Requirements</span></span>

|<span data-ttu-id="c0128-197">Requisito</span><span class="sxs-lookup"><span data-stu-id="c0128-197">Requirement</span></span>| <span data-ttu-id="c0128-198">Valor</span><span class="sxs-lookup"><span data-stu-id="c0128-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0128-199">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c0128-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c0128-200">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-200">1.1</span></span>|
|[<span data-ttu-id="c0128-201">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c0128-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c0128-202">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c0128-203">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c0128-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="c0128-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="c0128-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="c0128-205">Obtém o aplicativo do Office que está hospedando o suplemento.</span><span class="sxs-lookup"><span data-stu-id="c0128-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="c0128-206">Como alternativa, você pode usar a propriedade [Office. Context. Diagnostics](#diagnostics-contextinformation) para obter o host.</span><span class="sxs-lookup"><span data-stu-id="c0128-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="c0128-207">Tipo</span><span class="sxs-lookup"><span data-stu-id="c0128-207">Type</span></span>

*   [<span data-ttu-id="c0128-208">HostType</span><span class="sxs-lookup"><span data-stu-id="c0128-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="c0128-209">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c0128-209">Requirements</span></span>

|<span data-ttu-id="c0128-210">Requisito</span><span class="sxs-lookup"><span data-stu-id="c0128-210">Requirement</span></span>| <span data-ttu-id="c0128-211">Valor</span><span class="sxs-lookup"><span data-stu-id="c0128-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0128-212">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c0128-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c0128-213">1,5</span><span class="sxs-lookup"><span data-stu-id="c0128-213">1.5</span></span>|
|[<span data-ttu-id="c0128-214">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c0128-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c0128-215">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c0128-216">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c0128-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="c0128-217">Platform: [platformtype](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="c0128-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="c0128-218">Fornece a plataforma na qual o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="c0128-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="c0128-219">Como alternativa, você pode usar a propriedade [Office. Context. Diagnostics](#diagnostics-contextinformation) para obter a plataforma.</span><span class="sxs-lookup"><span data-stu-id="c0128-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="c0128-220">Tipo</span><span class="sxs-lookup"><span data-stu-id="c0128-220">Type</span></span>

*   [<span data-ttu-id="c0128-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="c0128-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="c0128-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c0128-222">Requirements</span></span>

|<span data-ttu-id="c0128-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="c0128-223">Requirement</span></span>| <span data-ttu-id="c0128-224">Valor</span><span class="sxs-lookup"><span data-stu-id="c0128-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0128-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c0128-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c0128-226">1,5</span><span class="sxs-lookup"><span data-stu-id="c0128-226">1.5</span></span>|
|[<span data-ttu-id="c0128-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c0128-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c0128-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c0128-229">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c0128-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="c0128-230">requisitos: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="c0128-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="c0128-231">Fornece um método para determinar quais conjuntos de requisitos são compatíveis com o aplicativo atual e a plataforma.</span><span class="sxs-lookup"><span data-stu-id="c0128-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="c0128-232">Tipo</span><span class="sxs-lookup"><span data-stu-id="c0128-232">Type</span></span>

*   [<span data-ttu-id="c0128-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="c0128-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="c0128-234">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c0128-234">Requirements</span></span>

|<span data-ttu-id="c0128-235">Requisito</span><span class="sxs-lookup"><span data-stu-id="c0128-235">Requirement</span></span>| <span data-ttu-id="c0128-236">Valor</span><span class="sxs-lookup"><span data-stu-id="c0128-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0128-237">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c0128-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c0128-238">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-238">1.1</span></span>|
|[<span data-ttu-id="c0128-239">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c0128-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c0128-240">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c0128-241">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c0128-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="c0128-242">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="c0128-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="c0128-243">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="c0128-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="c0128-244">O `RoamingSettings` objeto permite armazenar e acessar dados de um suplemento de email armazenado na caixa de correio de um usuário, para que esteja disponível para esse suplemento quando ele estiver sendo executado a partir de qualquer cliente do Outlook usado para acessar a caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="c0128-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c0128-245">Tipo</span><span class="sxs-lookup"><span data-stu-id="c0128-245">Type</span></span>

*   [<span data-ttu-id="c0128-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c0128-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="c0128-247">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c0128-247">Requirements</span></span>

|<span data-ttu-id="c0128-248">Requisito</span><span class="sxs-lookup"><span data-stu-id="c0128-248">Requirement</span></span>| <span data-ttu-id="c0128-249">Valor</span><span class="sxs-lookup"><span data-stu-id="c0128-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0128-250">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c0128-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c0128-251">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-251">1.1</span></span>|
|[<span data-ttu-id="c0128-252">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c0128-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="c0128-253">Restrito</span><span class="sxs-lookup"><span data-stu-id="c0128-253">Restricted</span></span>|
|[<span data-ttu-id="c0128-254">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c0128-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c0128-255">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="c0128-256">UI: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="c0128-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="c0128-257">Fornece objetos e métodos que você pode usar para criar e manipular componentes da interface do usuário, como caixas de diálogo, em seus suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="c0128-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c0128-258">Tipo</span><span class="sxs-lookup"><span data-stu-id="c0128-258">Type</span></span>

*   [<span data-ttu-id="c0128-259">UI</span><span class="sxs-lookup"><span data-stu-id="c0128-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="c0128-260">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c0128-260">Requirements</span></span>

|<span data-ttu-id="c0128-261">Requisito</span><span class="sxs-lookup"><span data-stu-id="c0128-261">Requirement</span></span>| <span data-ttu-id="c0128-262">Valor</span><span class="sxs-lookup"><span data-stu-id="c0128-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0128-263">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c0128-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c0128-264">1.1</span><span class="sxs-lookup"><span data-stu-id="c0128-264">1.1</span></span>|
|[<span data-ttu-id="c0128-265">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c0128-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c0128-266">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c0128-266">Compose or Read</span></span>|
