---
title: Office. Context – conjunto de requisitos de visualização
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: f06509e291325c635581d902d1f4f440bd255314
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696460"
---
# <a name="context"></a><span data-ttu-id="cc37f-102">context</span><span class="sxs-lookup"><span data-stu-id="cc37f-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="cc37f-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="cc37f-103">[Office](Office.md).context</span></span>

<span data-ttu-id="cc37f-104">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="cc37f-104">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="cc37f-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="cc37f-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc37f-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cc37f-106">Requirements</span></span>

|<span data-ttu-id="cc37f-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="cc37f-107">Requirement</span></span>| <span data-ttu-id="cc37f-108">Valor</span><span class="sxs-lookup"><span data-stu-id="cc37f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc37f-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cc37f-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc37f-110">1.0</span><span class="sxs-lookup"><span data-stu-id="cc37f-110">1.0</span></span>|
|[<span data-ttu-id="cc37f-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cc37f-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc37f-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cc37f-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="cc37f-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="cc37f-113">Members and methods</span></span>

| <span data-ttu-id="cc37f-114">Membro</span><span class="sxs-lookup"><span data-stu-id="cc37f-114">Member</span></span> | <span data-ttu-id="cc37f-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="cc37f-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="cc37f-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="cc37f-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="cc37f-117">Membro</span><span class="sxs-lookup"><span data-stu-id="cc37f-117">Member</span></span> |
| [<span data-ttu-id="cc37f-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="cc37f-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="cc37f-119">Membro</span><span class="sxs-lookup"><span data-stu-id="cc37f-119">Member</span></span> |
| [<span data-ttu-id="cc37f-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="cc37f-120">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="cc37f-121">Membro</span><span class="sxs-lookup"><span data-stu-id="cc37f-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="cc37f-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="cc37f-122">Namespaces</span></span>

<span data-ttu-id="cc37f-123">[Mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="cc37f-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="cc37f-124">Members</span><span class="sxs-lookup"><span data-stu-id="cc37f-124">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="cc37f-125">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="cc37f-125">displayLanguage: String</span></span>

<span data-ttu-id="cc37f-126">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="cc37f-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="cc37f-127">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="cc37f-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="cc37f-128">Tipo</span><span class="sxs-lookup"><span data-stu-id="cc37f-128">Type</span></span>

*   <span data-ttu-id="cc37f-129">String</span><span class="sxs-lookup"><span data-stu-id="cc37f-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc37f-130">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cc37f-130">Requirements</span></span>

|<span data-ttu-id="cc37f-131">Requisito</span><span class="sxs-lookup"><span data-stu-id="cc37f-131">Requirement</span></span>| <span data-ttu-id="cc37f-132">Valor</span><span class="sxs-lookup"><span data-stu-id="cc37f-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc37f-133">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cc37f-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc37f-134">1.0</span><span class="sxs-lookup"><span data-stu-id="cc37f-134">1.0</span></span>|
|[<span data-ttu-id="cc37f-135">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cc37f-135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc37f-136">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cc37f-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc37f-137">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cc37f-137">Example</span></span>

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

#### <a name="officetheme-object"></a><span data-ttu-id="cc37f-138">officeTheme: objeto</span><span class="sxs-lookup"><span data-stu-id="cc37f-138">officeTheme: Object</span></span>

<span data-ttu-id="cc37f-139">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="cc37f-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="cc37f-140">Só há suporte para esse membro no Outlook no Windows.</span><span class="sxs-lookup"><span data-stu-id="cc37f-140">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="cc37f-141">O uso de cores de temas do Office permite coordenar o esquema de cores do seu suplemento com o tema atual do Office selecionado pelo usuário com a **conta de arquivo > office > Office Theme UI**, que é aplicada em todos os aplicativos host do Office.</span><span class="sxs-lookup"><span data-stu-id="cc37f-141">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="cc37f-142">Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="cc37f-142">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="cc37f-143">Tipo</span><span class="sxs-lookup"><span data-stu-id="cc37f-143">Type</span></span>

*   <span data-ttu-id="cc37f-144">Objeto</span><span class="sxs-lookup"><span data-stu-id="cc37f-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="cc37f-145">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="cc37f-145">Properties:</span></span>

|<span data-ttu-id="cc37f-146">Nome</span><span class="sxs-lookup"><span data-stu-id="cc37f-146">Name</span></span>| <span data-ttu-id="cc37f-147">Tipo</span><span class="sxs-lookup"><span data-stu-id="cc37f-147">Type</span></span>| <span data-ttu-id="cc37f-148">Descrição</span><span class="sxs-lookup"><span data-stu-id="cc37f-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="cc37f-149">String</span><span class="sxs-lookup"><span data-stu-id="cc37f-149">String</span></span>|<span data-ttu-id="cc37f-150">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="cc37f-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="cc37f-151">String</span><span class="sxs-lookup"><span data-stu-id="cc37f-151">String</span></span>|<span data-ttu-id="cc37f-152">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="cc37f-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="cc37f-153">String</span><span class="sxs-lookup"><span data-stu-id="cc37f-153">String</span></span>|<span data-ttu-id="cc37f-154">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="cc37f-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="cc37f-155">String</span><span class="sxs-lookup"><span data-stu-id="cc37f-155">String</span></span>|<span data-ttu-id="cc37f-156">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="cc37f-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc37f-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cc37f-157">Requirements</span></span>

|<span data-ttu-id="cc37f-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="cc37f-158">Requirement</span></span>| <span data-ttu-id="cc37f-159">Valor</span><span class="sxs-lookup"><span data-stu-id="cc37f-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc37f-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cc37f-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc37f-161">Visualização</span><span class="sxs-lookup"><span data-stu-id="cc37f-161">Preview</span></span>|
|[<span data-ttu-id="cc37f-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cc37f-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc37f-163">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cc37f-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc37f-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="cc37f-164">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="cc37f-165">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="cc37f-165">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span></span>

<span data-ttu-id="cc37f-166">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="cc37f-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="cc37f-167">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="cc37f-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="cc37f-168">Tipo</span><span class="sxs-lookup"><span data-stu-id="cc37f-168">Type</span></span>

*   [<span data-ttu-id="cc37f-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="cc37f-169">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="cc37f-170">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cc37f-170">Requirements</span></span>

|<span data-ttu-id="cc37f-171">Requisito</span><span class="sxs-lookup"><span data-stu-id="cc37f-171">Requirement</span></span>| <span data-ttu-id="cc37f-172">Valor</span><span class="sxs-lookup"><span data-stu-id="cc37f-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc37f-173">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cc37f-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc37f-174">1.0</span><span class="sxs-lookup"><span data-stu-id="cc37f-174">1.0</span></span>|
|[<span data-ttu-id="cc37f-175">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="cc37f-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc37f-176">Restrito</span><span class="sxs-lookup"><span data-stu-id="cc37f-176">Restricted</span></span>|
|[<span data-ttu-id="cc37f-177">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cc37f-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc37f-178">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cc37f-178">Compose or Read</span></span>|
