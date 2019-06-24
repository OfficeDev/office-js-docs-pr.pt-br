---
title: Office. Context – conjunto de requisitos 1,7
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: ff816b3bb51ebb5dc8ef124af8488405fdc3fd39
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127132"
---
# <a name="context"></a><span data-ttu-id="184fb-102">context</span><span class="sxs-lookup"><span data-stu-id="184fb-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="184fb-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="184fb-103">[Office](Office.md).context</span></span>

<span data-ttu-id="184fb-104">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="184fb-104">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="184fb-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="184fb-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="184fb-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="184fb-106">Requirements</span></span>

|<span data-ttu-id="184fb-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="184fb-107">Requirement</span></span>| <span data-ttu-id="184fb-108">Valor</span><span class="sxs-lookup"><span data-stu-id="184fb-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="184fb-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="184fb-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="184fb-110">1.0</span><span class="sxs-lookup"><span data-stu-id="184fb-110">1.0</span></span>|
|[<span data-ttu-id="184fb-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="184fb-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="184fb-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="184fb-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="184fb-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="184fb-113">Members and methods</span></span>

| <span data-ttu-id="184fb-114">Membro</span><span class="sxs-lookup"><span data-stu-id="184fb-114">Member</span></span> | <span data-ttu-id="184fb-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="184fb-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="184fb-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="184fb-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="184fb-117">Membro</span><span class="sxs-lookup"><span data-stu-id="184fb-117">Member</span></span> |
| [<span data-ttu-id="184fb-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="184fb-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="184fb-119">Membro</span><span class="sxs-lookup"><span data-stu-id="184fb-119">Member</span></span> |
| [<span data-ttu-id="184fb-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="184fb-120">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="184fb-121">Membro</span><span class="sxs-lookup"><span data-stu-id="184fb-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="184fb-122">Namespaces</span><span class="sxs-lookup"><span data-stu-id="184fb-122">Namespaces</span></span>

<span data-ttu-id="184fb-123">[Mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="184fb-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="184fb-124">Members</span><span class="sxs-lookup"><span data-stu-id="184fb-124">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="184fb-125">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="184fb-125">displayLanguage: String</span></span>

<span data-ttu-id="184fb-126">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="184fb-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="184fb-127">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="184fb-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="184fb-128">Tipo</span><span class="sxs-lookup"><span data-stu-id="184fb-128">Type</span></span>

*   <span data-ttu-id="184fb-129">String</span><span class="sxs-lookup"><span data-stu-id="184fb-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="184fb-130">Requisitos</span><span class="sxs-lookup"><span data-stu-id="184fb-130">Requirements</span></span>

|<span data-ttu-id="184fb-131">Requisito</span><span class="sxs-lookup"><span data-stu-id="184fb-131">Requirement</span></span>| <span data-ttu-id="184fb-132">Valor</span><span class="sxs-lookup"><span data-stu-id="184fb-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="184fb-133">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="184fb-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="184fb-134">1.0</span><span class="sxs-lookup"><span data-stu-id="184fb-134">1.0</span></span>|
|[<span data-ttu-id="184fb-135">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="184fb-135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="184fb-136">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="184fb-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="184fb-137">Exemplo</span><span class="sxs-lookup"><span data-stu-id="184fb-137">Example</span></span>

```javascript
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

---
---

#### <a name="officetheme-object"></a><span data-ttu-id="184fb-138">officeTheme: objeto</span><span class="sxs-lookup"><span data-stu-id="184fb-138">officeTheme: Object</span></span>

<span data-ttu-id="184fb-139">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="184fb-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="184fb-140">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="184fb-140">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="184fb-p102">Usar as cores de tema do Office possibilita coordenar o esquema de cores de seu suplemento com o tema do Office atualmente selecionado pelo usuário em \*\*Arquivo > Conta do Office > Tema da interface de usuário do Office \*\*, que é aplicado a todos os aplicativos host do Office. Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="184fb-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="184fb-143">Tipo</span><span class="sxs-lookup"><span data-stu-id="184fb-143">Type</span></span>

*   <span data-ttu-id="184fb-144">Objeto</span><span class="sxs-lookup"><span data-stu-id="184fb-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="184fb-145">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="184fb-145">Properties:</span></span>

|<span data-ttu-id="184fb-146">Nome</span><span class="sxs-lookup"><span data-stu-id="184fb-146">Name</span></span>| <span data-ttu-id="184fb-147">Tipo</span><span class="sxs-lookup"><span data-stu-id="184fb-147">Type</span></span>| <span data-ttu-id="184fb-148">Descrição</span><span class="sxs-lookup"><span data-stu-id="184fb-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="184fb-149">String</span><span class="sxs-lookup"><span data-stu-id="184fb-149">String</span></span>|<span data-ttu-id="184fb-150">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="184fb-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="184fb-151">String</span><span class="sxs-lookup"><span data-stu-id="184fb-151">String</span></span>|<span data-ttu-id="184fb-152">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="184fb-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="184fb-153">String</span><span class="sxs-lookup"><span data-stu-id="184fb-153">String</span></span>|<span data-ttu-id="184fb-154">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="184fb-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="184fb-155">String</span><span class="sxs-lookup"><span data-stu-id="184fb-155">String</span></span>|<span data-ttu-id="184fb-156">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="184fb-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="184fb-157">Requisitos</span><span class="sxs-lookup"><span data-stu-id="184fb-157">Requirements</span></span>

|<span data-ttu-id="184fb-158">Requisito</span><span class="sxs-lookup"><span data-stu-id="184fb-158">Requirement</span></span>| <span data-ttu-id="184fb-159">Valor</span><span class="sxs-lookup"><span data-stu-id="184fb-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="184fb-160">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="184fb-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="184fb-161">1.3</span><span class="sxs-lookup"><span data-stu-id="184fb-161">1.3</span></span>|
|[<span data-ttu-id="184fb-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="184fb-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="184fb-163">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="184fb-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="184fb-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="184fb-164">Example</span></span>

```javascript
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

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook17officeroamingsettings"></a><span data-ttu-id="184fb-165">roamingSettings: [roamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="184fb-165">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span></span>

<span data-ttu-id="184fb-166">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="184fb-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="184fb-167">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="184fb-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="184fb-168">Tipo</span><span class="sxs-lookup"><span data-stu-id="184fb-168">Type</span></span>

*   [<span data-ttu-id="184fb-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="184fb-169">RoamingSettings</span></span>](/javascript/api/outlook_1_7/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="184fb-170">Requisitos</span><span class="sxs-lookup"><span data-stu-id="184fb-170">Requirements</span></span>

|<span data-ttu-id="184fb-171">Requisito</span><span class="sxs-lookup"><span data-stu-id="184fb-171">Requirement</span></span>| <span data-ttu-id="184fb-172">Valor</span><span class="sxs-lookup"><span data-stu-id="184fb-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="184fb-173">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="184fb-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="184fb-174">1.0</span><span class="sxs-lookup"><span data-stu-id="184fb-174">1.0</span></span>|
|[<span data-ttu-id="184fb-175">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="184fb-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="184fb-176">Restrito</span><span class="sxs-lookup"><span data-stu-id="184fb-176">Restricted</span></span>|
|[<span data-ttu-id="184fb-177">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="184fb-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="184fb-178">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="184fb-178">Compose or Read</span></span>|
