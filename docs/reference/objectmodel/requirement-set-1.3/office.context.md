---
title: Office. Context – conjunto de requisitos 1,3
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: ccc0a3addb5a440daf841452883019aa9f2b80c8
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127412"
---
# <a name="context"></a><span data-ttu-id="a574c-102">context</span><span class="sxs-lookup"><span data-stu-id="a574c-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="a574c-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="a574c-103">[Office](Office.md).context</span></span>

<span data-ttu-id="a574c-104">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="a574c-104">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="a574c-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="a574c-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a574c-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a574c-106">Requirements</span></span>

|<span data-ttu-id="a574c-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="a574c-107">Requirement</span></span>| <span data-ttu-id="a574c-108">Valor</span><span class="sxs-lookup"><span data-stu-id="a574c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="a574c-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a574c-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a574c-110">1.0</span><span class="sxs-lookup"><span data-stu-id="a574c-110">1.0</span></span>|
|[<span data-ttu-id="a574c-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a574c-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a574c-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a574c-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="a574c-113">Namespaces</span><span class="sxs-lookup"><span data-stu-id="a574c-113">Namespaces</span></span>

<span data-ttu-id="a574c-114">[Mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="a574c-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="a574c-115">Members</span><span class="sxs-lookup"><span data-stu-id="a574c-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="a574c-116">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a574c-116">displayLanguage: String</span></span>

<span data-ttu-id="a574c-117">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="a574c-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="a574c-118">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="a574c-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="a574c-119">Tipo</span><span class="sxs-lookup"><span data-stu-id="a574c-119">Type</span></span>

*   <span data-ttu-id="a574c-120">String</span><span class="sxs-lookup"><span data-stu-id="a574c-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a574c-121">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a574c-121">Requirements</span></span>

|<span data-ttu-id="a574c-122">Requisito</span><span class="sxs-lookup"><span data-stu-id="a574c-122">Requirement</span></span>| <span data-ttu-id="a574c-123">Valor</span><span class="sxs-lookup"><span data-stu-id="a574c-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="a574c-124">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a574c-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a574c-125">1.0</span><span class="sxs-lookup"><span data-stu-id="a574c-125">1.0</span></span>|
|[<span data-ttu-id="a574c-126">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a574c-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a574c-127">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a574c-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a574c-128">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a574c-128">Example</span></span>

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

#### <a name="officetheme-object"></a><span data-ttu-id="a574c-129">officeTheme: objeto</span><span class="sxs-lookup"><span data-stu-id="a574c-129">officeTheme: Object</span></span>

<span data-ttu-id="a574c-130">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="a574c-130">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="a574c-131">Não há suporte para esse membro no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="a574c-131">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a574c-p102">Usar as cores de tema do Office possibilita coordenar o esquema de cores de seu suplemento com o tema do Office atualmente selecionado pelo usuário em \*\*Arquivo > Conta do Office > Tema da interface de usuário do Office \*\*, que é aplicado a todos os aplicativos host do Office. Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="a574c-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="a574c-134">Tipo</span><span class="sxs-lookup"><span data-stu-id="a574c-134">Type</span></span>

*   <span data-ttu-id="a574c-135">Objeto</span><span class="sxs-lookup"><span data-stu-id="a574c-135">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="a574c-136">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a574c-136">Properties:</span></span>

|<span data-ttu-id="a574c-137">Nome</span><span class="sxs-lookup"><span data-stu-id="a574c-137">Name</span></span>| <span data-ttu-id="a574c-138">Tipo</span><span class="sxs-lookup"><span data-stu-id="a574c-138">Type</span></span>| <span data-ttu-id="a574c-139">Descrição</span><span class="sxs-lookup"><span data-stu-id="a574c-139">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="a574c-140">String</span><span class="sxs-lookup"><span data-stu-id="a574c-140">String</span></span>|<span data-ttu-id="a574c-141">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="a574c-141">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="a574c-142">String</span><span class="sxs-lookup"><span data-stu-id="a574c-142">String</span></span>|<span data-ttu-id="a574c-143">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="a574c-143">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="a574c-144">String</span><span class="sxs-lookup"><span data-stu-id="a574c-144">String</span></span>|<span data-ttu-id="a574c-145">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="a574c-145">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="a574c-146">String</span><span class="sxs-lookup"><span data-stu-id="a574c-146">String</span></span>|<span data-ttu-id="a574c-147">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="a574c-147">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a574c-148">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a574c-148">Requirements</span></span>

|<span data-ttu-id="a574c-149">Requisito</span><span class="sxs-lookup"><span data-stu-id="a574c-149">Requirement</span></span>| <span data-ttu-id="a574c-150">Valor</span><span class="sxs-lookup"><span data-stu-id="a574c-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="a574c-151">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a574c-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a574c-152">1.3</span><span class="sxs-lookup"><span data-stu-id="a574c-152">1.3</span></span>|
|[<span data-ttu-id="a574c-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a574c-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a574c-154">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a574c-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a574c-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a574c-155">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook13officeroamingsettings"></a><span data-ttu-id="a574c-156">roamingSettings: [roamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="a574c-156">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span></span>

<span data-ttu-id="a574c-157">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="a574c-157">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="a574c-158">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="a574c-158">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="a574c-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="a574c-159">Type</span></span>

*   [<span data-ttu-id="a574c-160">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a574c-160">RoamingSettings</span></span>](/javascript/api/outlook_1_3/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="a574c-161">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a574c-161">Requirements</span></span>

|<span data-ttu-id="a574c-162">Requisito</span><span class="sxs-lookup"><span data-stu-id="a574c-162">Requirement</span></span>| <span data-ttu-id="a574c-163">Valor</span><span class="sxs-lookup"><span data-stu-id="a574c-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="a574c-164">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a574c-164">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a574c-165">1.0</span><span class="sxs-lookup"><span data-stu-id="a574c-165">1.0</span></span>|
|[<span data-ttu-id="a574c-166">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a574c-166">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a574c-167">Restrito</span><span class="sxs-lookup"><span data-stu-id="a574c-167">Restricted</span></span>|
|[<span data-ttu-id="a574c-168">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a574c-168">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a574c-169">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a574c-169">Compose or Read</span></span>|
