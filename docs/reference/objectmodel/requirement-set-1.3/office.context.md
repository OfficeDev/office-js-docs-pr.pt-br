---
title: Office.context – conjunto de requisitos da versão 1.3
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: 9652a1d4dcd48c437bb1156e9abc4c2aff575d59
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457541"
---
# <a name="context"></a><span data-ttu-id="b346f-102">context</span><span class="sxs-lookup"><span data-stu-id="b346f-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="b346f-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="b346f-103">[Office](Office.md).context</span></span>

<span data-ttu-id="b346f-104">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="b346f-104">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="b346f-105">Esta listagem documenta apenas as interfaces usadas pelos suplementos do Outlook. Para uma listagem completa de namespace Office.context, confira a [referência Office.context na API comum](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="b346f-105">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b346f-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b346f-106">Requirements</span></span>

|<span data-ttu-id="b346f-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="b346f-107">Requirement</span></span>| <span data-ttu-id="b346f-108">Valor</span><span class="sxs-lookup"><span data-stu-id="b346f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b346f-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b346f-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b346f-110">1.0</span><span class="sxs-lookup"><span data-stu-id="b346f-110">1.0</span></span>|
|[<span data-ttu-id="b346f-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b346f-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b346f-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b346f-112">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="b346f-113">Namespaces</span><span class="sxs-lookup"><span data-stu-id="b346f-113">Namespaces</span></span>

<span data-ttu-id="b346f-114">[mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="b346f-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="b346f-115">Membros</span><span class="sxs-lookup"><span data-stu-id="b346f-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="b346f-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="b346f-116">displayLanguage :String</span></span>

<span data-ttu-id="b346f-117">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="b346f-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="b346f-118">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="b346f-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="b346f-119">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b346f-119">Type:</span></span>

*   <span data-ttu-id="b346f-120">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b346f-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b346f-121">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b346f-121">Requirements</span></span>

|<span data-ttu-id="b346f-122">Requisito</span><span class="sxs-lookup"><span data-stu-id="b346f-122">Requirement</span></span>| <span data-ttu-id="b346f-123">Valor</span><span class="sxs-lookup"><span data-stu-id="b346f-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="b346f-124">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b346f-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b346f-125">1.0</span><span class="sxs-lookup"><span data-stu-id="b346f-125">1.0</span></span>|
|[<span data-ttu-id="b346f-126">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b346f-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b346f-127">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b346f-127">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b346f-128">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b346f-128">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="b346f-129">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="b346f-129">officeTheme :Object</span></span>

<span data-ttu-id="b346f-130">Fornece acesso às propriedades de cores de temas do Office.</span><span class="sxs-lookup"><span data-stu-id="b346f-130">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="b346f-131">Não há suporte para esse membro no Outlook para iOS ou no Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="b346f-131">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b346f-p102">Usar as cores de tema do Office possibilita coordenar o esquema de cores de seu suplemento com o tema do Office atualmente selecionado pelo usuário em \*\*Arquivo > Conta do Office > Tema da interface de usuário do Office \*\*, que é aplicado a todos os aplicativos host do Office. Usar cores de temas do Office é apropriado suplementos de email e painéis de tarefas.</span><span class="sxs-lookup"><span data-stu-id="b346f-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="b346f-134">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b346f-134">Type:</span></span>

*   <span data-ttu-id="b346f-135">Objeto</span><span class="sxs-lookup"><span data-stu-id="b346f-135">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="b346f-136">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b346f-136">Properties:</span></span>

|<span data-ttu-id="b346f-137">Nome</span><span class="sxs-lookup"><span data-stu-id="b346f-137">Name</span></span>| <span data-ttu-id="b346f-138">Tipo</span><span class="sxs-lookup"><span data-stu-id="b346f-138">Type</span></span>| <span data-ttu-id="b346f-139">Descrição</span><span class="sxs-lookup"><span data-stu-id="b346f-139">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="b346f-140">String</span><span class="sxs-lookup"><span data-stu-id="b346f-140">String</span></span>|<span data-ttu-id="b346f-141">Obtém a cor de plano de fundo do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="b346f-141">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="b346f-142">String</span><span class="sxs-lookup"><span data-stu-id="b346f-142">String</span></span>|<span data-ttu-id="b346f-143">Obtém a cor de primeiro plano do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="b346f-143">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="b346f-144">String</span><span class="sxs-lookup"><span data-stu-id="b346f-144">String</span></span>|<span data-ttu-id="b346f-145">Obtém a cor de plano de fundo do controle do tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="b346f-145">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="b346f-146">String</span><span class="sxs-lookup"><span data-stu-id="b346f-146">String</span></span>|<span data-ttu-id="b346f-147">Obtém a cor de controle do corpo de tema do Office como um tripleto hexadecimal de cores.</span><span class="sxs-lookup"><span data-stu-id="b346f-147">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b346f-148">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b346f-148">Requirements</span></span>

|<span data-ttu-id="b346f-149">Requisito</span><span class="sxs-lookup"><span data-stu-id="b346f-149">Requirement</span></span>| <span data-ttu-id="b346f-150">Valor</span><span class="sxs-lookup"><span data-stu-id="b346f-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="b346f-151">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b346f-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b346f-152">1.3</span><span class="sxs-lookup"><span data-stu-id="b346f-152">1.3</span></span>|
|[<span data-ttu-id="b346f-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b346f-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b346f-154">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="b346f-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b346f-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b346f-155">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook13officeroamingsettings"></a><span data-ttu-id="b346f-156">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="b346f-156">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span></span>

<span data-ttu-id="b346f-157">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="b346f-157">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="b346f-158">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="b346f-158">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="b346f-159">Tipo:</span><span class="sxs-lookup"><span data-stu-id="b346f-159">Type:</span></span>

*   [<span data-ttu-id="b346f-160">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b346f-160">RoamingSettings</span></span>](/javascript/api/outlook_1_3/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="b346f-161">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b346f-161">Requirements</span></span>

|<span data-ttu-id="b346f-162">Requisito</span><span class="sxs-lookup"><span data-stu-id="b346f-162">Requirement</span></span>| <span data-ttu-id="b346f-163">Valor</span><span class="sxs-lookup"><span data-stu-id="b346f-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="b346f-164">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b346f-164">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b346f-165">1.0</span><span class="sxs-lookup"><span data-stu-id="b346f-165">1.0</span></span>|
|[<span data-ttu-id="b346f-166">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b346f-166">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b346f-167">Restrito</span><span class="sxs-lookup"><span data-stu-id="b346f-167">Restricted</span></span>|
|[<span data-ttu-id="b346f-168">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b346f-168">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b346f-169">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b346f-169">Compose or read</span></span>|