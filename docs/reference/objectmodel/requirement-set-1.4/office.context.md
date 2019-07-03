---
title: Office. Context – conjunto de requisitos 1,4
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: cd5cd136c8d801f5ac8a607da2fa1a165961905e
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454822"
---
# <a name="context"></a><span data-ttu-id="a4751-102">context</span><span class="sxs-lookup"><span data-stu-id="a4751-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="a4751-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="a4751-103">[Office](Office.md).context</span></span>

<span data-ttu-id="a4751-104">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="a4751-104">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="a4751-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="a4751-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4751-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a4751-106">Requirements</span></span>

|<span data-ttu-id="a4751-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="a4751-107">Requirement</span></span>| <span data-ttu-id="a4751-108">Valor</span><span class="sxs-lookup"><span data-stu-id="a4751-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4751-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a4751-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4751-110">1.0</span><span class="sxs-lookup"><span data-stu-id="a4751-110">1.0</span></span>|
|[<span data-ttu-id="a4751-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a4751-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a4751-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a4751-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="a4751-113">Namespaces</span><span class="sxs-lookup"><span data-stu-id="a4751-113">Namespaces</span></span>

<span data-ttu-id="a4751-114">[Mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="a4751-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="a4751-115">Members</span><span class="sxs-lookup"><span data-stu-id="a4751-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="a4751-116">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a4751-116">displayLanguage: String</span></span>

<span data-ttu-id="a4751-117">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="a4751-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="a4751-118">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="a4751-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="a4751-119">Tipo</span><span class="sxs-lookup"><span data-stu-id="a4751-119">Type</span></span>

*   <span data-ttu-id="a4751-120">String</span><span class="sxs-lookup"><span data-stu-id="a4751-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4751-121">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a4751-121">Requirements</span></span>

|<span data-ttu-id="a4751-122">Requisito</span><span class="sxs-lookup"><span data-stu-id="a4751-122">Requirement</span></span>| <span data-ttu-id="a4751-123">Valor</span><span class="sxs-lookup"><span data-stu-id="a4751-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4751-124">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a4751-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4751-125">1.0</span><span class="sxs-lookup"><span data-stu-id="a4751-125">1.0</span></span>|
|[<span data-ttu-id="a4751-126">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a4751-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a4751-127">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a4751-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a4751-128">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a4751-128">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook14officeroamingsettings"></a><span data-ttu-id="a4751-129">roamingSettings: [roamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="a4751-129">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span></span>

<span data-ttu-id="a4751-130">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="a4751-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="a4751-131">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="a4751-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="a4751-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="a4751-132">Type</span></span>

*   [<span data-ttu-id="a4751-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a4751-133">RoamingSettings</span></span>](/javascript/api/outlook_1_4/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="a4751-134">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a4751-134">Requirements</span></span>

|<span data-ttu-id="a4751-135">Requisito</span><span class="sxs-lookup"><span data-stu-id="a4751-135">Requirement</span></span>| <span data-ttu-id="a4751-136">Valor</span><span class="sxs-lookup"><span data-stu-id="a4751-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4751-137">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a4751-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4751-138">1.0</span><span class="sxs-lookup"><span data-stu-id="a4751-138">1.0</span></span>|
|[<span data-ttu-id="a4751-139">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a4751-139">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a4751-140">Restrito</span><span class="sxs-lookup"><span data-stu-id="a4751-140">Restricted</span></span>|
|[<span data-ttu-id="a4751-141">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a4751-141">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a4751-142">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a4751-142">Compose or Read</span></span>|
