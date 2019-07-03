---
title: Office. Context – conjunto de requisitos 1,3
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: c98c8b41dda566ff9f66818ebe1398d11d4d0749
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454864"
---
# <a name="context"></a><span data-ttu-id="a1078-102">context</span><span class="sxs-lookup"><span data-stu-id="a1078-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="a1078-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="a1078-103">[Office](Office.md).context</span></span>

<span data-ttu-id="a1078-104">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="a1078-104">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="a1078-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="a1078-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1078-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a1078-106">Requirements</span></span>

|<span data-ttu-id="a1078-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="a1078-107">Requirement</span></span>| <span data-ttu-id="a1078-108">Valor</span><span class="sxs-lookup"><span data-stu-id="a1078-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1078-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a1078-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1078-110">1.0</span><span class="sxs-lookup"><span data-stu-id="a1078-110">1.0</span></span>|
|[<span data-ttu-id="a1078-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a1078-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a1078-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a1078-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="a1078-113">Namespaces</span><span class="sxs-lookup"><span data-stu-id="a1078-113">Namespaces</span></span>

<span data-ttu-id="a1078-114">[Mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="a1078-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="a1078-115">Members</span><span class="sxs-lookup"><span data-stu-id="a1078-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="a1078-116">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1078-116">displayLanguage: String</span></span>

<span data-ttu-id="a1078-117">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="a1078-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="a1078-118">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="a1078-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="a1078-119">Tipo</span><span class="sxs-lookup"><span data-stu-id="a1078-119">Type</span></span>

*   <span data-ttu-id="a1078-120">String</span><span class="sxs-lookup"><span data-stu-id="a1078-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1078-121">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a1078-121">Requirements</span></span>

|<span data-ttu-id="a1078-122">Requisito</span><span class="sxs-lookup"><span data-stu-id="a1078-122">Requirement</span></span>| <span data-ttu-id="a1078-123">Valor</span><span class="sxs-lookup"><span data-stu-id="a1078-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1078-124">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a1078-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1078-125">1.0</span><span class="sxs-lookup"><span data-stu-id="a1078-125">1.0</span></span>|
|[<span data-ttu-id="a1078-126">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a1078-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a1078-127">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a1078-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1078-128">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a1078-128">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook13officeroamingsettings"></a><span data-ttu-id="a1078-129">roamingSettings: [roamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="a1078-129">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span></span>

<span data-ttu-id="a1078-130">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="a1078-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="a1078-131">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="a1078-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="a1078-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="a1078-132">Type</span></span>

*   [<span data-ttu-id="a1078-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a1078-133">RoamingSettings</span></span>](/javascript/api/outlook_1_3/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="a1078-134">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a1078-134">Requirements</span></span>

|<span data-ttu-id="a1078-135">Requisito</span><span class="sxs-lookup"><span data-stu-id="a1078-135">Requirement</span></span>| <span data-ttu-id="a1078-136">Valor</span><span class="sxs-lookup"><span data-stu-id="a1078-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1078-137">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a1078-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a1078-138">1.0</span><span class="sxs-lookup"><span data-stu-id="a1078-138">1.0</span></span>|
|[<span data-ttu-id="a1078-139">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a1078-139">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a1078-140">Restrito</span><span class="sxs-lookup"><span data-stu-id="a1078-140">Restricted</span></span>|
|[<span data-ttu-id="a1078-141">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a1078-141">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a1078-142">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a1078-142">Compose or Read</span></span>|
