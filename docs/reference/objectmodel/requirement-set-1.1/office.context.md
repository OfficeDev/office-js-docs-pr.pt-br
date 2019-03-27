---
title: Office. Context – conjunto de requisitos 1,1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c7c33f31a63fc2a57b9457fc788a80c925d3eef1
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871196"
---
# <a name="context"></a><span data-ttu-id="d393d-102">context</span><span class="sxs-lookup"><span data-stu-id="d393d-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="d393d-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="d393d-103">[Office](Office.md).context</span></span>

<span data-ttu-id="d393d-104">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="d393d-104">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="d393d-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="d393d-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="d393d-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d393d-106">Requirements</span></span>

|<span data-ttu-id="d393d-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="d393d-107">Requirement</span></span>| <span data-ttu-id="d393d-108">Valor</span><span class="sxs-lookup"><span data-stu-id="d393d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d393d-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d393d-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d393d-110">1.0</span><span class="sxs-lookup"><span data-stu-id="d393d-110">1.0</span></span>|
|[<span data-ttu-id="d393d-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d393d-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d393d-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d393d-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="d393d-113">Namespaces</span><span class="sxs-lookup"><span data-stu-id="d393d-113">Namespaces</span></span>

<span data-ttu-id="d393d-114">[mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="d393d-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="d393d-115">Membros</span><span class="sxs-lookup"><span data-stu-id="d393d-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="d393d-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="d393d-116">displayLanguage :String</span></span>

<span data-ttu-id="d393d-117">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="d393d-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="d393d-118">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="d393d-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="d393d-119">Tipo</span><span class="sxs-lookup"><span data-stu-id="d393d-119">Type</span></span>

*   <span data-ttu-id="d393d-120">String</span><span class="sxs-lookup"><span data-stu-id="d393d-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d393d-121">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d393d-121">Requirements</span></span>

|<span data-ttu-id="d393d-122">Requisito</span><span class="sxs-lookup"><span data-stu-id="d393d-122">Requirement</span></span>| <span data-ttu-id="d393d-123">Valor</span><span class="sxs-lookup"><span data-stu-id="d393d-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="d393d-124">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d393d-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d393d-125">1.0</span><span class="sxs-lookup"><span data-stu-id="d393d-125">1.0</span></span>|
|[<span data-ttu-id="d393d-126">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d393d-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d393d-127">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d393d-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d393d-128">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d393d-128">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook11officeroamingsettings"></a><span data-ttu-id="d393d-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="d393d-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)</span></span>

<span data-ttu-id="d393d-130">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="d393d-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="d393d-131">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="d393d-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="d393d-132">Tipo</span><span class="sxs-lookup"><span data-stu-id="d393d-132">Type</span></span>

*   [<span data-ttu-id="d393d-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="d393d-133">RoamingSettings</span></span>](/javascript/api/outlook_1_1/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="d393d-134">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d393d-134">Requirements</span></span>

|<span data-ttu-id="d393d-135">Requisito</span><span class="sxs-lookup"><span data-stu-id="d393d-135">Requirement</span></span>| <span data-ttu-id="d393d-136">Valor</span><span class="sxs-lookup"><span data-stu-id="d393d-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="d393d-137">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d393d-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d393d-138">1.0</span><span class="sxs-lookup"><span data-stu-id="d393d-138">1.0</span></span>|
|[<span data-ttu-id="d393d-139">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d393d-139">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d393d-140">Restrito</span><span class="sxs-lookup"><span data-stu-id="d393d-140">Restricted</span></span>|
|[<span data-ttu-id="d393d-141">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d393d-141">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d393d-142">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d393d-142">Compose or Read</span></span>|
