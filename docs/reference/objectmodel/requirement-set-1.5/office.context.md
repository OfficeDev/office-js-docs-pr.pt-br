---
title: Office. Context – conjunto de requisitos 1,5
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: 816e28a5ea8d270b2223ff5c24ca11ab3a762852
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454843"
---
# <a name="context"></a><span data-ttu-id="de658-102">context</span><span class="sxs-lookup"><span data-stu-id="de658-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="de658-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="de658-103">[Office](Office.md).context</span></span>

<span data-ttu-id="de658-104">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="de658-104">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="de658-105">Esta listagem documenta apenas as interfaces usados pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a referência [Office.context na API Comum](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="de658-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="de658-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="de658-106">Requirements</span></span>

|<span data-ttu-id="de658-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="de658-107">Requirement</span></span>| <span data-ttu-id="de658-108">Valor</span><span class="sxs-lookup"><span data-stu-id="de658-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="de658-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="de658-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de658-110">1.0</span><span class="sxs-lookup"><span data-stu-id="de658-110">1.0</span></span>|
|[<span data-ttu-id="de658-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="de658-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de658-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="de658-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="de658-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="de658-113">Members and methods</span></span>

| <span data-ttu-id="de658-114">Membro</span><span class="sxs-lookup"><span data-stu-id="de658-114">Member</span></span> | <span data-ttu-id="de658-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="de658-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="de658-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="de658-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="de658-117">Membro</span><span class="sxs-lookup"><span data-stu-id="de658-117">Member</span></span> |
| [<span data-ttu-id="de658-118">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="de658-118">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="de658-119">Membro</span><span class="sxs-lookup"><span data-stu-id="de658-119">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="de658-120">Namespaces</span><span class="sxs-lookup"><span data-stu-id="de658-120">Namespaces</span></span>

<span data-ttu-id="de658-121">[Mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto do suplemento do Outlook para o Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="de658-121">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="de658-122">Members</span><span class="sxs-lookup"><span data-stu-id="de658-122">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="de658-123">displayLanguage: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="de658-123">displayLanguage: String</span></span>

<span data-ttu-id="de658-124">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="de658-124">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="de658-125">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="de658-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="de658-126">Tipo</span><span class="sxs-lookup"><span data-stu-id="de658-126">Type</span></span>

*   <span data-ttu-id="de658-127">String</span><span class="sxs-lookup"><span data-stu-id="de658-127">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="de658-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="de658-128">Requirements</span></span>

|<span data-ttu-id="de658-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="de658-129">Requirement</span></span>| <span data-ttu-id="de658-130">Valor</span><span class="sxs-lookup"><span data-stu-id="de658-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="de658-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="de658-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de658-132">1.0</span><span class="sxs-lookup"><span data-stu-id="de658-132">1.0</span></span>|
|[<span data-ttu-id="de658-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="de658-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de658-134">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="de658-134">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="de658-135">Exemplo</span><span class="sxs-lookup"><span data-stu-id="de658-135">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook15officeroamingsettings"></a><span data-ttu-id="de658-136">roamingSettings: [roamingSettings](/javascript/api/outlook_1_5/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="de658-136">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_5/office.RoamingSettings)</span></span>

<span data-ttu-id="de658-137">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="de658-137">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="de658-138">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="de658-138">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="de658-139">Tipo</span><span class="sxs-lookup"><span data-stu-id="de658-139">Type</span></span>

*   [<span data-ttu-id="de658-140">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="de658-140">RoamingSettings</span></span>](/javascript/api/outlook_1_5/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="de658-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="de658-141">Requirements</span></span>

|<span data-ttu-id="de658-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="de658-142">Requirement</span></span>| <span data-ttu-id="de658-143">Valor</span><span class="sxs-lookup"><span data-stu-id="de658-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="de658-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="de658-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de658-145">1.0</span><span class="sxs-lookup"><span data-stu-id="de658-145">1.0</span></span>|
|[<span data-ttu-id="de658-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="de658-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de658-147">Restrito</span><span class="sxs-lookup"><span data-stu-id="de658-147">Restricted</span></span>|
|[<span data-ttu-id="de658-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="de658-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="de658-149">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="de658-149">Compose or Read</span></span>|
