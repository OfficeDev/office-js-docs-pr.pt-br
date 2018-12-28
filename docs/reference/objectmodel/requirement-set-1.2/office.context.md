---
title: Office.context – conjunto de requisitos 1.2
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: b77e63a21c9ca211a948d05d60701d8c97554981
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433422"
---
# <a name="context"></a><span data-ttu-id="7c441-102">context</span><span class="sxs-lookup"><span data-stu-id="7c441-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="7c441-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="7c441-103">[Office](Office.md).context</span></span>

<span data-ttu-id="7c441-p101">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa do namespace Office.context, confira a [Referência sobre o Office.context na API compartilhada](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="7c441-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="7c441-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c441-106">Requirements</span></span>

|<span data-ttu-id="7c441-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c441-107">Requirement</span></span>| <span data-ttu-id="7c441-108">Valor</span><span class="sxs-lookup"><span data-stu-id="7c441-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c441-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c441-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c441-110">1.0</span><span class="sxs-lookup"><span data-stu-id="7c441-110">1.0</span></span>|
|[<span data-ttu-id="7c441-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c441-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7c441-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="7c441-112">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="7c441-113">Namespaces</span><span class="sxs-lookup"><span data-stu-id="7c441-113">Namespaces</span></span>

<span data-ttu-id="7c441-114">[mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto de Suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="7c441-114">[mailbox](office.context.mailbox.md) - Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="7c441-115">Membros</span><span class="sxs-lookup"><span data-stu-id="7c441-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="7c441-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="7c441-116">displayLanguage :String</span></span>

<span data-ttu-id="7c441-117">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="7c441-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="7c441-118">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="7c441-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="7c441-119">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7c441-119">Type:</span></span>

*   <span data-ttu-id="7c441-120">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c441-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c441-121">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c441-121">Requirements</span></span>

|<span data-ttu-id="7c441-122">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c441-122">Requirement</span></span>| <span data-ttu-id="7c441-123">Valor</span><span class="sxs-lookup"><span data-stu-id="7c441-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c441-124">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c441-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c441-125">1.0</span><span class="sxs-lookup"><span data-stu-id="7c441-125">1.0</span></span>|
|[<span data-ttu-id="7c441-126">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c441-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7c441-127">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="7c441-127">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c441-128">Exemplo</span><span class="sxs-lookup"><span data-stu-id="7c441-128">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook12officeroamingsettings"></a><span data-ttu-id="7c441-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="7c441-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span></span>

<span data-ttu-id="7c441-130">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="7c441-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="7c441-131">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="7c441-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="7c441-132">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7c441-132">Type:</span></span>

*   [<span data-ttu-id="7c441-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="7c441-133">RoamingSettings</span></span>](/javascript/api/outlook_1_2/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="7c441-134">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c441-134">Requirements</span></span>

|<span data-ttu-id="7c441-135">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c441-135">Requirement</span></span>| <span data-ttu-id="7c441-136">Valor</span><span class="sxs-lookup"><span data-stu-id="7c441-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c441-137">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c441-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c441-138">1.0</span><span class="sxs-lookup"><span data-stu-id="7c441-138">1.0</span></span>|
|[<span data-ttu-id="7c441-139">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="7c441-139">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c441-140">Restrito</span><span class="sxs-lookup"><span data-stu-id="7c441-140">Restricted</span></span>|
|[<span data-ttu-id="7c441-141">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c441-141">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7c441-142">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="7c441-142">Compose or read</span></span>|