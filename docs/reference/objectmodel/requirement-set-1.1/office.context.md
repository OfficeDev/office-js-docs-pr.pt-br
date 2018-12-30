---
title: 'Office.context: conjunto de requisitos da versão 1.1'
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: a736588233db33b04e64c517f4b0802c17084d99
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457982"
---
# <a name="context"></a><span data-ttu-id="9966f-102">context</span><span class="sxs-lookup"><span data-stu-id="9966f-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="9966f-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="9966f-103">[Office](Office.md).context</span></span>

<span data-ttu-id="9966f-104">O namespace Office.context fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="9966f-104">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="9966f-105">Esta listagem documenta apenas as interfaces usadas pelos suplementos do Outlook. Para uma listagem completa do namespace Office.context, veja a [referência Office.context no API Comum](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="9966f-105">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="9966f-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9966f-106">Requirements</span></span>

|<span data-ttu-id="9966f-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="9966f-107">Requirement</span></span>| <span data-ttu-id="9966f-108">Valor</span><span class="sxs-lookup"><span data-stu-id="9966f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="9966f-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9966f-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9966f-110">1.0</span><span class="sxs-lookup"><span data-stu-id="9966f-110">1.0</span></span>|
|[<span data-ttu-id="9966f-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9966f-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9966f-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9966f-112">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="9966f-113">Namespaces</span><span class="sxs-lookup"><span data-stu-id="9966f-113">Namespaces</span></span>

<span data-ttu-id="9966f-114">[mailbox](office.context.mailbox.md): fornece acesso ao modelo de objeto de suplemento do Outlook para o Microsoft Outlook e o Microsoft Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="9966f-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="9966f-115">Membros</span><span class="sxs-lookup"><span data-stu-id="9966f-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="9966f-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="9966f-116">displayLanguage :String</span></span>

<span data-ttu-id="9966f-117">Obtém a localidade (idioma) no formato de marca de idioma RFC 1766 especificado pelo usuário para a interface do usuário do aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="9966f-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="9966f-118">O valor `displayLanguage` reflete a configuração atual de **Display Language** especificada com **Arquivo > Opções > Idioma** no aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="9966f-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="9966f-119">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9966f-119">Type:</span></span>

*   <span data-ttu-id="9966f-120">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9966f-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9966f-121">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9966f-121">Requirements</span></span>

|<span data-ttu-id="9966f-122">Requisito</span><span class="sxs-lookup"><span data-stu-id="9966f-122">Requirement</span></span>| <span data-ttu-id="9966f-123">Valor</span><span class="sxs-lookup"><span data-stu-id="9966f-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="9966f-124">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9966f-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9966f-125">1.0</span><span class="sxs-lookup"><span data-stu-id="9966f-125">1.0</span></span>|
|[<span data-ttu-id="9966f-126">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9966f-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9966f-127">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9966f-127">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9966f-128">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9966f-128">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook11officeroamingsettings"></a><span data-ttu-id="9966f-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="9966f-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)</span></span>

<span data-ttu-id="9966f-130">Obtém um objeto que representa as configurações personalizadas ou o estado de um suplemento de email do Outlook salvos na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="9966f-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="9966f-131">O objeto `RoamingSettings` permite armazenar e acessar os dados de um suplemento de email que está armazenado na caixa de correio do usuário, para que fiquem disponíveis para esse suplemento quando ele for executado em qualquer aplicativo host de cliente usado para acessar essa caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="9966f-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="9966f-132">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9966f-132">Type:</span></span>

*   [<span data-ttu-id="9966f-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9966f-133">RoamingSettings</span></span>](/javascript/api/outlook_1_1/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="9966f-134">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9966f-134">Requirements</span></span>

|<span data-ttu-id="9966f-135">Requisito</span><span class="sxs-lookup"><span data-stu-id="9966f-135">Requirement</span></span>| <span data-ttu-id="9966f-136">Valor</span><span class="sxs-lookup"><span data-stu-id="9966f-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="9966f-137">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9966f-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9966f-138">1.0</span><span class="sxs-lookup"><span data-stu-id="9966f-138">1.0</span></span>|
|[<span data-ttu-id="9966f-139">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9966f-139">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9966f-140">Restrito</span><span class="sxs-lookup"><span data-stu-id="9966f-140">Restricted</span></span>|
|[<span data-ttu-id="9966f-141">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9966f-141">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9966f-142">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="9966f-142">Compose or read</span></span>|