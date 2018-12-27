---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.2
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: e5548fa514cff9b452c2747324f11e5df8a06def
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432239"
---
# <a name="userprofile"></a><span data-ttu-id="3abfd-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="3abfd-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="3abfd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="3abfd-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="3abfd-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3abfd-104">Requirements</span></span>

|<span data-ttu-id="3abfd-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="3abfd-105">Requirement</span></span>| <span data-ttu-id="3abfd-106">Valor</span><span class="sxs-lookup"><span data-stu-id="3abfd-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="3abfd-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3abfd-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3abfd-108">1.0</span><span class="sxs-lookup"><span data-stu-id="3abfd-108">1.0</span></span>|
|[<span data-ttu-id="3abfd-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3abfd-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3abfd-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3abfd-110">ReadItem</span></span>|
|[<span data-ttu-id="3abfd-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3abfd-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3abfd-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="3abfd-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="3abfd-113">Membros</span><span class="sxs-lookup"><span data-stu-id="3abfd-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="3abfd-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="3abfd-114">displayName :String</span></span>

<span data-ttu-id="3abfd-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="3abfd-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="3abfd-116">Tipo:</span><span class="sxs-lookup"><span data-stu-id="3abfd-116">Type:</span></span>

*   <span data-ttu-id="3abfd-117">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3abfd-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3abfd-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3abfd-118">Requirements</span></span>

|<span data-ttu-id="3abfd-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="3abfd-119">Requirement</span></span>| <span data-ttu-id="3abfd-120">Valor</span><span class="sxs-lookup"><span data-stu-id="3abfd-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="3abfd-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3abfd-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3abfd-122">1.0</span><span class="sxs-lookup"><span data-stu-id="3abfd-122">1.0</span></span>|
|[<span data-ttu-id="3abfd-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3abfd-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3abfd-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3abfd-124">ReadItem</span></span>|
|[<span data-ttu-id="3abfd-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3abfd-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3abfd-126">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="3abfd-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3abfd-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3abfd-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="3abfd-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="3abfd-128">emailAddress :String</span></span>

<span data-ttu-id="3abfd-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="3abfd-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="3abfd-130">Tipo:</span><span class="sxs-lookup"><span data-stu-id="3abfd-130">Type:</span></span>

*   <span data-ttu-id="3abfd-131">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3abfd-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3abfd-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3abfd-132">Requirements</span></span>

|<span data-ttu-id="3abfd-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="3abfd-133">Requirement</span></span>| <span data-ttu-id="3abfd-134">Valor</span><span class="sxs-lookup"><span data-stu-id="3abfd-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="3abfd-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3abfd-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3abfd-136">1.0</span><span class="sxs-lookup"><span data-stu-id="3abfd-136">1.0</span></span>|
|[<span data-ttu-id="3abfd-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3abfd-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3abfd-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3abfd-138">ReadItem</span></span>|
|[<span data-ttu-id="3abfd-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3abfd-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3abfd-140">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="3abfd-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3abfd-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3abfd-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="3abfd-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="3abfd-142">timeZone :String</span></span>

<span data-ttu-id="3abfd-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="3abfd-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="3abfd-144">Tipo:</span><span class="sxs-lookup"><span data-stu-id="3abfd-144">Type:</span></span>

*   <span data-ttu-id="3abfd-145">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3abfd-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3abfd-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3abfd-146">Requirements</span></span>

|<span data-ttu-id="3abfd-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="3abfd-147">Requirement</span></span>| <span data-ttu-id="3abfd-148">Valor</span><span class="sxs-lookup"><span data-stu-id="3abfd-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="3abfd-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3abfd-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3abfd-150">1.0</span><span class="sxs-lookup"><span data-stu-id="3abfd-150">1.0</span></span>|
|[<span data-ttu-id="3abfd-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="3abfd-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3abfd-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3abfd-152">ReadItem</span></span>|
|[<span data-ttu-id="3abfd-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3abfd-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3abfd-154">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="3abfd-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3abfd-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3abfd-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```