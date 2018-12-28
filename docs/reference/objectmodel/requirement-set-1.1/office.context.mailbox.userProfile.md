---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.1
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 312cba4d5aace980b7c9b205899fac51d3da3de5
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433170"
---
# <a name="userprofile"></a><span data-ttu-id="912f9-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="912f9-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="912f9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="912f9-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="912f9-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="912f9-104">Requirements</span></span>

|<span data-ttu-id="912f9-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="912f9-105">Requirement</span></span>| <span data-ttu-id="912f9-106">Valor</span><span class="sxs-lookup"><span data-stu-id="912f9-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="912f9-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="912f9-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="912f9-108">1.0</span><span class="sxs-lookup"><span data-stu-id="912f9-108">1.0</span></span>|
|[<span data-ttu-id="912f9-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="912f9-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="912f9-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="912f9-110">ReadItem</span></span>|
|[<span data-ttu-id="912f9-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="912f9-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="912f9-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="912f9-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="912f9-113">Membros</span><span class="sxs-lookup"><span data-stu-id="912f9-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="912f9-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="912f9-114">displayName :String</span></span>

<span data-ttu-id="912f9-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="912f9-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="912f9-116">Tipo:</span><span class="sxs-lookup"><span data-stu-id="912f9-116">Type:</span></span>

*   <span data-ttu-id="912f9-117">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="912f9-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="912f9-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="912f9-118">Requirements</span></span>

|<span data-ttu-id="912f9-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="912f9-119">Requirement</span></span>| <span data-ttu-id="912f9-120">Valor</span><span class="sxs-lookup"><span data-stu-id="912f9-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="912f9-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="912f9-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="912f9-122">1.0</span><span class="sxs-lookup"><span data-stu-id="912f9-122">1.0</span></span>|
|[<span data-ttu-id="912f9-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="912f9-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="912f9-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="912f9-124">ReadItem</span></span>|
|[<span data-ttu-id="912f9-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="912f9-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="912f9-126">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="912f9-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="912f9-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="912f9-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="912f9-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="912f9-128">emailAddress :String</span></span>

<span data-ttu-id="912f9-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="912f9-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="912f9-130">Tipo:</span><span class="sxs-lookup"><span data-stu-id="912f9-130">Type:</span></span>

*   <span data-ttu-id="912f9-131">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="912f9-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="912f9-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="912f9-132">Requirements</span></span>

|<span data-ttu-id="912f9-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="912f9-133">Requirement</span></span>| <span data-ttu-id="912f9-134">Valor</span><span class="sxs-lookup"><span data-stu-id="912f9-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="912f9-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="912f9-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="912f9-136">1.0</span><span class="sxs-lookup"><span data-stu-id="912f9-136">1.0</span></span>|
|[<span data-ttu-id="912f9-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="912f9-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="912f9-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="912f9-138">ReadItem</span></span>|
|[<span data-ttu-id="912f9-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="912f9-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="912f9-140">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="912f9-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="912f9-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="912f9-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="912f9-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="912f9-142">timeZone :String</span></span>

<span data-ttu-id="912f9-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="912f9-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="912f9-144">Tipo:</span><span class="sxs-lookup"><span data-stu-id="912f9-144">Type:</span></span>

*   <span data-ttu-id="912f9-145">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="912f9-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="912f9-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="912f9-146">Requirements</span></span>

|<span data-ttu-id="912f9-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="912f9-147">Requirement</span></span>| <span data-ttu-id="912f9-148">Valor</span><span class="sxs-lookup"><span data-stu-id="912f9-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="912f9-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="912f9-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="912f9-150">1.0</span><span class="sxs-lookup"><span data-stu-id="912f9-150">1.0</span></span>|
|[<span data-ttu-id="912f9-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="912f9-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="912f9-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="912f9-152">ReadItem</span></span>|
|[<span data-ttu-id="912f9-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="912f9-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="912f9-154">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="912f9-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="912f9-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="912f9-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```