---
title: 'Office.context.mailbox.userProfile: conjunto de requisitos da versão 1.4'
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 55d0a789c8e46fd3f6ee69f39cf33f7e7d94c322
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432645"
---
# <a name="userprofile"></a><span data-ttu-id="59476-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="59476-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="59476-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="59476-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="59476-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="59476-104">Requirements</span></span>

|<span data-ttu-id="59476-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="59476-105">Requirement</span></span>| <span data-ttu-id="59476-106">Valor</span><span class="sxs-lookup"><span data-stu-id="59476-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="59476-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="59476-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="59476-108">1.0</span><span class="sxs-lookup"><span data-stu-id="59476-108">1.0</span></span>|
|[<span data-ttu-id="59476-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="59476-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="59476-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59476-110">ReadItem</span></span>|
|[<span data-ttu-id="59476-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="59476-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="59476-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="59476-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="59476-113">Membros</span><span class="sxs-lookup"><span data-stu-id="59476-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="59476-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="59476-114">displayName :String</span></span>

<span data-ttu-id="59476-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="59476-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="59476-116">Tipo:</span><span class="sxs-lookup"><span data-stu-id="59476-116">Type:</span></span>

*   <span data-ttu-id="59476-117">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="59476-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59476-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="59476-118">Requirements</span></span>

|<span data-ttu-id="59476-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="59476-119">Requirement</span></span>| <span data-ttu-id="59476-120">Valor</span><span class="sxs-lookup"><span data-stu-id="59476-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="59476-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="59476-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="59476-122">1.0</span><span class="sxs-lookup"><span data-stu-id="59476-122">1.0</span></span>|
|[<span data-ttu-id="59476-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="59476-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="59476-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59476-124">ReadItem</span></span>|
|[<span data-ttu-id="59476-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="59476-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="59476-126">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="59476-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="59476-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="59476-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="59476-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="59476-128">emailAddress :String</span></span>

<span data-ttu-id="59476-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="59476-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="59476-130">Tipo:</span><span class="sxs-lookup"><span data-stu-id="59476-130">Type:</span></span>

*   <span data-ttu-id="59476-131">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="59476-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59476-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="59476-132">Requirements</span></span>

|<span data-ttu-id="59476-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="59476-133">Requirement</span></span>| <span data-ttu-id="59476-134">Valor</span><span class="sxs-lookup"><span data-stu-id="59476-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="59476-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="59476-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="59476-136">1.0</span><span class="sxs-lookup"><span data-stu-id="59476-136">1.0</span></span>|
|[<span data-ttu-id="59476-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="59476-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="59476-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59476-138">ReadItem</span></span>|
|[<span data-ttu-id="59476-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="59476-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="59476-140">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="59476-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="59476-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="59476-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="59476-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="59476-142">timeZone :String</span></span>

<span data-ttu-id="59476-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="59476-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="59476-144">Tipo:</span><span class="sxs-lookup"><span data-stu-id="59476-144">Type:</span></span>

*   <span data-ttu-id="59476-145">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="59476-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59476-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="59476-146">Requirements</span></span>

|<span data-ttu-id="59476-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="59476-147">Requirement</span></span>| <span data-ttu-id="59476-148">Valor</span><span class="sxs-lookup"><span data-stu-id="59476-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="59476-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="59476-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="59476-150">1.0</span><span class="sxs-lookup"><span data-stu-id="59476-150">1.0</span></span>|
|[<span data-ttu-id="59476-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="59476-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="59476-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59476-152">ReadItem</span></span>|
|[<span data-ttu-id="59476-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="59476-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="59476-154">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="59476-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="59476-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="59476-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```