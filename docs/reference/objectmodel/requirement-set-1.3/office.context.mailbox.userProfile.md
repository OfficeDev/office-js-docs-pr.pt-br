---
title: 'Office.context.mailbox.userProfile: conjunto de requisitos da versão 1.3'
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 9f36b5f1d31ad6709cf2c43ce7dcb3f91a35bd00
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432218"
---
# <a name="userprofile"></a><span data-ttu-id="8a56a-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="8a56a-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="8a56a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="8a56a-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a56a-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a56a-104">Requirements</span></span>

|<span data-ttu-id="8a56a-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a56a-105">Requirement</span></span>| <span data-ttu-id="8a56a-106">Valor</span><span class="sxs-lookup"><span data-stu-id="8a56a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a56a-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a56a-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a56a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="8a56a-108">1.0</span></span>|
|[<span data-ttu-id="8a56a-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a56a-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a56a-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a56a-110">ReadItem</span></span>|
|[<span data-ttu-id="8a56a-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a56a-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8a56a-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a56a-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="8a56a-113">Membros</span><span class="sxs-lookup"><span data-stu-id="8a56a-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="8a56a-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="8a56a-114">displayName :String</span></span>

<span data-ttu-id="8a56a-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="8a56a-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="8a56a-116">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a56a-116">Type:</span></span>

*   <span data-ttu-id="8a56a-117">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a56a-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a56a-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a56a-118">Requirements</span></span>

|<span data-ttu-id="8a56a-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a56a-119">Requirement</span></span>| <span data-ttu-id="8a56a-120">Valor</span><span class="sxs-lookup"><span data-stu-id="8a56a-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a56a-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a56a-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a56a-122">1.0</span><span class="sxs-lookup"><span data-stu-id="8a56a-122">1.0</span></span>|
|[<span data-ttu-id="8a56a-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a56a-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a56a-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a56a-124">ReadItem</span></span>|
|[<span data-ttu-id="8a56a-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a56a-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8a56a-126">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a56a-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a56a-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a56a-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="8a56a-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="8a56a-128">emailAddress :String</span></span>

<span data-ttu-id="8a56a-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="8a56a-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="8a56a-130">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a56a-130">Type:</span></span>

*   <span data-ttu-id="8a56a-131">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a56a-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a56a-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a56a-132">Requirements</span></span>

|<span data-ttu-id="8a56a-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a56a-133">Requirement</span></span>| <span data-ttu-id="8a56a-134">Valor</span><span class="sxs-lookup"><span data-stu-id="8a56a-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a56a-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a56a-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a56a-136">1.0</span><span class="sxs-lookup"><span data-stu-id="8a56a-136">1.0</span></span>|
|[<span data-ttu-id="8a56a-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a56a-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a56a-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a56a-138">ReadItem</span></span>|
|[<span data-ttu-id="8a56a-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a56a-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8a56a-140">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a56a-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a56a-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a56a-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="8a56a-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="8a56a-142">timeZone :String</span></span>

<span data-ttu-id="8a56a-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="8a56a-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="8a56a-144">Tipo:</span><span class="sxs-lookup"><span data-stu-id="8a56a-144">Type:</span></span>

*   <span data-ttu-id="8a56a-145">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="8a56a-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a56a-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="8a56a-146">Requirements</span></span>

|<span data-ttu-id="8a56a-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="8a56a-147">Requirement</span></span>| <span data-ttu-id="8a56a-148">Valor</span><span class="sxs-lookup"><span data-stu-id="8a56a-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a56a-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="8a56a-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a56a-150">1.0</span><span class="sxs-lookup"><span data-stu-id="8a56a-150">1.0</span></span>|
|[<span data-ttu-id="8a56a-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="8a56a-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a56a-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a56a-152">ReadItem</span></span>|
|[<span data-ttu-id="8a56a-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="8a56a-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8a56a-154">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="8a56a-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a56a-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="8a56a-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```