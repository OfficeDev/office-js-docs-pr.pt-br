---
title: 'Office.context.mailbox.userProfile: conjunto de requisitos da versão 1.3'
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: e29cf90d1c5d4c288417ef98f6e9d22eaf908b67
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067927"
---
# <a name="userprofile"></a><span data-ttu-id="b13d0-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="b13d0-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="b13d0-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="b13d0-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="b13d0-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b13d0-104">Requirements</span></span>

|<span data-ttu-id="b13d0-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="b13d0-105">Requirement</span></span>| <span data-ttu-id="b13d0-106">Valor</span><span class="sxs-lookup"><span data-stu-id="b13d0-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b13d0-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b13d0-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b13d0-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b13d0-108">1.0</span></span>|
|[<span data-ttu-id="b13d0-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b13d0-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b13d0-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b13d0-110">ReadItem</span></span>|
|[<span data-ttu-id="b13d0-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b13d0-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b13d0-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b13d0-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="b13d0-113">Membros</span><span class="sxs-lookup"><span data-stu-id="b13d0-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="b13d0-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="b13d0-114">displayName :String</span></span>

<span data-ttu-id="b13d0-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="b13d0-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="b13d0-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="b13d0-116">Type</span></span>

*   <span data-ttu-id="b13d0-117">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b13d0-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b13d0-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b13d0-118">Requirements</span></span>

|<span data-ttu-id="b13d0-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="b13d0-119">Requirement</span></span>| <span data-ttu-id="b13d0-120">Valor</span><span class="sxs-lookup"><span data-stu-id="b13d0-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="b13d0-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b13d0-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b13d0-122">1.0</span><span class="sxs-lookup"><span data-stu-id="b13d0-122">1.0</span></span>|
|[<span data-ttu-id="b13d0-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b13d0-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b13d0-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b13d0-124">ReadItem</span></span>|
|[<span data-ttu-id="b13d0-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b13d0-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b13d0-126">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b13d0-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b13d0-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b13d0-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="b13d0-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="b13d0-128">emailAddress :String</span></span>

<span data-ttu-id="b13d0-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="b13d0-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="b13d0-130">Tipo</span><span class="sxs-lookup"><span data-stu-id="b13d0-130">Type</span></span>

*   <span data-ttu-id="b13d0-131">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b13d0-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b13d0-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b13d0-132">Requirements</span></span>

|<span data-ttu-id="b13d0-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="b13d0-133">Requirement</span></span>| <span data-ttu-id="b13d0-134">Valor</span><span class="sxs-lookup"><span data-stu-id="b13d0-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="b13d0-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b13d0-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b13d0-136">1.0</span><span class="sxs-lookup"><span data-stu-id="b13d0-136">1.0</span></span>|
|[<span data-ttu-id="b13d0-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b13d0-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b13d0-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b13d0-138">ReadItem</span></span>|
|[<span data-ttu-id="b13d0-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b13d0-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b13d0-140">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b13d0-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b13d0-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b13d0-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="b13d0-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="b13d0-142">timeZone :String</span></span>

<span data-ttu-id="b13d0-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="b13d0-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="b13d0-144">Tipo</span><span class="sxs-lookup"><span data-stu-id="b13d0-144">Type</span></span>

*   <span data-ttu-id="b13d0-145">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b13d0-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b13d0-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b13d0-146">Requirements</span></span>

|<span data-ttu-id="b13d0-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="b13d0-147">Requirement</span></span>| <span data-ttu-id="b13d0-148">Valor</span><span class="sxs-lookup"><span data-stu-id="b13d0-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="b13d0-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b13d0-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b13d0-150">1.0</span><span class="sxs-lookup"><span data-stu-id="b13d0-150">1.0</span></span>|
|[<span data-ttu-id="b13d0-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b13d0-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b13d0-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b13d0-152">ReadItem</span></span>|
|[<span data-ttu-id="b13d0-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b13d0-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b13d0-154">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b13d0-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b13d0-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b13d0-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
