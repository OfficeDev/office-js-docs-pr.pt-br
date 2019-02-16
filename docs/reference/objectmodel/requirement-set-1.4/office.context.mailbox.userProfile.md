---
title: 'Office.context.mailbox.userProfile: conjunto de requisitos da versão 1.4'
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 7facc0ea555dca7d6784a09f798c3d8fa25f2731
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067843"
---
# <a name="userprofile"></a><span data-ttu-id="1d61f-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="1d61f-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="1d61f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="1d61f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d61f-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d61f-104">Requirements</span></span>

|<span data-ttu-id="1d61f-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d61f-105">Requirement</span></span>| <span data-ttu-id="1d61f-106">Valor</span><span class="sxs-lookup"><span data-stu-id="1d61f-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d61f-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d61f-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d61f-108">1.0</span><span class="sxs-lookup"><span data-stu-id="1d61f-108">1.0</span></span>|
|[<span data-ttu-id="1d61f-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d61f-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d61f-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d61f-110">ReadItem</span></span>|
|[<span data-ttu-id="1d61f-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d61f-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d61f-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d61f-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="1d61f-113">Membros</span><span class="sxs-lookup"><span data-stu-id="1d61f-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="1d61f-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="1d61f-114">displayName :String</span></span>

<span data-ttu-id="1d61f-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="1d61f-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="1d61f-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d61f-116">Type</span></span>

*   <span data-ttu-id="1d61f-117">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d61f-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d61f-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d61f-118">Requirements</span></span>

|<span data-ttu-id="1d61f-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d61f-119">Requirement</span></span>| <span data-ttu-id="1d61f-120">Valor</span><span class="sxs-lookup"><span data-stu-id="1d61f-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d61f-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d61f-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d61f-122">1.0</span><span class="sxs-lookup"><span data-stu-id="1d61f-122">1.0</span></span>|
|[<span data-ttu-id="1d61f-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d61f-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d61f-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d61f-124">ReadItem</span></span>|
|[<span data-ttu-id="1d61f-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d61f-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d61f-126">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d61f-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d61f-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d61f-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="1d61f-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="1d61f-128">emailAddress :String</span></span>

<span data-ttu-id="1d61f-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="1d61f-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="1d61f-130">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d61f-130">Type</span></span>

*   <span data-ttu-id="1d61f-131">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d61f-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d61f-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d61f-132">Requirements</span></span>

|<span data-ttu-id="1d61f-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d61f-133">Requirement</span></span>| <span data-ttu-id="1d61f-134">Valor</span><span class="sxs-lookup"><span data-stu-id="1d61f-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d61f-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d61f-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d61f-136">1.0</span><span class="sxs-lookup"><span data-stu-id="1d61f-136">1.0</span></span>|
|[<span data-ttu-id="1d61f-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d61f-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d61f-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d61f-138">ReadItem</span></span>|
|[<span data-ttu-id="1d61f-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d61f-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d61f-140">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d61f-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d61f-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d61f-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="1d61f-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="1d61f-142">timeZone :String</span></span>

<span data-ttu-id="1d61f-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="1d61f-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="1d61f-144">Tipo</span><span class="sxs-lookup"><span data-stu-id="1d61f-144">Type</span></span>

*   <span data-ttu-id="1d61f-145">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1d61f-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1d61f-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1d61f-146">Requirements</span></span>

|<span data-ttu-id="1d61f-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="1d61f-147">Requirement</span></span>| <span data-ttu-id="1d61f-148">Valor</span><span class="sxs-lookup"><span data-stu-id="1d61f-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="1d61f-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1d61f-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1d61f-150">1.0</span><span class="sxs-lookup"><span data-stu-id="1d61f-150">1.0</span></span>|
|[<span data-ttu-id="1d61f-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="1d61f-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1d61f-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1d61f-152">ReadItem</span></span>|
|[<span data-ttu-id="1d61f-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1d61f-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1d61f-154">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="1d61f-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1d61f-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="1d61f-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
