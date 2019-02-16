---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.5
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: e98e88cde184db121e69fdd267dff4e39d887b1f
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067823"
---
# <a name="userprofile"></a><span data-ttu-id="b5523-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="b5523-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="b5523-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="b5523-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="b5523-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b5523-104">Requirements</span></span>

|<span data-ttu-id="b5523-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="b5523-105">Requirement</span></span>| <span data-ttu-id="b5523-106">Valor</span><span class="sxs-lookup"><span data-stu-id="b5523-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b5523-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b5523-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b5523-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b5523-108">1.0</span></span>|
|[<span data-ttu-id="b5523-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b5523-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b5523-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b5523-110">ReadItem</span></span>|
|[<span data-ttu-id="b5523-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b5523-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b5523-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b5523-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b5523-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="b5523-113">Members and methods</span></span>

| <span data-ttu-id="b5523-114">Membro</span><span class="sxs-lookup"><span data-stu-id="b5523-114">Member</span></span> | <span data-ttu-id="b5523-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="b5523-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b5523-116">displayName</span><span class="sxs-lookup"><span data-stu-id="b5523-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="b5523-117">Membro</span><span class="sxs-lookup"><span data-stu-id="b5523-117">Member</span></span> |
| [<span data-ttu-id="b5523-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="b5523-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="b5523-119">Membro</span><span class="sxs-lookup"><span data-stu-id="b5523-119">Member</span></span> |
| [<span data-ttu-id="b5523-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="b5523-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="b5523-121">Membro</span><span class="sxs-lookup"><span data-stu-id="b5523-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="b5523-122">Membros</span><span class="sxs-lookup"><span data-stu-id="b5523-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="b5523-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="b5523-123">displayName :String</span></span>

<span data-ttu-id="b5523-124">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="b5523-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="b5523-125">Tipo</span><span class="sxs-lookup"><span data-stu-id="b5523-125">Type</span></span>

*   <span data-ttu-id="b5523-126">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b5523-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b5523-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b5523-127">Requirements</span></span>

|<span data-ttu-id="b5523-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="b5523-128">Requirement</span></span>| <span data-ttu-id="b5523-129">Valor</span><span class="sxs-lookup"><span data-stu-id="b5523-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="b5523-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b5523-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b5523-131">1.0</span><span class="sxs-lookup"><span data-stu-id="b5523-131">1.0</span></span>|
|[<span data-ttu-id="b5523-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b5523-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b5523-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b5523-133">ReadItem</span></span>|
|[<span data-ttu-id="b5523-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b5523-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b5523-135">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b5523-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b5523-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b5523-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="b5523-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="b5523-137">emailAddress :String</span></span>

<span data-ttu-id="b5523-138">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="b5523-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="b5523-139">Tipo</span><span class="sxs-lookup"><span data-stu-id="b5523-139">Type</span></span>

*   <span data-ttu-id="b5523-140">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b5523-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b5523-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b5523-141">Requirements</span></span>

|<span data-ttu-id="b5523-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="b5523-142">Requirement</span></span>| <span data-ttu-id="b5523-143">Valor</span><span class="sxs-lookup"><span data-stu-id="b5523-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="b5523-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b5523-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b5523-145">1.0</span><span class="sxs-lookup"><span data-stu-id="b5523-145">1.0</span></span>|
|[<span data-ttu-id="b5523-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b5523-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b5523-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b5523-147">ReadItem</span></span>|
|[<span data-ttu-id="b5523-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b5523-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b5523-149">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b5523-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b5523-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b5523-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="b5523-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="b5523-151">timeZone :String</span></span>

<span data-ttu-id="b5523-152">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="b5523-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="b5523-153">Tipo</span><span class="sxs-lookup"><span data-stu-id="b5523-153">Type</span></span>

*   <span data-ttu-id="b5523-154">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b5523-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b5523-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b5523-155">Requirements</span></span>

|<span data-ttu-id="b5523-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="b5523-156">Requirement</span></span>| <span data-ttu-id="b5523-157">Valor</span><span class="sxs-lookup"><span data-stu-id="b5523-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="b5523-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b5523-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b5523-159">1.0</span><span class="sxs-lookup"><span data-stu-id="b5523-159">1.0</span></span>|
|[<span data-ttu-id="b5523-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="b5523-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b5523-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b5523-161">ReadItem</span></span>|
|[<span data-ttu-id="b5523-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b5523-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b5523-163">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="b5523-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b5523-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="b5523-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
