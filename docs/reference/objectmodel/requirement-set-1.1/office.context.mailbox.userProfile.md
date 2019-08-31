---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 06492623e0b9ab16792d6b23dfaeb27d99125ff1
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696397"
---
# <a name="userprofile"></a><span data-ttu-id="92934-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="92934-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="92934-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="92934-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="92934-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="92934-104">Requirements</span></span>

|<span data-ttu-id="92934-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="92934-105">Requirement</span></span>| <span data-ttu-id="92934-106">Valor</span><span class="sxs-lookup"><span data-stu-id="92934-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="92934-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="92934-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92934-108">1.0</span><span class="sxs-lookup"><span data-stu-id="92934-108">1.0</span></span>|
|[<span data-ttu-id="92934-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="92934-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92934-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92934-110">ReadItem</span></span>|
|[<span data-ttu-id="92934-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="92934-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92934-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="92934-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="92934-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="92934-113">Members and methods</span></span>

| <span data-ttu-id="92934-114">Membro</span><span class="sxs-lookup"><span data-stu-id="92934-114">Member</span></span> | <span data-ttu-id="92934-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="92934-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="92934-116">displayName</span><span class="sxs-lookup"><span data-stu-id="92934-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="92934-117">Membro</span><span class="sxs-lookup"><span data-stu-id="92934-117">Member</span></span> |
| [<span data-ttu-id="92934-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="92934-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="92934-119">Membro</span><span class="sxs-lookup"><span data-stu-id="92934-119">Member</span></span> |
| [<span data-ttu-id="92934-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="92934-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="92934-121">Membro</span><span class="sxs-lookup"><span data-stu-id="92934-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="92934-122">Membros</span><span class="sxs-lookup"><span data-stu-id="92934-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="92934-123">displayName: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="92934-123">displayName: String</span></span>

<span data-ttu-id="92934-124">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="92934-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="92934-125">Tipo</span><span class="sxs-lookup"><span data-stu-id="92934-125">Type</span></span>

*   <span data-ttu-id="92934-126">String</span><span class="sxs-lookup"><span data-stu-id="92934-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92934-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="92934-127">Requirements</span></span>

|<span data-ttu-id="92934-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="92934-128">Requirement</span></span>| <span data-ttu-id="92934-129">Valor</span><span class="sxs-lookup"><span data-stu-id="92934-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="92934-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="92934-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92934-131">1.0</span><span class="sxs-lookup"><span data-stu-id="92934-131">1.0</span></span>|
|[<span data-ttu-id="92934-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="92934-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92934-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92934-133">ReadItem</span></span>|
|[<span data-ttu-id="92934-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="92934-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92934-135">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="92934-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92934-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="92934-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="92934-137">emailAddress: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="92934-137">emailAddress: String</span></span>

<span data-ttu-id="92934-138">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="92934-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="92934-139">Tipo</span><span class="sxs-lookup"><span data-stu-id="92934-139">Type</span></span>

*   <span data-ttu-id="92934-140">String</span><span class="sxs-lookup"><span data-stu-id="92934-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92934-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="92934-141">Requirements</span></span>

|<span data-ttu-id="92934-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="92934-142">Requirement</span></span>| <span data-ttu-id="92934-143">Valor</span><span class="sxs-lookup"><span data-stu-id="92934-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="92934-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="92934-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92934-145">1.0</span><span class="sxs-lookup"><span data-stu-id="92934-145">1.0</span></span>|
|[<span data-ttu-id="92934-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="92934-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92934-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92934-147">ReadItem</span></span>|
|[<span data-ttu-id="92934-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="92934-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92934-149">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="92934-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92934-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="92934-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="92934-151">timeZone: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="92934-151">timeZone: String</span></span>

<span data-ttu-id="92934-152">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="92934-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="92934-153">Tipo</span><span class="sxs-lookup"><span data-stu-id="92934-153">Type</span></span>

*   <span data-ttu-id="92934-154">String</span><span class="sxs-lookup"><span data-stu-id="92934-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92934-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="92934-155">Requirements</span></span>

|<span data-ttu-id="92934-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="92934-156">Requirement</span></span>| <span data-ttu-id="92934-157">Valor</span><span class="sxs-lookup"><span data-stu-id="92934-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="92934-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="92934-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92934-159">1.0</span><span class="sxs-lookup"><span data-stu-id="92934-159">1.0</span></span>|
|[<span data-ttu-id="92934-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="92934-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92934-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92934-161">ReadItem</span></span>|
|[<span data-ttu-id="92934-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="92934-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92934-163">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="92934-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92934-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="92934-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
