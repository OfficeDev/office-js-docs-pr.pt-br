---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 8ba2a21b16c51c827155d793241b80c5c510dd5a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696334"
---
# <a name="userprofile"></a><span data-ttu-id="89e53-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="89e53-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="89e53-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="89e53-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="89e53-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="89e53-104">Requirements</span></span>

|<span data-ttu-id="89e53-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="89e53-105">Requirement</span></span>| <span data-ttu-id="89e53-106">Valor</span><span class="sxs-lookup"><span data-stu-id="89e53-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="89e53-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="89e53-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="89e53-108">1.0</span><span class="sxs-lookup"><span data-stu-id="89e53-108">1.0</span></span>|
|[<span data-ttu-id="89e53-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="89e53-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="89e53-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="89e53-110">ReadItem</span></span>|
|[<span data-ttu-id="89e53-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="89e53-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="89e53-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="89e53-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="89e53-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="89e53-113">Members and methods</span></span>

| <span data-ttu-id="89e53-114">Membro</span><span class="sxs-lookup"><span data-stu-id="89e53-114">Member</span></span> | <span data-ttu-id="89e53-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="89e53-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="89e53-116">displayName</span><span class="sxs-lookup"><span data-stu-id="89e53-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="89e53-117">Membro</span><span class="sxs-lookup"><span data-stu-id="89e53-117">Member</span></span> |
| [<span data-ttu-id="89e53-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="89e53-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="89e53-119">Membro</span><span class="sxs-lookup"><span data-stu-id="89e53-119">Member</span></span> |
| [<span data-ttu-id="89e53-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="89e53-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="89e53-121">Membro</span><span class="sxs-lookup"><span data-stu-id="89e53-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="89e53-122">Membros</span><span class="sxs-lookup"><span data-stu-id="89e53-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="89e53-123">displayName: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="89e53-123">displayName: String</span></span>

<span data-ttu-id="89e53-124">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="89e53-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="89e53-125">Tipo</span><span class="sxs-lookup"><span data-stu-id="89e53-125">Type</span></span>

*   <span data-ttu-id="89e53-126">String</span><span class="sxs-lookup"><span data-stu-id="89e53-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="89e53-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="89e53-127">Requirements</span></span>

|<span data-ttu-id="89e53-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="89e53-128">Requirement</span></span>| <span data-ttu-id="89e53-129">Valor</span><span class="sxs-lookup"><span data-stu-id="89e53-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="89e53-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="89e53-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="89e53-131">1.0</span><span class="sxs-lookup"><span data-stu-id="89e53-131">1.0</span></span>|
|[<span data-ttu-id="89e53-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="89e53-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="89e53-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="89e53-133">ReadItem</span></span>|
|[<span data-ttu-id="89e53-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="89e53-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="89e53-135">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="89e53-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="89e53-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="89e53-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="89e53-137">emailAddress: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="89e53-137">emailAddress: String</span></span>

<span data-ttu-id="89e53-138">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="89e53-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="89e53-139">Tipo</span><span class="sxs-lookup"><span data-stu-id="89e53-139">Type</span></span>

*   <span data-ttu-id="89e53-140">String</span><span class="sxs-lookup"><span data-stu-id="89e53-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="89e53-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="89e53-141">Requirements</span></span>

|<span data-ttu-id="89e53-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="89e53-142">Requirement</span></span>| <span data-ttu-id="89e53-143">Valor</span><span class="sxs-lookup"><span data-stu-id="89e53-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="89e53-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="89e53-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="89e53-145">1.0</span><span class="sxs-lookup"><span data-stu-id="89e53-145">1.0</span></span>|
|[<span data-ttu-id="89e53-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="89e53-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="89e53-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="89e53-147">ReadItem</span></span>|
|[<span data-ttu-id="89e53-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="89e53-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="89e53-149">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="89e53-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="89e53-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="89e53-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="89e53-151">timeZone: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="89e53-151">timeZone: String</span></span>

<span data-ttu-id="89e53-152">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="89e53-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="89e53-153">Tipo</span><span class="sxs-lookup"><span data-stu-id="89e53-153">Type</span></span>

*   <span data-ttu-id="89e53-154">String</span><span class="sxs-lookup"><span data-stu-id="89e53-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="89e53-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="89e53-155">Requirements</span></span>

|<span data-ttu-id="89e53-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="89e53-156">Requirement</span></span>| <span data-ttu-id="89e53-157">Valor</span><span class="sxs-lookup"><span data-stu-id="89e53-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="89e53-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="89e53-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="89e53-159">1.0</span><span class="sxs-lookup"><span data-stu-id="89e53-159">1.0</span></span>|
|[<span data-ttu-id="89e53-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="89e53-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="89e53-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="89e53-161">ReadItem</span></span>|
|[<span data-ttu-id="89e53-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="89e53-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="89e53-163">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="89e53-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="89e53-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="89e53-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
