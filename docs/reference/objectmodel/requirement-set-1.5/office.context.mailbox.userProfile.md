---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,5
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 993fad674fcc616483ac927619e7ca64d81b7326
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696089"
---
# <a name="userprofile"></a><span data-ttu-id="2a22a-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="2a22a-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="2a22a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="2a22a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a22a-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2a22a-104">Requirements</span></span>

|<span data-ttu-id="2a22a-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="2a22a-105">Requirement</span></span>| <span data-ttu-id="2a22a-106">Valor</span><span class="sxs-lookup"><span data-stu-id="2a22a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a22a-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2a22a-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2a22a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="2a22a-108">1.0</span></span>|
|[<span data-ttu-id="2a22a-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2a22a-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2a22a-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a22a-110">ReadItem</span></span>|
|[<span data-ttu-id="2a22a-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2a22a-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a22a-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2a22a-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2a22a-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="2a22a-113">Members and methods</span></span>

| <span data-ttu-id="2a22a-114">Membro</span><span class="sxs-lookup"><span data-stu-id="2a22a-114">Member</span></span> | <span data-ttu-id="2a22a-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="2a22a-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2a22a-116">displayName</span><span class="sxs-lookup"><span data-stu-id="2a22a-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="2a22a-117">Membro</span><span class="sxs-lookup"><span data-stu-id="2a22a-117">Member</span></span> |
| [<span data-ttu-id="2a22a-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="2a22a-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="2a22a-119">Membro</span><span class="sxs-lookup"><span data-stu-id="2a22a-119">Member</span></span> |
| [<span data-ttu-id="2a22a-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="2a22a-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="2a22a-121">Membro</span><span class="sxs-lookup"><span data-stu-id="2a22a-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="2a22a-122">Membros</span><span class="sxs-lookup"><span data-stu-id="2a22a-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="2a22a-123">displayName: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2a22a-123">displayName: String</span></span>

<span data-ttu-id="2a22a-124">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="2a22a-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="2a22a-125">Tipo</span><span class="sxs-lookup"><span data-stu-id="2a22a-125">Type</span></span>

*   <span data-ttu-id="2a22a-126">String</span><span class="sxs-lookup"><span data-stu-id="2a22a-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a22a-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2a22a-127">Requirements</span></span>

|<span data-ttu-id="2a22a-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="2a22a-128">Requirement</span></span>| <span data-ttu-id="2a22a-129">Valor</span><span class="sxs-lookup"><span data-stu-id="2a22a-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a22a-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2a22a-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2a22a-131">1.0</span><span class="sxs-lookup"><span data-stu-id="2a22a-131">1.0</span></span>|
|[<span data-ttu-id="2a22a-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2a22a-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2a22a-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a22a-133">ReadItem</span></span>|
|[<span data-ttu-id="2a22a-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2a22a-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a22a-135">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2a22a-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a22a-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2a22a-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="2a22a-137">emailAddress: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2a22a-137">emailAddress: String</span></span>

<span data-ttu-id="2a22a-138">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="2a22a-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="2a22a-139">Tipo</span><span class="sxs-lookup"><span data-stu-id="2a22a-139">Type</span></span>

*   <span data-ttu-id="2a22a-140">String</span><span class="sxs-lookup"><span data-stu-id="2a22a-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a22a-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2a22a-141">Requirements</span></span>

|<span data-ttu-id="2a22a-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="2a22a-142">Requirement</span></span>| <span data-ttu-id="2a22a-143">Valor</span><span class="sxs-lookup"><span data-stu-id="2a22a-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a22a-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2a22a-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2a22a-145">1.0</span><span class="sxs-lookup"><span data-stu-id="2a22a-145">1.0</span></span>|
|[<span data-ttu-id="2a22a-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2a22a-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2a22a-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a22a-147">ReadItem</span></span>|
|[<span data-ttu-id="2a22a-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2a22a-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a22a-149">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2a22a-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a22a-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2a22a-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="2a22a-151">timeZone: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2a22a-151">timeZone: String</span></span>

<span data-ttu-id="2a22a-152">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="2a22a-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="2a22a-153">Tipo</span><span class="sxs-lookup"><span data-stu-id="2a22a-153">Type</span></span>

*   <span data-ttu-id="2a22a-154">String</span><span class="sxs-lookup"><span data-stu-id="2a22a-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a22a-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2a22a-155">Requirements</span></span>

|<span data-ttu-id="2a22a-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="2a22a-156">Requirement</span></span>| <span data-ttu-id="2a22a-157">Valor</span><span class="sxs-lookup"><span data-stu-id="2a22a-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a22a-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2a22a-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2a22a-159">1.0</span><span class="sxs-lookup"><span data-stu-id="2a22a-159">1.0</span></span>|
|[<span data-ttu-id="2a22a-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2a22a-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2a22a-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2a22a-161">ReadItem</span></span>|
|[<span data-ttu-id="2a22a-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2a22a-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a22a-163">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2a22a-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2a22a-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2a22a-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
