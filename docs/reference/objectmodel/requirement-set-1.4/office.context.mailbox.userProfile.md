---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,4
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 7a728ebbec0136e0b2eddfb4402e45abe3f02ad4
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268632"
---
# <a name="userprofile"></a><span data-ttu-id="6e0de-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="6e0de-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="6e0de-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="6e0de-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e0de-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6e0de-104">Requirements</span></span>

|<span data-ttu-id="6e0de-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="6e0de-105">Requirement</span></span>| <span data-ttu-id="6e0de-106">Valor</span><span class="sxs-lookup"><span data-stu-id="6e0de-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e0de-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6e0de-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e0de-108">1.0</span><span class="sxs-lookup"><span data-stu-id="6e0de-108">1.0</span></span>|
|[<span data-ttu-id="6e0de-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6e0de-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e0de-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e0de-110">ReadItem</span></span>|
|[<span data-ttu-id="6e0de-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6e0de-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e0de-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6e0de-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6e0de-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="6e0de-113">Members and methods</span></span>

| <span data-ttu-id="6e0de-114">Membro</span><span class="sxs-lookup"><span data-stu-id="6e0de-114">Member</span></span> | <span data-ttu-id="6e0de-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="6e0de-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6e0de-116">displayName</span><span class="sxs-lookup"><span data-stu-id="6e0de-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="6e0de-117">Membro</span><span class="sxs-lookup"><span data-stu-id="6e0de-117">Member</span></span> |
| [<span data-ttu-id="6e0de-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="6e0de-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="6e0de-119">Membro</span><span class="sxs-lookup"><span data-stu-id="6e0de-119">Member</span></span> |
| [<span data-ttu-id="6e0de-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="6e0de-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="6e0de-121">Membro</span><span class="sxs-lookup"><span data-stu-id="6e0de-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="6e0de-122">Membros</span><span class="sxs-lookup"><span data-stu-id="6e0de-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="6e0de-123">displayName: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6e0de-123">displayName: String</span></span>

<span data-ttu-id="6e0de-124">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="6e0de-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="6e0de-125">Tipo</span><span class="sxs-lookup"><span data-stu-id="6e0de-125">Type</span></span>

*   <span data-ttu-id="6e0de-126">String</span><span class="sxs-lookup"><span data-stu-id="6e0de-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e0de-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6e0de-127">Requirements</span></span>

|<span data-ttu-id="6e0de-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="6e0de-128">Requirement</span></span>| <span data-ttu-id="6e0de-129">Valor</span><span class="sxs-lookup"><span data-stu-id="6e0de-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e0de-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6e0de-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e0de-131">1.0</span><span class="sxs-lookup"><span data-stu-id="6e0de-131">1.0</span></span>|
|[<span data-ttu-id="6e0de-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6e0de-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e0de-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e0de-133">ReadItem</span></span>|
|[<span data-ttu-id="6e0de-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6e0de-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e0de-135">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6e0de-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e0de-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6e0de-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="6e0de-137">emailAddress: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6e0de-137">emailAddress: String</span></span>

<span data-ttu-id="6e0de-138">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="6e0de-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="6e0de-139">Tipo</span><span class="sxs-lookup"><span data-stu-id="6e0de-139">Type</span></span>

*   <span data-ttu-id="6e0de-140">String</span><span class="sxs-lookup"><span data-stu-id="6e0de-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e0de-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6e0de-141">Requirements</span></span>

|<span data-ttu-id="6e0de-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="6e0de-142">Requirement</span></span>| <span data-ttu-id="6e0de-143">Valor</span><span class="sxs-lookup"><span data-stu-id="6e0de-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e0de-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6e0de-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e0de-145">1.0</span><span class="sxs-lookup"><span data-stu-id="6e0de-145">1.0</span></span>|
|[<span data-ttu-id="6e0de-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6e0de-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e0de-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e0de-147">ReadItem</span></span>|
|[<span data-ttu-id="6e0de-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6e0de-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e0de-149">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6e0de-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e0de-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6e0de-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="6e0de-151">timeZone: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6e0de-151">timeZone: String</span></span>

<span data-ttu-id="6e0de-152">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="6e0de-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="6e0de-153">Tipo</span><span class="sxs-lookup"><span data-stu-id="6e0de-153">Type</span></span>

*   <span data-ttu-id="6e0de-154">String</span><span class="sxs-lookup"><span data-stu-id="6e0de-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6e0de-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6e0de-155">Requirements</span></span>

|<span data-ttu-id="6e0de-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="6e0de-156">Requirement</span></span>| <span data-ttu-id="6e0de-157">Valor</span><span class="sxs-lookup"><span data-stu-id="6e0de-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="6e0de-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6e0de-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6e0de-159">1.0</span><span class="sxs-lookup"><span data-stu-id="6e0de-159">1.0</span></span>|
|[<span data-ttu-id="6e0de-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6e0de-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6e0de-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6e0de-161">ReadItem</span></span>|
|[<span data-ttu-id="6e0de-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6e0de-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6e0de-163">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6e0de-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6e0de-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6e0de-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
