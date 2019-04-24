---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: fc20497cc8df8d091ba0195f7dca9b283ff4d1c2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451819"
---
# <a name="userprofile"></a><span data-ttu-id="52f7d-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="52f7d-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="52f7d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="52f7d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="52f7d-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="52f7d-104">Requirements</span></span>

|<span data-ttu-id="52f7d-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="52f7d-105">Requirement</span></span>| <span data-ttu-id="52f7d-106">Valor</span><span class="sxs-lookup"><span data-stu-id="52f7d-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="52f7d-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="52f7d-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="52f7d-108">1.0</span><span class="sxs-lookup"><span data-stu-id="52f7d-108">1.0</span></span>|
|[<span data-ttu-id="52f7d-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="52f7d-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="52f7d-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="52f7d-110">ReadItem</span></span>|
|[<span data-ttu-id="52f7d-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="52f7d-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="52f7d-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="52f7d-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="52f7d-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="52f7d-113">Members and methods</span></span>

| <span data-ttu-id="52f7d-114">Membro</span><span class="sxs-lookup"><span data-stu-id="52f7d-114">Member</span></span> | <span data-ttu-id="52f7d-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="52f7d-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="52f7d-116">displayName</span><span class="sxs-lookup"><span data-stu-id="52f7d-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="52f7d-117">Member</span><span class="sxs-lookup"><span data-stu-id="52f7d-117">Member</span></span> |
| [<span data-ttu-id="52f7d-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="52f7d-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="52f7d-119">Member</span><span class="sxs-lookup"><span data-stu-id="52f7d-119">Member</span></span> |
| [<span data-ttu-id="52f7d-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="52f7d-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="52f7d-121">Membro</span><span class="sxs-lookup"><span data-stu-id="52f7d-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="52f7d-122">Membros</span><span class="sxs-lookup"><span data-stu-id="52f7d-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="52f7d-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="52f7d-123">displayName :String</span></span>

<span data-ttu-id="52f7d-124">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="52f7d-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="52f7d-125">Tipo</span><span class="sxs-lookup"><span data-stu-id="52f7d-125">Type</span></span>

*   <span data-ttu-id="52f7d-126">String</span><span class="sxs-lookup"><span data-stu-id="52f7d-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="52f7d-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="52f7d-127">Requirements</span></span>

|<span data-ttu-id="52f7d-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="52f7d-128">Requirement</span></span>| <span data-ttu-id="52f7d-129">Valor</span><span class="sxs-lookup"><span data-stu-id="52f7d-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="52f7d-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="52f7d-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="52f7d-131">1.0</span><span class="sxs-lookup"><span data-stu-id="52f7d-131">1.0</span></span>|
|[<span data-ttu-id="52f7d-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="52f7d-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="52f7d-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="52f7d-133">ReadItem</span></span>|
|[<span data-ttu-id="52f7d-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="52f7d-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="52f7d-135">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="52f7d-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="52f7d-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="52f7d-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="52f7d-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="52f7d-137">emailAddress :String</span></span>

<span data-ttu-id="52f7d-138">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="52f7d-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="52f7d-139">Tipo</span><span class="sxs-lookup"><span data-stu-id="52f7d-139">Type</span></span>

*   <span data-ttu-id="52f7d-140">String</span><span class="sxs-lookup"><span data-stu-id="52f7d-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="52f7d-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="52f7d-141">Requirements</span></span>

|<span data-ttu-id="52f7d-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="52f7d-142">Requirement</span></span>| <span data-ttu-id="52f7d-143">Valor</span><span class="sxs-lookup"><span data-stu-id="52f7d-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="52f7d-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="52f7d-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="52f7d-145">1.0</span><span class="sxs-lookup"><span data-stu-id="52f7d-145">1.0</span></span>|
|[<span data-ttu-id="52f7d-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="52f7d-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="52f7d-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="52f7d-147">ReadItem</span></span>|
|[<span data-ttu-id="52f7d-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="52f7d-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="52f7d-149">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="52f7d-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="52f7d-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="52f7d-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="52f7d-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="52f7d-151">timeZone :String</span></span>

<span data-ttu-id="52f7d-152">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="52f7d-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="52f7d-153">Tipo</span><span class="sxs-lookup"><span data-stu-id="52f7d-153">Type</span></span>

*   <span data-ttu-id="52f7d-154">String</span><span class="sxs-lookup"><span data-stu-id="52f7d-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="52f7d-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="52f7d-155">Requirements</span></span>

|<span data-ttu-id="52f7d-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="52f7d-156">Requirement</span></span>| <span data-ttu-id="52f7d-157">Valor</span><span class="sxs-lookup"><span data-stu-id="52f7d-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="52f7d-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="52f7d-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="52f7d-159">1.0</span><span class="sxs-lookup"><span data-stu-id="52f7d-159">1.0</span></span>|
|[<span data-ttu-id="52f7d-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="52f7d-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="52f7d-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="52f7d-161">ReadItem</span></span>|
|[<span data-ttu-id="52f7d-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="52f7d-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="52f7d-163">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="52f7d-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="52f7d-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="52f7d-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
