---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: fc20497cc8df8d091ba0195f7dca9b283ff4d1c2
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871014"
---
# <a name="userprofile"></a><span data-ttu-id="ab9c7-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ab9c7-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ab9c7-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ab9c7-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab9c7-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab9c7-104">Requirements</span></span>

|<span data-ttu-id="ab9c7-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab9c7-105">Requirement</span></span>| <span data-ttu-id="ab9c7-106">Valor</span><span class="sxs-lookup"><span data-stu-id="ab9c7-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab9c7-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab9c7-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab9c7-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ab9c7-108">1.0</span></span>|
|[<span data-ttu-id="ab9c7-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab9c7-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab9c7-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab9c7-110">ReadItem</span></span>|
|[<span data-ttu-id="ab9c7-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab9c7-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ab9c7-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ab9c7-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ab9c7-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="ab9c7-113">Members and methods</span></span>

| <span data-ttu-id="ab9c7-114">Membro</span><span class="sxs-lookup"><span data-stu-id="ab9c7-114">Member</span></span> | <span data-ttu-id="ab9c7-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab9c7-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ab9c7-116">displayName</span><span class="sxs-lookup"><span data-stu-id="ab9c7-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="ab9c7-117">Member</span><span class="sxs-lookup"><span data-stu-id="ab9c7-117">Member</span></span> |
| [<span data-ttu-id="ab9c7-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="ab9c7-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="ab9c7-119">Member</span><span class="sxs-lookup"><span data-stu-id="ab9c7-119">Member</span></span> |
| [<span data-ttu-id="ab9c7-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="ab9c7-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="ab9c7-121">Membro</span><span class="sxs-lookup"><span data-stu-id="ab9c7-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="ab9c7-122">Membros</span><span class="sxs-lookup"><span data-stu-id="ab9c7-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="ab9c7-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ab9c7-123">displayName :String</span></span>

<span data-ttu-id="ab9c7-124">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="ab9c7-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ab9c7-125">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab9c7-125">Type</span></span>

*   <span data-ttu-id="ab9c7-126">String</span><span class="sxs-lookup"><span data-stu-id="ab9c7-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab9c7-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab9c7-127">Requirements</span></span>

|<span data-ttu-id="ab9c7-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab9c7-128">Requirement</span></span>| <span data-ttu-id="ab9c7-129">Valor</span><span class="sxs-lookup"><span data-stu-id="ab9c7-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab9c7-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab9c7-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab9c7-131">1.0</span><span class="sxs-lookup"><span data-stu-id="ab9c7-131">1.0</span></span>|
|[<span data-ttu-id="ab9c7-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab9c7-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab9c7-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab9c7-133">ReadItem</span></span>|
|[<span data-ttu-id="ab9c7-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab9c7-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ab9c7-135">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ab9c7-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab9c7-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab9c7-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ab9c7-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ab9c7-137">emailAddress :String</span></span>

<span data-ttu-id="ab9c7-138">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="ab9c7-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ab9c7-139">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab9c7-139">Type</span></span>

*   <span data-ttu-id="ab9c7-140">String</span><span class="sxs-lookup"><span data-stu-id="ab9c7-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab9c7-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab9c7-141">Requirements</span></span>

|<span data-ttu-id="ab9c7-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab9c7-142">Requirement</span></span>| <span data-ttu-id="ab9c7-143">Valor</span><span class="sxs-lookup"><span data-stu-id="ab9c7-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab9c7-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab9c7-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab9c7-145">1.0</span><span class="sxs-lookup"><span data-stu-id="ab9c7-145">1.0</span></span>|
|[<span data-ttu-id="ab9c7-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab9c7-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab9c7-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab9c7-147">ReadItem</span></span>|
|[<span data-ttu-id="ab9c7-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab9c7-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ab9c7-149">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ab9c7-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab9c7-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab9c7-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ab9c7-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ab9c7-151">timeZone :String</span></span>

<span data-ttu-id="ab9c7-152">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="ab9c7-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ab9c7-153">Tipo</span><span class="sxs-lookup"><span data-stu-id="ab9c7-153">Type</span></span>

*   <span data-ttu-id="ab9c7-154">String</span><span class="sxs-lookup"><span data-stu-id="ab9c7-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ab9c7-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ab9c7-155">Requirements</span></span>

|<span data-ttu-id="ab9c7-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="ab9c7-156">Requirement</span></span>| <span data-ttu-id="ab9c7-157">Valor</span><span class="sxs-lookup"><span data-stu-id="ab9c7-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="ab9c7-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ab9c7-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ab9c7-159">1.0</span><span class="sxs-lookup"><span data-stu-id="ab9c7-159">1.0</span></span>|
|[<span data-ttu-id="ab9c7-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ab9c7-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ab9c7-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ab9c7-161">ReadItem</span></span>|
|[<span data-ttu-id="ab9c7-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ab9c7-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ab9c7-163">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ab9c7-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ab9c7-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ab9c7-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
