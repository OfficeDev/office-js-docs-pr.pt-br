---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 7258195e7ec0ef2432723d0f32f3d9ef1a3acf2b
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268681"
---
# <a name="userprofile"></a><span data-ttu-id="6c37c-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="6c37c-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="6c37c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="6c37c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="6c37c-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6c37c-104">Requirements</span></span>

|<span data-ttu-id="6c37c-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="6c37c-105">Requirement</span></span>| <span data-ttu-id="6c37c-106">Valor</span><span class="sxs-lookup"><span data-stu-id="6c37c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c37c-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6c37c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c37c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="6c37c-108">1.0</span></span>|
|[<span data-ttu-id="6c37c-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6c37c-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c37c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6c37c-110">ReadItem</span></span>|
|[<span data-ttu-id="6c37c-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6c37c-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6c37c-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6c37c-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6c37c-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="6c37c-113">Members and methods</span></span>

| <span data-ttu-id="6c37c-114">Membro</span><span class="sxs-lookup"><span data-stu-id="6c37c-114">Member</span></span> | <span data-ttu-id="6c37c-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="6c37c-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6c37c-116">displayName</span><span class="sxs-lookup"><span data-stu-id="6c37c-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="6c37c-117">Membro</span><span class="sxs-lookup"><span data-stu-id="6c37c-117">Member</span></span> |
| [<span data-ttu-id="6c37c-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="6c37c-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="6c37c-119">Membro</span><span class="sxs-lookup"><span data-stu-id="6c37c-119">Member</span></span> |
| [<span data-ttu-id="6c37c-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="6c37c-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="6c37c-121">Membro</span><span class="sxs-lookup"><span data-stu-id="6c37c-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="6c37c-122">Membros</span><span class="sxs-lookup"><span data-stu-id="6c37c-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="6c37c-123">displayName: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6c37c-123">displayName: String</span></span>

<span data-ttu-id="6c37c-124">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="6c37c-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="6c37c-125">Tipo</span><span class="sxs-lookup"><span data-stu-id="6c37c-125">Type</span></span>

*   <span data-ttu-id="6c37c-126">String</span><span class="sxs-lookup"><span data-stu-id="6c37c-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6c37c-127">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6c37c-127">Requirements</span></span>

|<span data-ttu-id="6c37c-128">Requisito</span><span class="sxs-lookup"><span data-stu-id="6c37c-128">Requirement</span></span>| <span data-ttu-id="6c37c-129">Valor</span><span class="sxs-lookup"><span data-stu-id="6c37c-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c37c-130">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6c37c-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c37c-131">1.0</span><span class="sxs-lookup"><span data-stu-id="6c37c-131">1.0</span></span>|
|[<span data-ttu-id="6c37c-132">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6c37c-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c37c-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6c37c-133">ReadItem</span></span>|
|[<span data-ttu-id="6c37c-134">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6c37c-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6c37c-135">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6c37c-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c37c-136">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6c37c-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="6c37c-137">emailAddress: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6c37c-137">emailAddress: String</span></span>

<span data-ttu-id="6c37c-138">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="6c37c-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="6c37c-139">Tipo</span><span class="sxs-lookup"><span data-stu-id="6c37c-139">Type</span></span>

*   <span data-ttu-id="6c37c-140">String</span><span class="sxs-lookup"><span data-stu-id="6c37c-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6c37c-141">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6c37c-141">Requirements</span></span>

|<span data-ttu-id="6c37c-142">Requisito</span><span class="sxs-lookup"><span data-stu-id="6c37c-142">Requirement</span></span>| <span data-ttu-id="6c37c-143">Valor</span><span class="sxs-lookup"><span data-stu-id="6c37c-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c37c-144">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6c37c-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c37c-145">1.0</span><span class="sxs-lookup"><span data-stu-id="6c37c-145">1.0</span></span>|
|[<span data-ttu-id="6c37c-146">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6c37c-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c37c-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6c37c-147">ReadItem</span></span>|
|[<span data-ttu-id="6c37c-148">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6c37c-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6c37c-149">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6c37c-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c37c-150">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6c37c-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="6c37c-151">timeZone: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6c37c-151">timeZone: String</span></span>

<span data-ttu-id="6c37c-152">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="6c37c-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="6c37c-153">Tipo</span><span class="sxs-lookup"><span data-stu-id="6c37c-153">Type</span></span>

*   <span data-ttu-id="6c37c-154">String</span><span class="sxs-lookup"><span data-stu-id="6c37c-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6c37c-155">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6c37c-155">Requirements</span></span>

|<span data-ttu-id="6c37c-156">Requisito</span><span class="sxs-lookup"><span data-stu-id="6c37c-156">Requirement</span></span>| <span data-ttu-id="6c37c-157">Valor</span><span class="sxs-lookup"><span data-stu-id="6c37c-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c37c-158">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6c37c-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6c37c-159">1.0</span><span class="sxs-lookup"><span data-stu-id="6c37c-159">1.0</span></span>|
|[<span data-ttu-id="6c37c-160">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="6c37c-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6c37c-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6c37c-161">ReadItem</span></span>|
|[<span data-ttu-id="6c37c-162">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6c37c-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6c37c-163">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6c37c-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c37c-164">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6c37c-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
