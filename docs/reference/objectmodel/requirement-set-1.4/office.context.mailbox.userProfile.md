---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2798b07b3353e9d89f757a22e6bed19dbd94a1c5
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450314"
---
# <a name="userprofile"></a><span data-ttu-id="5a7b2-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="5a7b2-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="5a7b2-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="5a7b2-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="5a7b2-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5a7b2-104">Requirements</span></span>

|<span data-ttu-id="5a7b2-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="5a7b2-105">Requirement</span></span>| <span data-ttu-id="5a7b2-106">Valor</span><span class="sxs-lookup"><span data-stu-id="5a7b2-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="5a7b2-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5a7b2-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5a7b2-108">1.0</span><span class="sxs-lookup"><span data-stu-id="5a7b2-108">1.0</span></span>|
|[<span data-ttu-id="5a7b2-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="5a7b2-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5a7b2-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5a7b2-110">ReadItem</span></span>|
|[<span data-ttu-id="5a7b2-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5a7b2-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5a7b2-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5a7b2-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="5a7b2-113">Membros</span><span class="sxs-lookup"><span data-stu-id="5a7b2-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="5a7b2-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="5a7b2-114">displayName :String</span></span>

<span data-ttu-id="5a7b2-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="5a7b2-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="5a7b2-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="5a7b2-116">Type</span></span>

*   <span data-ttu-id="5a7b2-117">String</span><span class="sxs-lookup"><span data-stu-id="5a7b2-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5a7b2-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5a7b2-118">Requirements</span></span>

|<span data-ttu-id="5a7b2-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="5a7b2-119">Requirement</span></span>| <span data-ttu-id="5a7b2-120">Valor</span><span class="sxs-lookup"><span data-stu-id="5a7b2-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="5a7b2-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5a7b2-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5a7b2-122">1.0</span><span class="sxs-lookup"><span data-stu-id="5a7b2-122">1.0</span></span>|
|[<span data-ttu-id="5a7b2-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="5a7b2-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5a7b2-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5a7b2-124">ReadItem</span></span>|
|[<span data-ttu-id="5a7b2-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5a7b2-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5a7b2-126">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5a7b2-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5a7b2-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="5a7b2-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="5a7b2-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="5a7b2-128">emailAddress :String</span></span>

<span data-ttu-id="5a7b2-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="5a7b2-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="5a7b2-130">Tipo</span><span class="sxs-lookup"><span data-stu-id="5a7b2-130">Type</span></span>

*   <span data-ttu-id="5a7b2-131">String</span><span class="sxs-lookup"><span data-stu-id="5a7b2-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5a7b2-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5a7b2-132">Requirements</span></span>

|<span data-ttu-id="5a7b2-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="5a7b2-133">Requirement</span></span>| <span data-ttu-id="5a7b2-134">Valor</span><span class="sxs-lookup"><span data-stu-id="5a7b2-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="5a7b2-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5a7b2-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5a7b2-136">1.0</span><span class="sxs-lookup"><span data-stu-id="5a7b2-136">1.0</span></span>|
|[<span data-ttu-id="5a7b2-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="5a7b2-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5a7b2-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5a7b2-138">ReadItem</span></span>|
|[<span data-ttu-id="5a7b2-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5a7b2-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5a7b2-140">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5a7b2-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5a7b2-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="5a7b2-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="5a7b2-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="5a7b2-142">timeZone :String</span></span>

<span data-ttu-id="5a7b2-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="5a7b2-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="5a7b2-144">Tipo</span><span class="sxs-lookup"><span data-stu-id="5a7b2-144">Type</span></span>

*   <span data-ttu-id="5a7b2-145">String</span><span class="sxs-lookup"><span data-stu-id="5a7b2-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5a7b2-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5a7b2-146">Requirements</span></span>

|<span data-ttu-id="5a7b2-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="5a7b2-147">Requirement</span></span>| <span data-ttu-id="5a7b2-148">Valor</span><span class="sxs-lookup"><span data-stu-id="5a7b2-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="5a7b2-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5a7b2-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5a7b2-150">1.0</span><span class="sxs-lookup"><span data-stu-id="5a7b2-150">1.0</span></span>|
|[<span data-ttu-id="5a7b2-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="5a7b2-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5a7b2-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5a7b2-152">ReadItem</span></span>|
|[<span data-ttu-id="5a7b2-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5a7b2-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5a7b2-154">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5a7b2-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5a7b2-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="5a7b2-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
