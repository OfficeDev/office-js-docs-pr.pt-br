---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2798b07b3353e9d89f757a22e6bed19dbd94a1c5
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870041"
---
# <a name="userprofile"></a><span data-ttu-id="daa74-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="daa74-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="daa74-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="daa74-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa74-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="daa74-104">Requirements</span></span>

|<span data-ttu-id="daa74-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="daa74-105">Requirement</span></span>| <span data-ttu-id="daa74-106">Valor</span><span class="sxs-lookup"><span data-stu-id="daa74-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa74-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="daa74-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa74-108">1.0</span><span class="sxs-lookup"><span data-stu-id="daa74-108">1.0</span></span>|
|[<span data-ttu-id="daa74-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="daa74-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa74-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa74-110">ReadItem</span></span>|
|[<span data-ttu-id="daa74-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="daa74-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa74-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="daa74-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="daa74-113">Membros</span><span class="sxs-lookup"><span data-stu-id="daa74-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="daa74-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="daa74-114">displayName :String</span></span>

<span data-ttu-id="daa74-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="daa74-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="daa74-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="daa74-116">Type</span></span>

*   <span data-ttu-id="daa74-117">String</span><span class="sxs-lookup"><span data-stu-id="daa74-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa74-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="daa74-118">Requirements</span></span>

|<span data-ttu-id="daa74-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="daa74-119">Requirement</span></span>| <span data-ttu-id="daa74-120">Valor</span><span class="sxs-lookup"><span data-stu-id="daa74-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa74-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="daa74-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa74-122">1.0</span><span class="sxs-lookup"><span data-stu-id="daa74-122">1.0</span></span>|
|[<span data-ttu-id="daa74-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="daa74-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa74-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa74-124">ReadItem</span></span>|
|[<span data-ttu-id="daa74-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="daa74-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa74-126">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="daa74-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa74-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="daa74-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="daa74-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="daa74-128">emailAddress :String</span></span>

<span data-ttu-id="daa74-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="daa74-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="daa74-130">Tipo</span><span class="sxs-lookup"><span data-stu-id="daa74-130">Type</span></span>

*   <span data-ttu-id="daa74-131">String</span><span class="sxs-lookup"><span data-stu-id="daa74-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa74-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="daa74-132">Requirements</span></span>

|<span data-ttu-id="daa74-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="daa74-133">Requirement</span></span>| <span data-ttu-id="daa74-134">Valor</span><span class="sxs-lookup"><span data-stu-id="daa74-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa74-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="daa74-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa74-136">1.0</span><span class="sxs-lookup"><span data-stu-id="daa74-136">1.0</span></span>|
|[<span data-ttu-id="daa74-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="daa74-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa74-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa74-138">ReadItem</span></span>|
|[<span data-ttu-id="daa74-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="daa74-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa74-140">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="daa74-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa74-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="daa74-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="daa74-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="daa74-142">timeZone :String</span></span>

<span data-ttu-id="daa74-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="daa74-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="daa74-144">Tipo</span><span class="sxs-lookup"><span data-stu-id="daa74-144">Type</span></span>

*   <span data-ttu-id="daa74-145">String</span><span class="sxs-lookup"><span data-stu-id="daa74-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa74-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="daa74-146">Requirements</span></span>

|<span data-ttu-id="daa74-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="daa74-147">Requirement</span></span>| <span data-ttu-id="daa74-148">Valor</span><span class="sxs-lookup"><span data-stu-id="daa74-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa74-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="daa74-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa74-150">1.0</span><span class="sxs-lookup"><span data-stu-id="daa74-150">1.0</span></span>|
|[<span data-ttu-id="daa74-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="daa74-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa74-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa74-152">ReadItem</span></span>|
|[<span data-ttu-id="daa74-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="daa74-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa74-154">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="daa74-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa74-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="daa74-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
