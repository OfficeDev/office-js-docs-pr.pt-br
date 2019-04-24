---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 7a10a35887d31a8803d0662eedbe190543d2326a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451917"
---
# <a name="userprofile"></a><span data-ttu-id="2b224-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="2b224-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="2b224-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="2b224-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b224-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2b224-104">Requirements</span></span>

|<span data-ttu-id="2b224-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="2b224-105">Requirement</span></span>| <span data-ttu-id="2b224-106">Valor</span><span class="sxs-lookup"><span data-stu-id="2b224-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b224-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2b224-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b224-108">1.0</span><span class="sxs-lookup"><span data-stu-id="2b224-108">1.0</span></span>|
|[<span data-ttu-id="2b224-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2b224-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2b224-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2b224-110">ReadItem</span></span>|
|[<span data-ttu-id="2b224-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2b224-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2b224-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2b224-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="2b224-113">Membros</span><span class="sxs-lookup"><span data-stu-id="2b224-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="2b224-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="2b224-114">displayName :String</span></span>

<span data-ttu-id="2b224-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="2b224-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="2b224-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="2b224-116">Type</span></span>

*   <span data-ttu-id="2b224-117">String</span><span class="sxs-lookup"><span data-stu-id="2b224-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b224-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2b224-118">Requirements</span></span>

|<span data-ttu-id="2b224-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="2b224-119">Requirement</span></span>| <span data-ttu-id="2b224-120">Valor</span><span class="sxs-lookup"><span data-stu-id="2b224-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b224-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2b224-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b224-122">1.0</span><span class="sxs-lookup"><span data-stu-id="2b224-122">1.0</span></span>|
|[<span data-ttu-id="2b224-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2b224-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2b224-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2b224-124">ReadItem</span></span>|
|[<span data-ttu-id="2b224-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2b224-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2b224-126">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2b224-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b224-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2b224-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="2b224-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="2b224-128">emailAddress :String</span></span>

<span data-ttu-id="2b224-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="2b224-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="2b224-130">Tipo</span><span class="sxs-lookup"><span data-stu-id="2b224-130">Type</span></span>

*   <span data-ttu-id="2b224-131">String</span><span class="sxs-lookup"><span data-stu-id="2b224-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b224-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2b224-132">Requirements</span></span>

|<span data-ttu-id="2b224-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="2b224-133">Requirement</span></span>| <span data-ttu-id="2b224-134">Valor</span><span class="sxs-lookup"><span data-stu-id="2b224-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b224-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2b224-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b224-136">1.0</span><span class="sxs-lookup"><span data-stu-id="2b224-136">1.0</span></span>|
|[<span data-ttu-id="2b224-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2b224-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2b224-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2b224-138">ReadItem</span></span>|
|[<span data-ttu-id="2b224-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2b224-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2b224-140">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2b224-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b224-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2b224-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="2b224-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="2b224-142">timeZone :String</span></span>

<span data-ttu-id="2b224-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="2b224-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="2b224-144">Tipo</span><span class="sxs-lookup"><span data-stu-id="2b224-144">Type</span></span>

*   <span data-ttu-id="2b224-145">String</span><span class="sxs-lookup"><span data-stu-id="2b224-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b224-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2b224-146">Requirements</span></span>

|<span data-ttu-id="2b224-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="2b224-147">Requirement</span></span>| <span data-ttu-id="2b224-148">Valor</span><span class="sxs-lookup"><span data-stu-id="2b224-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b224-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2b224-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b224-150">1.0</span><span class="sxs-lookup"><span data-stu-id="2b224-150">1.0</span></span>|
|[<span data-ttu-id="2b224-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="2b224-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2b224-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2b224-152">ReadItem</span></span>|
|[<span data-ttu-id="2b224-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2b224-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2b224-154">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2b224-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b224-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="2b224-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
