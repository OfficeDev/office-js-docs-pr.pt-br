---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 496a59f4ef02f03cda95fde0bf14634b1db13f77
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450335"
---
# <a name="userprofile"></a><span data-ttu-id="c5b45-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="c5b45-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="c5b45-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="c5b45-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5b45-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5b45-104">Requirements</span></span>

|<span data-ttu-id="c5b45-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5b45-105">Requirement</span></span>| <span data-ttu-id="c5b45-106">Valor</span><span class="sxs-lookup"><span data-stu-id="c5b45-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5b45-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5b45-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5b45-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c5b45-108">1.0</span></span>|
|[<span data-ttu-id="c5b45-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c5b45-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c5b45-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c5b45-110">ReadItem</span></span>|
|[<span data-ttu-id="c5b45-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5b45-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5b45-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5b45-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="c5b45-113">Membros</span><span class="sxs-lookup"><span data-stu-id="c5b45-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="c5b45-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c5b45-114">displayName :String</span></span>

<span data-ttu-id="c5b45-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="c5b45-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c5b45-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="c5b45-116">Type</span></span>

*   <span data-ttu-id="c5b45-117">String</span><span class="sxs-lookup"><span data-stu-id="c5b45-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5b45-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5b45-118">Requirements</span></span>

|<span data-ttu-id="c5b45-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5b45-119">Requirement</span></span>| <span data-ttu-id="c5b45-120">Valor</span><span class="sxs-lookup"><span data-stu-id="c5b45-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5b45-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5b45-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5b45-122">1.0</span><span class="sxs-lookup"><span data-stu-id="c5b45-122">1.0</span></span>|
|[<span data-ttu-id="c5b45-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c5b45-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c5b45-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c5b45-124">ReadItem</span></span>|
|[<span data-ttu-id="c5b45-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5b45-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5b45-126">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5b45-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5b45-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c5b45-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c5b45-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c5b45-128">emailAddress :String</span></span>

<span data-ttu-id="c5b45-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="c5b45-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c5b45-130">Tipo</span><span class="sxs-lookup"><span data-stu-id="c5b45-130">Type</span></span>

*   <span data-ttu-id="c5b45-131">String</span><span class="sxs-lookup"><span data-stu-id="c5b45-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5b45-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5b45-132">Requirements</span></span>

|<span data-ttu-id="c5b45-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5b45-133">Requirement</span></span>| <span data-ttu-id="c5b45-134">Valor</span><span class="sxs-lookup"><span data-stu-id="c5b45-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5b45-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5b45-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5b45-136">1.0</span><span class="sxs-lookup"><span data-stu-id="c5b45-136">1.0</span></span>|
|[<span data-ttu-id="c5b45-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c5b45-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c5b45-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c5b45-138">ReadItem</span></span>|
|[<span data-ttu-id="c5b45-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5b45-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5b45-140">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5b45-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5b45-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c5b45-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c5b45-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c5b45-142">timeZone :String</span></span>

<span data-ttu-id="c5b45-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="c5b45-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c5b45-144">Tipo</span><span class="sxs-lookup"><span data-stu-id="c5b45-144">Type</span></span>

*   <span data-ttu-id="c5b45-145">String</span><span class="sxs-lookup"><span data-stu-id="c5b45-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5b45-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c5b45-146">Requirements</span></span>

|<span data-ttu-id="c5b45-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="c5b45-147">Requirement</span></span>| <span data-ttu-id="c5b45-148">Valor</span><span class="sxs-lookup"><span data-stu-id="c5b45-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5b45-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c5b45-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5b45-150">1.0</span><span class="sxs-lookup"><span data-stu-id="c5b45-150">1.0</span></span>|
|[<span data-ttu-id="c5b45-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="c5b45-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c5b45-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c5b45-152">ReadItem</span></span>|
|[<span data-ttu-id="c5b45-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c5b45-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5b45-154">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c5b45-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5b45-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="c5b45-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
