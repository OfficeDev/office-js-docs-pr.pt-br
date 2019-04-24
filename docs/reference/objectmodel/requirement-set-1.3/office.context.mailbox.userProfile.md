---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 03cdc13845bff0fbd3855f29f43298cd770e5ad9
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451840"
---
# <a name="userprofile"></a><span data-ttu-id="741de-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="741de-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="741de-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="741de-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="741de-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="741de-104">Requirements</span></span>

|<span data-ttu-id="741de-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="741de-105">Requirement</span></span>| <span data-ttu-id="741de-106">Valor</span><span class="sxs-lookup"><span data-stu-id="741de-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="741de-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="741de-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="741de-108">1.0</span><span class="sxs-lookup"><span data-stu-id="741de-108">1.0</span></span>|
|[<span data-ttu-id="741de-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="741de-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="741de-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="741de-110">ReadItem</span></span>|
|[<span data-ttu-id="741de-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="741de-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="741de-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="741de-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="741de-113">Membros</span><span class="sxs-lookup"><span data-stu-id="741de-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="741de-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="741de-114">displayName :String</span></span>

<span data-ttu-id="741de-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="741de-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="741de-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="741de-116">Type</span></span>

*   <span data-ttu-id="741de-117">String</span><span class="sxs-lookup"><span data-stu-id="741de-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="741de-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="741de-118">Requirements</span></span>

|<span data-ttu-id="741de-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="741de-119">Requirement</span></span>| <span data-ttu-id="741de-120">Valor</span><span class="sxs-lookup"><span data-stu-id="741de-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="741de-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="741de-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="741de-122">1.0</span><span class="sxs-lookup"><span data-stu-id="741de-122">1.0</span></span>|
|[<span data-ttu-id="741de-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="741de-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="741de-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="741de-124">ReadItem</span></span>|
|[<span data-ttu-id="741de-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="741de-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="741de-126">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="741de-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="741de-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="741de-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="741de-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="741de-128">emailAddress :String</span></span>

<span data-ttu-id="741de-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="741de-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="741de-130">Tipo</span><span class="sxs-lookup"><span data-stu-id="741de-130">Type</span></span>

*   <span data-ttu-id="741de-131">String</span><span class="sxs-lookup"><span data-stu-id="741de-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="741de-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="741de-132">Requirements</span></span>

|<span data-ttu-id="741de-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="741de-133">Requirement</span></span>| <span data-ttu-id="741de-134">Valor</span><span class="sxs-lookup"><span data-stu-id="741de-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="741de-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="741de-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="741de-136">1.0</span><span class="sxs-lookup"><span data-stu-id="741de-136">1.0</span></span>|
|[<span data-ttu-id="741de-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="741de-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="741de-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="741de-138">ReadItem</span></span>|
|[<span data-ttu-id="741de-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="741de-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="741de-140">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="741de-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="741de-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="741de-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="741de-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="741de-142">timeZone :String</span></span>

<span data-ttu-id="741de-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="741de-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="741de-144">Tipo</span><span class="sxs-lookup"><span data-stu-id="741de-144">Type</span></span>

*   <span data-ttu-id="741de-145">String</span><span class="sxs-lookup"><span data-stu-id="741de-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="741de-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="741de-146">Requirements</span></span>

|<span data-ttu-id="741de-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="741de-147">Requirement</span></span>| <span data-ttu-id="741de-148">Valor</span><span class="sxs-lookup"><span data-stu-id="741de-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="741de-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="741de-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="741de-150">1.0</span><span class="sxs-lookup"><span data-stu-id="741de-150">1.0</span></span>|
|[<span data-ttu-id="741de-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="741de-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="741de-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="741de-152">ReadItem</span></span>|
|[<span data-ttu-id="741de-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="741de-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="741de-154">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="741de-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="741de-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="741de-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
