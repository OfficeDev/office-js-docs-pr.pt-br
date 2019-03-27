---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 7a10a35887d31a8803d0662eedbe190543d2326a
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870188"
---
# <a name="userprofile"></a><span data-ttu-id="957fb-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="957fb-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="957fb-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="957fb-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="957fb-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="957fb-104">Requirements</span></span>

|<span data-ttu-id="957fb-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="957fb-105">Requirement</span></span>| <span data-ttu-id="957fb-106">Valor</span><span class="sxs-lookup"><span data-stu-id="957fb-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="957fb-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="957fb-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="957fb-108">1.0</span><span class="sxs-lookup"><span data-stu-id="957fb-108">1.0</span></span>|
|[<span data-ttu-id="957fb-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="957fb-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="957fb-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="957fb-110">ReadItem</span></span>|
|[<span data-ttu-id="957fb-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="957fb-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="957fb-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="957fb-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="957fb-113">Membros</span><span class="sxs-lookup"><span data-stu-id="957fb-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="957fb-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="957fb-114">displayName :String</span></span>

<span data-ttu-id="957fb-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="957fb-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="957fb-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="957fb-116">Type</span></span>

*   <span data-ttu-id="957fb-117">String</span><span class="sxs-lookup"><span data-stu-id="957fb-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="957fb-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="957fb-118">Requirements</span></span>

|<span data-ttu-id="957fb-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="957fb-119">Requirement</span></span>| <span data-ttu-id="957fb-120">Valor</span><span class="sxs-lookup"><span data-stu-id="957fb-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="957fb-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="957fb-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="957fb-122">1.0</span><span class="sxs-lookup"><span data-stu-id="957fb-122">1.0</span></span>|
|[<span data-ttu-id="957fb-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="957fb-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="957fb-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="957fb-124">ReadItem</span></span>|
|[<span data-ttu-id="957fb-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="957fb-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="957fb-126">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="957fb-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="957fb-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="957fb-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="957fb-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="957fb-128">emailAddress :String</span></span>

<span data-ttu-id="957fb-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="957fb-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="957fb-130">Tipo</span><span class="sxs-lookup"><span data-stu-id="957fb-130">Type</span></span>

*   <span data-ttu-id="957fb-131">String</span><span class="sxs-lookup"><span data-stu-id="957fb-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="957fb-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="957fb-132">Requirements</span></span>

|<span data-ttu-id="957fb-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="957fb-133">Requirement</span></span>| <span data-ttu-id="957fb-134">Valor</span><span class="sxs-lookup"><span data-stu-id="957fb-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="957fb-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="957fb-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="957fb-136">1.0</span><span class="sxs-lookup"><span data-stu-id="957fb-136">1.0</span></span>|
|[<span data-ttu-id="957fb-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="957fb-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="957fb-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="957fb-138">ReadItem</span></span>|
|[<span data-ttu-id="957fb-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="957fb-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="957fb-140">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="957fb-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="957fb-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="957fb-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="957fb-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="957fb-142">timeZone :String</span></span>

<span data-ttu-id="957fb-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="957fb-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="957fb-144">Tipo</span><span class="sxs-lookup"><span data-stu-id="957fb-144">Type</span></span>

*   <span data-ttu-id="957fb-145">String</span><span class="sxs-lookup"><span data-stu-id="957fb-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="957fb-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="957fb-146">Requirements</span></span>

|<span data-ttu-id="957fb-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="957fb-147">Requirement</span></span>| <span data-ttu-id="957fb-148">Valor</span><span class="sxs-lookup"><span data-stu-id="957fb-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="957fb-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="957fb-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="957fb-150">1.0</span><span class="sxs-lookup"><span data-stu-id="957fb-150">1.0</span></span>|
|[<span data-ttu-id="957fb-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="957fb-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="957fb-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="957fb-152">ReadItem</span></span>|
|[<span data-ttu-id="957fb-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="957fb-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="957fb-154">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="957fb-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="957fb-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="957fb-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
