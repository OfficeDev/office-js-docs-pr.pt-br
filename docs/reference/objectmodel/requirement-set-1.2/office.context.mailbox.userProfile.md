---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 4a6739c9b463e49d41e320094a4c9cb1a32655f4
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067822"
---
# <a name="userprofile"></a><span data-ttu-id="816e3-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="816e3-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="816e3-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="816e3-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="816e3-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="816e3-104">Requirements</span></span>

|<span data-ttu-id="816e3-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="816e3-105">Requirement</span></span>| <span data-ttu-id="816e3-106">Valor</span><span class="sxs-lookup"><span data-stu-id="816e3-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="816e3-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="816e3-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="816e3-108">1.0</span><span class="sxs-lookup"><span data-stu-id="816e3-108">1.0</span></span>|
|[<span data-ttu-id="816e3-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="816e3-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="816e3-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="816e3-110">ReadItem</span></span>|
|[<span data-ttu-id="816e3-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="816e3-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="816e3-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="816e3-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="816e3-113">Membros</span><span class="sxs-lookup"><span data-stu-id="816e3-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="816e3-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="816e3-114">displayName :String</span></span>

<span data-ttu-id="816e3-115">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="816e3-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="816e3-116">Tipo</span><span class="sxs-lookup"><span data-stu-id="816e3-116">Type</span></span>

*   <span data-ttu-id="816e3-117">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="816e3-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="816e3-118">Requisitos</span><span class="sxs-lookup"><span data-stu-id="816e3-118">Requirements</span></span>

|<span data-ttu-id="816e3-119">Requisito</span><span class="sxs-lookup"><span data-stu-id="816e3-119">Requirement</span></span>| <span data-ttu-id="816e3-120">Valor</span><span class="sxs-lookup"><span data-stu-id="816e3-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="816e3-121">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="816e3-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="816e3-122">1.0</span><span class="sxs-lookup"><span data-stu-id="816e3-122">1.0</span></span>|
|[<span data-ttu-id="816e3-123">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="816e3-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="816e3-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="816e3-124">ReadItem</span></span>|
|[<span data-ttu-id="816e3-125">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="816e3-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="816e3-126">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="816e3-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="816e3-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="816e3-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="816e3-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="816e3-128">emailAddress :String</span></span>

<span data-ttu-id="816e3-129">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="816e3-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="816e3-130">Tipo</span><span class="sxs-lookup"><span data-stu-id="816e3-130">Type</span></span>

*   <span data-ttu-id="816e3-131">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="816e3-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="816e3-132">Requisitos</span><span class="sxs-lookup"><span data-stu-id="816e3-132">Requirements</span></span>

|<span data-ttu-id="816e3-133">Requisito</span><span class="sxs-lookup"><span data-stu-id="816e3-133">Requirement</span></span>| <span data-ttu-id="816e3-134">Valor</span><span class="sxs-lookup"><span data-stu-id="816e3-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="816e3-135">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="816e3-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="816e3-136">1.0</span><span class="sxs-lookup"><span data-stu-id="816e3-136">1.0</span></span>|
|[<span data-ttu-id="816e3-137">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="816e3-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="816e3-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="816e3-138">ReadItem</span></span>|
|[<span data-ttu-id="816e3-139">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="816e3-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="816e3-140">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="816e3-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="816e3-141">Exemplo</span><span class="sxs-lookup"><span data-stu-id="816e3-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="816e3-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="816e3-142">timeZone :String</span></span>

<span data-ttu-id="816e3-143">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="816e3-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="816e3-144">Tipo</span><span class="sxs-lookup"><span data-stu-id="816e3-144">Type</span></span>

*   <span data-ttu-id="816e3-145">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="816e3-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="816e3-146">Requisitos</span><span class="sxs-lookup"><span data-stu-id="816e3-146">Requirements</span></span>

|<span data-ttu-id="816e3-147">Requisito</span><span class="sxs-lookup"><span data-stu-id="816e3-147">Requirement</span></span>| <span data-ttu-id="816e3-148">Valor</span><span class="sxs-lookup"><span data-stu-id="816e3-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="816e3-149">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="816e3-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="816e3-150">1.0</span><span class="sxs-lookup"><span data-stu-id="816e3-150">1.0</span></span>|
|[<span data-ttu-id="816e3-151">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="816e3-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="816e3-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="816e3-152">ReadItem</span></span>|
|[<span data-ttu-id="816e3-153">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="816e3-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="816e3-154">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="816e3-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="816e3-155">Exemplo</span><span class="sxs-lookup"><span data-stu-id="816e3-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
