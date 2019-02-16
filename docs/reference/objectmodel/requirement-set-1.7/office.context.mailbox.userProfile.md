---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.7
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: fb55d11fd46a9957dab124514ef3bfe5a7c138eb
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067864"
---
# <a name="userprofile"></a><span data-ttu-id="be78d-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="be78d-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="be78d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="be78d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="be78d-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="be78d-104">Requirements</span></span>

|<span data-ttu-id="be78d-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="be78d-105">Requirement</span></span>| <span data-ttu-id="be78d-106">Valor</span><span class="sxs-lookup"><span data-stu-id="be78d-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="be78d-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="be78d-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be78d-108">1.0</span><span class="sxs-lookup"><span data-stu-id="be78d-108">1.0</span></span>|
|[<span data-ttu-id="be78d-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="be78d-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be78d-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be78d-110">ReadItem</span></span>|
|[<span data-ttu-id="be78d-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="be78d-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be78d-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="be78d-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="be78d-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="be78d-113">Members and methods</span></span>

| <span data-ttu-id="be78d-114">Membro</span><span class="sxs-lookup"><span data-stu-id="be78d-114">Member</span></span> | <span data-ttu-id="be78d-115">Type</span><span class="sxs-lookup"><span data-stu-id="be78d-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="be78d-116">accountType</span><span class="sxs-lookup"><span data-stu-id="be78d-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="be78d-117">Member</span><span class="sxs-lookup"><span data-stu-id="be78d-117">Member</span></span> |
| [<span data-ttu-id="be78d-118">displayName</span><span class="sxs-lookup"><span data-stu-id="be78d-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="be78d-119">Membro</span><span class="sxs-lookup"><span data-stu-id="be78d-119">Member</span></span> |
| [<span data-ttu-id="be78d-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="be78d-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="be78d-121">Membro</span><span class="sxs-lookup"><span data-stu-id="be78d-121">Member</span></span> |
| [<span data-ttu-id="be78d-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="be78d-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="be78d-123">Membro</span><span class="sxs-lookup"><span data-stu-id="be78d-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="be78d-124">Members</span><span class="sxs-lookup"><span data-stu-id="be78d-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="be78d-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="be78d-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="be78d-126">Atualmente, esse membro só tem suporte no Outlook 2016 para Mac (Build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="be78d-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="be78d-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="be78d-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="be78d-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="be78d-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="be78d-129">Value</span><span class="sxs-lookup"><span data-stu-id="be78d-129">Value</span></span> | <span data-ttu-id="be78d-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="be78d-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="be78d-131">A caixa de correio está em um servidor local do Exchange.</span><span class="sxs-lookup"><span data-stu-id="be78d-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="be78d-132">A caixa de correio está associada a uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="be78d-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="be78d-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="be78d-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="be78d-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="be78d-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="be78d-135">Tipo</span><span class="sxs-lookup"><span data-stu-id="be78d-135">Type</span></span>

*   <span data-ttu-id="be78d-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="be78d-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be78d-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="be78d-137">Requirements</span></span>

|<span data-ttu-id="be78d-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="be78d-138">Requirement</span></span>| <span data-ttu-id="be78d-139">Valor</span><span class="sxs-lookup"><span data-stu-id="be78d-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="be78d-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="be78d-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be78d-141">1.6</span><span class="sxs-lookup"><span data-stu-id="be78d-141">1.6</span></span> |
|[<span data-ttu-id="be78d-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="be78d-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be78d-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be78d-143">ReadItem</span></span>|
|[<span data-ttu-id="be78d-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="be78d-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be78d-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="be78d-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be78d-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="be78d-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="be78d-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="be78d-147">displayName :String</span></span>

<span data-ttu-id="be78d-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="be78d-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="be78d-149">Tipo</span><span class="sxs-lookup"><span data-stu-id="be78d-149">Type</span></span>

*   <span data-ttu-id="be78d-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="be78d-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be78d-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="be78d-151">Requirements</span></span>

|<span data-ttu-id="be78d-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="be78d-152">Requirement</span></span>| <span data-ttu-id="be78d-153">Valor</span><span class="sxs-lookup"><span data-stu-id="be78d-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="be78d-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="be78d-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be78d-155">1.0</span><span class="sxs-lookup"><span data-stu-id="be78d-155">1.0</span></span>|
|[<span data-ttu-id="be78d-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="be78d-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be78d-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be78d-157">ReadItem</span></span>|
|[<span data-ttu-id="be78d-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="be78d-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be78d-159">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="be78d-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be78d-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="be78d-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="be78d-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="be78d-161">emailAddress :String</span></span>

<span data-ttu-id="be78d-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="be78d-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="be78d-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="be78d-163">Type</span></span>

*   <span data-ttu-id="be78d-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="be78d-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be78d-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="be78d-165">Requirements</span></span>

|<span data-ttu-id="be78d-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="be78d-166">Requirement</span></span>| <span data-ttu-id="be78d-167">Valor</span><span class="sxs-lookup"><span data-stu-id="be78d-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="be78d-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="be78d-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be78d-169">1.0</span><span class="sxs-lookup"><span data-stu-id="be78d-169">1.0</span></span>|
|[<span data-ttu-id="be78d-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="be78d-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be78d-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be78d-171">ReadItem</span></span>|
|[<span data-ttu-id="be78d-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="be78d-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be78d-173">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="be78d-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be78d-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="be78d-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="be78d-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="be78d-175">timeZone :String</span></span>

<span data-ttu-id="be78d-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="be78d-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="be78d-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="be78d-177">Type</span></span>

*   <span data-ttu-id="be78d-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="be78d-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be78d-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="be78d-179">Requirements</span></span>

|<span data-ttu-id="be78d-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="be78d-180">Requirement</span></span>| <span data-ttu-id="be78d-181">Valor</span><span class="sxs-lookup"><span data-stu-id="be78d-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="be78d-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="be78d-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be78d-183">1.0</span><span class="sxs-lookup"><span data-stu-id="be78d-183">1.0</span></span>|
|[<span data-ttu-id="be78d-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="be78d-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be78d-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be78d-185">ReadItem</span></span>|
|[<span data-ttu-id="be78d-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="be78d-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be78d-187">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="be78d-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be78d-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="be78d-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
