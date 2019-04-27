---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,7
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 036f18e4cb98cfe510a19d85a5a79f393ca8bd17
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353290"
---
# <a name="userprofile"></a><span data-ttu-id="43e2c-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="43e2c-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="43e2c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="43e2c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="43e2c-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43e2c-104">Requirements</span></span>

|<span data-ttu-id="43e2c-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="43e2c-105">Requirement</span></span>| <span data-ttu-id="43e2c-106">Valor</span><span class="sxs-lookup"><span data-stu-id="43e2c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="43e2c-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43e2c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43e2c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="43e2c-108">1.0</span></span>|
|[<span data-ttu-id="43e2c-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43e2c-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43e2c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43e2c-110">ReadItem</span></span>|
|[<span data-ttu-id="43e2c-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43e2c-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43e2c-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43e2c-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="43e2c-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="43e2c-113">Members and methods</span></span>

| <span data-ttu-id="43e2c-114">Membro</span><span class="sxs-lookup"><span data-stu-id="43e2c-114">Member</span></span> | <span data-ttu-id="43e2c-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="43e2c-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="43e2c-116">accountType</span><span class="sxs-lookup"><span data-stu-id="43e2c-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="43e2c-117">Member</span><span class="sxs-lookup"><span data-stu-id="43e2c-117">Member</span></span> |
| [<span data-ttu-id="43e2c-118">displayName</span><span class="sxs-lookup"><span data-stu-id="43e2c-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="43e2c-119">Member</span><span class="sxs-lookup"><span data-stu-id="43e2c-119">Member</span></span> |
| [<span data-ttu-id="43e2c-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="43e2c-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="43e2c-121">Member</span><span class="sxs-lookup"><span data-stu-id="43e2c-121">Member</span></span> |
| [<span data-ttu-id="43e2c-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="43e2c-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="43e2c-123">Membro</span><span class="sxs-lookup"><span data-stu-id="43e2c-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="43e2c-124">Membros</span><span class="sxs-lookup"><span data-stu-id="43e2c-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="43e2c-125">AccountType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="43e2c-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="43e2c-126">Atualmente, esse membro só tem suporte no Outlook 2016 para Mac (Build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="43e2c-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="43e2c-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="43e2c-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="43e2c-128">Os valores possíveis estão listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="43e2c-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="43e2c-129">Valor</span><span class="sxs-lookup"><span data-stu-id="43e2c-129">Value</span></span> | <span data-ttu-id="43e2c-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="43e2c-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="43e2c-131">A caixa de correio está em um servidor Exchange local.</span><span class="sxs-lookup"><span data-stu-id="43e2c-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="43e2c-132">A caixa de correio está associada a uma conta do gmail.</span><span class="sxs-lookup"><span data-stu-id="43e2c-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="43e2c-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="43e2c-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="43e2c-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="43e2c-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="43e2c-135">Tipo</span><span class="sxs-lookup"><span data-stu-id="43e2c-135">Type</span></span>

*   <span data-ttu-id="43e2c-136">String</span><span class="sxs-lookup"><span data-stu-id="43e2c-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43e2c-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43e2c-137">Requirements</span></span>

|<span data-ttu-id="43e2c-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="43e2c-138">Requirement</span></span>| <span data-ttu-id="43e2c-139">Valor</span><span class="sxs-lookup"><span data-stu-id="43e2c-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="43e2c-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43e2c-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43e2c-141">1.6</span><span class="sxs-lookup"><span data-stu-id="43e2c-141">1.6</span></span> |
|[<span data-ttu-id="43e2c-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43e2c-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43e2c-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43e2c-143">ReadItem</span></span>|
|[<span data-ttu-id="43e2c-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43e2c-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43e2c-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43e2c-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43e2c-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43e2c-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

#### <a name="displayname-string"></a><span data-ttu-id="43e2c-147">displayName: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="43e2c-147">displayName: String</span></span>

<span data-ttu-id="43e2c-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="43e2c-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="43e2c-149">Tipo</span><span class="sxs-lookup"><span data-stu-id="43e2c-149">Type</span></span>

*   <span data-ttu-id="43e2c-150">String</span><span class="sxs-lookup"><span data-stu-id="43e2c-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43e2c-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43e2c-151">Requirements</span></span>

|<span data-ttu-id="43e2c-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="43e2c-152">Requirement</span></span>| <span data-ttu-id="43e2c-153">Valor</span><span class="sxs-lookup"><span data-stu-id="43e2c-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="43e2c-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43e2c-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43e2c-155">1.0</span><span class="sxs-lookup"><span data-stu-id="43e2c-155">1.0</span></span>|
|[<span data-ttu-id="43e2c-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43e2c-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43e2c-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43e2c-157">ReadItem</span></span>|
|[<span data-ttu-id="43e2c-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43e2c-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43e2c-159">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43e2c-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43e2c-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43e2c-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="43e2c-161">emailAddress: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="43e2c-161">emailAddress: String</span></span>

<span data-ttu-id="43e2c-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="43e2c-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="43e2c-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="43e2c-163">Type</span></span>

*   <span data-ttu-id="43e2c-164">String</span><span class="sxs-lookup"><span data-stu-id="43e2c-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43e2c-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43e2c-165">Requirements</span></span>

|<span data-ttu-id="43e2c-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="43e2c-166">Requirement</span></span>| <span data-ttu-id="43e2c-167">Valor</span><span class="sxs-lookup"><span data-stu-id="43e2c-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="43e2c-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43e2c-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43e2c-169">1.0</span><span class="sxs-lookup"><span data-stu-id="43e2c-169">1.0</span></span>|
|[<span data-ttu-id="43e2c-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43e2c-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43e2c-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43e2c-171">ReadItem</span></span>|
|[<span data-ttu-id="43e2c-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43e2c-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43e2c-173">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43e2c-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43e2c-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43e2c-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

#### <a name="timezone-string"></a><span data-ttu-id="43e2c-175">timeZone: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="43e2c-175">timeZone: String</span></span>

<span data-ttu-id="43e2c-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="43e2c-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="43e2c-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="43e2c-177">Type</span></span>

*   <span data-ttu-id="43e2c-178">String</span><span class="sxs-lookup"><span data-stu-id="43e2c-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43e2c-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43e2c-179">Requirements</span></span>

|<span data-ttu-id="43e2c-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="43e2c-180">Requirement</span></span>| <span data-ttu-id="43e2c-181">Valor</span><span class="sxs-lookup"><span data-stu-id="43e2c-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="43e2c-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43e2c-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43e2c-183">1.0</span><span class="sxs-lookup"><span data-stu-id="43e2c-183">1.0</span></span>|
|[<span data-ttu-id="43e2c-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="43e2c-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43e2c-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43e2c-185">ReadItem</span></span>|
|[<span data-ttu-id="43e2c-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43e2c-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43e2c-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43e2c-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="43e2c-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="43e2c-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
