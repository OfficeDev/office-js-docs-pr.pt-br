---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,7
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 45533fb3a879e4e34e91adfb04dd8ce55f815749
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127139"
---
# <a name="userprofile"></a><span data-ttu-id="24e03-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="24e03-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="24e03-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="24e03-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="24e03-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24e03-104">Requirements</span></span>

|<span data-ttu-id="24e03-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="24e03-105">Requirement</span></span>| <span data-ttu-id="24e03-106">Valor</span><span class="sxs-lookup"><span data-stu-id="24e03-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="24e03-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24e03-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24e03-108">1.0</span><span class="sxs-lookup"><span data-stu-id="24e03-108">1.0</span></span>|
|[<span data-ttu-id="24e03-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24e03-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24e03-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24e03-110">ReadItem</span></span>|
|[<span data-ttu-id="24e03-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24e03-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24e03-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24e03-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="24e03-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="24e03-113">Members and methods</span></span>

| <span data-ttu-id="24e03-114">Membro</span><span class="sxs-lookup"><span data-stu-id="24e03-114">Member</span></span> | <span data-ttu-id="24e03-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="24e03-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="24e03-116">accountType</span><span class="sxs-lookup"><span data-stu-id="24e03-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="24e03-117">Membro</span><span class="sxs-lookup"><span data-stu-id="24e03-117">Member</span></span> |
| [<span data-ttu-id="24e03-118">displayName</span><span class="sxs-lookup"><span data-stu-id="24e03-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="24e03-119">Membro</span><span class="sxs-lookup"><span data-stu-id="24e03-119">Member</span></span> |
| [<span data-ttu-id="24e03-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="24e03-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="24e03-121">Membro</span><span class="sxs-lookup"><span data-stu-id="24e03-121">Member</span></span> |
| [<span data-ttu-id="24e03-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="24e03-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="24e03-123">Membro</span><span class="sxs-lookup"><span data-stu-id="24e03-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="24e03-124">Membros</span><span class="sxs-lookup"><span data-stu-id="24e03-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="24e03-125">AccountType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="24e03-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="24e03-126">Atualmente, esse membro só tem suporte no Outlook 2016 ou posterior no Mac (Build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="24e03-126">This member is currently only supported by Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="24e03-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="24e03-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="24e03-128">Os valores possíveis estão listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="24e03-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="24e03-129">Valor</span><span class="sxs-lookup"><span data-stu-id="24e03-129">Value</span></span> | <span data-ttu-id="24e03-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="24e03-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="24e03-131">A caixa de correio está em um servidor Exchange local.</span><span class="sxs-lookup"><span data-stu-id="24e03-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="24e03-132">A caixa de correio está associada a uma conta do gmail.</span><span class="sxs-lookup"><span data-stu-id="24e03-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="24e03-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="24e03-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="24e03-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="24e03-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="24e03-135">Tipo</span><span class="sxs-lookup"><span data-stu-id="24e03-135">Type</span></span>

*   <span data-ttu-id="24e03-136">String</span><span class="sxs-lookup"><span data-stu-id="24e03-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="24e03-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24e03-137">Requirements</span></span>

|<span data-ttu-id="24e03-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="24e03-138">Requirement</span></span>| <span data-ttu-id="24e03-139">Valor</span><span class="sxs-lookup"><span data-stu-id="24e03-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="24e03-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24e03-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24e03-141">1.6</span><span class="sxs-lookup"><span data-stu-id="24e03-141">1.6</span></span> |
|[<span data-ttu-id="24e03-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24e03-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24e03-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24e03-143">ReadItem</span></span>|
|[<span data-ttu-id="24e03-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24e03-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24e03-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24e03-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24e03-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24e03-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

#### <a name="displayname-string"></a><span data-ttu-id="24e03-147">displayName: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="24e03-147">displayName: String</span></span>

<span data-ttu-id="24e03-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="24e03-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="24e03-149">Tipo</span><span class="sxs-lookup"><span data-stu-id="24e03-149">Type</span></span>

*   <span data-ttu-id="24e03-150">String</span><span class="sxs-lookup"><span data-stu-id="24e03-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="24e03-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24e03-151">Requirements</span></span>

|<span data-ttu-id="24e03-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="24e03-152">Requirement</span></span>| <span data-ttu-id="24e03-153">Valor</span><span class="sxs-lookup"><span data-stu-id="24e03-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="24e03-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24e03-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24e03-155">1.0</span><span class="sxs-lookup"><span data-stu-id="24e03-155">1.0</span></span>|
|[<span data-ttu-id="24e03-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24e03-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24e03-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24e03-157">ReadItem</span></span>|
|[<span data-ttu-id="24e03-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24e03-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24e03-159">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24e03-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24e03-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24e03-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="24e03-161">emailAddress: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="24e03-161">emailAddress: String</span></span>

<span data-ttu-id="24e03-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="24e03-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="24e03-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="24e03-163">Type</span></span>

*   <span data-ttu-id="24e03-164">String</span><span class="sxs-lookup"><span data-stu-id="24e03-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="24e03-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24e03-165">Requirements</span></span>

|<span data-ttu-id="24e03-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="24e03-166">Requirement</span></span>| <span data-ttu-id="24e03-167">Valor</span><span class="sxs-lookup"><span data-stu-id="24e03-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="24e03-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24e03-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24e03-169">1.0</span><span class="sxs-lookup"><span data-stu-id="24e03-169">1.0</span></span>|
|[<span data-ttu-id="24e03-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24e03-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24e03-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24e03-171">ReadItem</span></span>|
|[<span data-ttu-id="24e03-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24e03-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24e03-173">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24e03-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24e03-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24e03-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

#### <a name="timezone-string"></a><span data-ttu-id="24e03-175">timeZone: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="24e03-175">timeZone: String</span></span>

<span data-ttu-id="24e03-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="24e03-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="24e03-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="24e03-177">Type</span></span>

*   <span data-ttu-id="24e03-178">String</span><span class="sxs-lookup"><span data-stu-id="24e03-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="24e03-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="24e03-179">Requirements</span></span>

|<span data-ttu-id="24e03-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="24e03-180">Requirement</span></span>| <span data-ttu-id="24e03-181">Valor</span><span class="sxs-lookup"><span data-stu-id="24e03-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="24e03-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="24e03-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24e03-183">1.0</span><span class="sxs-lookup"><span data-stu-id="24e03-183">1.0</span></span>|
|[<span data-ttu-id="24e03-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="24e03-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24e03-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="24e03-185">ReadItem</span></span>|
|[<span data-ttu-id="24e03-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="24e03-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24e03-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="24e03-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24e03-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="24e03-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
