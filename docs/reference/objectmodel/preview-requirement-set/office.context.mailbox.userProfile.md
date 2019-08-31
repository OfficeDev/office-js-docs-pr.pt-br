---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos de visualização
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 5941c4e1276535091a3ffcf5b2fb6aa972ed8c4d
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696467"
---
# <a name="userprofile"></a><span data-ttu-id="710b0-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="710b0-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="710b0-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="710b0-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="710b0-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="710b0-104">Requirements</span></span>

|<span data-ttu-id="710b0-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="710b0-105">Requirement</span></span>| <span data-ttu-id="710b0-106">Valor</span><span class="sxs-lookup"><span data-stu-id="710b0-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="710b0-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="710b0-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="710b0-108">1.0</span><span class="sxs-lookup"><span data-stu-id="710b0-108">1.0</span></span>|
|[<span data-ttu-id="710b0-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="710b0-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="710b0-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="710b0-110">ReadItem</span></span>|
|[<span data-ttu-id="710b0-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="710b0-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="710b0-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="710b0-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="710b0-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="710b0-113">Members and methods</span></span>

| <span data-ttu-id="710b0-114">Membro</span><span class="sxs-lookup"><span data-stu-id="710b0-114">Member</span></span> | <span data-ttu-id="710b0-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="710b0-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="710b0-116">accountType</span><span class="sxs-lookup"><span data-stu-id="710b0-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="710b0-117">Membro</span><span class="sxs-lookup"><span data-stu-id="710b0-117">Member</span></span> |
| [<span data-ttu-id="710b0-118">displayName</span><span class="sxs-lookup"><span data-stu-id="710b0-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="710b0-119">Membro</span><span class="sxs-lookup"><span data-stu-id="710b0-119">Member</span></span> |
| [<span data-ttu-id="710b0-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="710b0-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="710b0-121">Membro</span><span class="sxs-lookup"><span data-stu-id="710b0-121">Member</span></span> |
| [<span data-ttu-id="710b0-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="710b0-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="710b0-123">Membro</span><span class="sxs-lookup"><span data-stu-id="710b0-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="710b0-124">Membros</span><span class="sxs-lookup"><span data-stu-id="710b0-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="710b0-125">AccountType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="710b0-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="710b0-126">Atualmente, esse membro só tem suporte no Outlook 2016 ou posterior no Mac (Build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="710b0-126">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="710b0-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="710b0-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="710b0-128">Os valores possíveis estão listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="710b0-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="710b0-129">Valor</span><span class="sxs-lookup"><span data-stu-id="710b0-129">Value</span></span> | <span data-ttu-id="710b0-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="710b0-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="710b0-131">A caixa de correio está em um servidor Exchange local.</span><span class="sxs-lookup"><span data-stu-id="710b0-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="710b0-132">A caixa de correio está associada a uma conta do gmail.</span><span class="sxs-lookup"><span data-stu-id="710b0-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="710b0-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="710b0-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="710b0-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="710b0-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="710b0-135">Tipo</span><span class="sxs-lookup"><span data-stu-id="710b0-135">Type</span></span>

*   <span data-ttu-id="710b0-136">String</span><span class="sxs-lookup"><span data-stu-id="710b0-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="710b0-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="710b0-137">Requirements</span></span>

|<span data-ttu-id="710b0-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="710b0-138">Requirement</span></span>| <span data-ttu-id="710b0-139">Valor</span><span class="sxs-lookup"><span data-stu-id="710b0-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="710b0-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="710b0-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="710b0-141">1.6</span><span class="sxs-lookup"><span data-stu-id="710b0-141">1.6</span></span> |
|[<span data-ttu-id="710b0-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="710b0-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="710b0-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="710b0-143">ReadItem</span></span>|
|[<span data-ttu-id="710b0-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="710b0-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="710b0-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="710b0-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="710b0-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="710b0-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="710b0-147">displayName: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="710b0-147">displayName: String</span></span>

<span data-ttu-id="710b0-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="710b0-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="710b0-149">Tipo</span><span class="sxs-lookup"><span data-stu-id="710b0-149">Type</span></span>

*   <span data-ttu-id="710b0-150">String</span><span class="sxs-lookup"><span data-stu-id="710b0-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="710b0-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="710b0-151">Requirements</span></span>

|<span data-ttu-id="710b0-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="710b0-152">Requirement</span></span>| <span data-ttu-id="710b0-153">Valor</span><span class="sxs-lookup"><span data-stu-id="710b0-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="710b0-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="710b0-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="710b0-155">1.0</span><span class="sxs-lookup"><span data-stu-id="710b0-155">1.0</span></span>|
|[<span data-ttu-id="710b0-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="710b0-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="710b0-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="710b0-157">ReadItem</span></span>|
|[<span data-ttu-id="710b0-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="710b0-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="710b0-159">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="710b0-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="710b0-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="710b0-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="710b0-161">emailAddress: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="710b0-161">emailAddress: String</span></span>

<span data-ttu-id="710b0-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="710b0-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="710b0-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="710b0-163">Type</span></span>

*   <span data-ttu-id="710b0-164">String</span><span class="sxs-lookup"><span data-stu-id="710b0-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="710b0-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="710b0-165">Requirements</span></span>

|<span data-ttu-id="710b0-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="710b0-166">Requirement</span></span>| <span data-ttu-id="710b0-167">Valor</span><span class="sxs-lookup"><span data-stu-id="710b0-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="710b0-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="710b0-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="710b0-169">1.0</span><span class="sxs-lookup"><span data-stu-id="710b0-169">1.0</span></span>|
|[<span data-ttu-id="710b0-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="710b0-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="710b0-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="710b0-171">ReadItem</span></span>|
|[<span data-ttu-id="710b0-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="710b0-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="710b0-173">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="710b0-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="710b0-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="710b0-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="710b0-175">timeZone: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="710b0-175">timeZone: String</span></span>

<span data-ttu-id="710b0-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="710b0-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="710b0-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="710b0-177">Type</span></span>

*   <span data-ttu-id="710b0-178">String</span><span class="sxs-lookup"><span data-stu-id="710b0-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="710b0-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="710b0-179">Requirements</span></span>

|<span data-ttu-id="710b0-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="710b0-180">Requirement</span></span>| <span data-ttu-id="710b0-181">Valor</span><span class="sxs-lookup"><span data-stu-id="710b0-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="710b0-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="710b0-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="710b0-183">1.0</span><span class="sxs-lookup"><span data-stu-id="710b0-183">1.0</span></span>|
|[<span data-ttu-id="710b0-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="710b0-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="710b0-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="710b0-185">ReadItem</span></span>|
|[<span data-ttu-id="710b0-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="710b0-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="710b0-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="710b0-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="710b0-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="710b0-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
