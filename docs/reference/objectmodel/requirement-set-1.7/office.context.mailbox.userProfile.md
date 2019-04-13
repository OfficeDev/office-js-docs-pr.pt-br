---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos 1,7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 8cfee874bbb5183d62cc3a9ce8b042a76617ec72
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838519"
---
# <a name="userprofile"></a><span data-ttu-id="d87eb-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="d87eb-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="d87eb-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="d87eb-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="d87eb-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d87eb-104">Requirements</span></span>

|<span data-ttu-id="d87eb-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="d87eb-105">Requirement</span></span>| <span data-ttu-id="d87eb-106">Valor</span><span class="sxs-lookup"><span data-stu-id="d87eb-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d87eb-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d87eb-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d87eb-108">1.0</span><span class="sxs-lookup"><span data-stu-id="d87eb-108">1.0</span></span>|
|[<span data-ttu-id="d87eb-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d87eb-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d87eb-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d87eb-110">ReadItem</span></span>|
|[<span data-ttu-id="d87eb-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d87eb-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d87eb-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d87eb-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d87eb-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="d87eb-113">Members and methods</span></span>

| <span data-ttu-id="d87eb-114">Membro</span><span class="sxs-lookup"><span data-stu-id="d87eb-114">Member</span></span> | <span data-ttu-id="d87eb-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="d87eb-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d87eb-116">accountType</span><span class="sxs-lookup"><span data-stu-id="d87eb-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="d87eb-117">Membro</span><span class="sxs-lookup"><span data-stu-id="d87eb-117">Member</span></span> |
| [<span data-ttu-id="d87eb-118">displayName</span><span class="sxs-lookup"><span data-stu-id="d87eb-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="d87eb-119">Membro</span><span class="sxs-lookup"><span data-stu-id="d87eb-119">Member</span></span> |
| [<span data-ttu-id="d87eb-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="d87eb-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="d87eb-121">Membro</span><span class="sxs-lookup"><span data-stu-id="d87eb-121">Member</span></span> |
| [<span data-ttu-id="d87eb-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="d87eb-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="d87eb-123">Membro</span><span class="sxs-lookup"><span data-stu-id="d87eb-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="d87eb-124">Membros</span><span class="sxs-lookup"><span data-stu-id="d87eb-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="d87eb-125">AccountType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d87eb-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="d87eb-126">Atualmente, esse membro só tem suporte no Outlook 2016 para Mac (Build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="d87eb-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="d87eb-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="d87eb-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="d87eb-128">Os valores possíveis estão listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="d87eb-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="d87eb-129">Valor</span><span class="sxs-lookup"><span data-stu-id="d87eb-129">Value</span></span> | <span data-ttu-id="d87eb-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="d87eb-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="d87eb-131">A caixa de correio está em um servidor Exchange local.</span><span class="sxs-lookup"><span data-stu-id="d87eb-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="d87eb-132">A caixa de correio está associada a uma conta do gmail.</span><span class="sxs-lookup"><span data-stu-id="d87eb-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="d87eb-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="d87eb-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="d87eb-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="d87eb-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="d87eb-135">Tipo</span><span class="sxs-lookup"><span data-stu-id="d87eb-135">Type</span></span>

*   <span data-ttu-id="d87eb-136">String</span><span class="sxs-lookup"><span data-stu-id="d87eb-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d87eb-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d87eb-137">Requirements</span></span>

|<span data-ttu-id="d87eb-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="d87eb-138">Requirement</span></span>| <span data-ttu-id="d87eb-139">Valor</span><span class="sxs-lookup"><span data-stu-id="d87eb-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="d87eb-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d87eb-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d87eb-141">1.6</span><span class="sxs-lookup"><span data-stu-id="d87eb-141">1.6</span></span> |
|[<span data-ttu-id="d87eb-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d87eb-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d87eb-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d87eb-143">ReadItem</span></span>|
|[<span data-ttu-id="d87eb-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d87eb-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d87eb-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d87eb-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d87eb-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d87eb-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

####  <a name="displayname-string"></a><span data-ttu-id="d87eb-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="d87eb-147">displayName :String</span></span>

<span data-ttu-id="d87eb-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="d87eb-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="d87eb-149">Tipo</span><span class="sxs-lookup"><span data-stu-id="d87eb-149">Type</span></span>

*   <span data-ttu-id="d87eb-150">String</span><span class="sxs-lookup"><span data-stu-id="d87eb-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d87eb-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d87eb-151">Requirements</span></span>

|<span data-ttu-id="d87eb-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="d87eb-152">Requirement</span></span>| <span data-ttu-id="d87eb-153">Valor</span><span class="sxs-lookup"><span data-stu-id="d87eb-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="d87eb-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d87eb-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d87eb-155">1.0</span><span class="sxs-lookup"><span data-stu-id="d87eb-155">1.0</span></span>|
|[<span data-ttu-id="d87eb-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d87eb-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d87eb-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d87eb-157">ReadItem</span></span>|
|[<span data-ttu-id="d87eb-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d87eb-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d87eb-159">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d87eb-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d87eb-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d87eb-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

####  <a name="emailaddress-string"></a><span data-ttu-id="d87eb-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="d87eb-161">emailAddress :String</span></span>

<span data-ttu-id="d87eb-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="d87eb-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="d87eb-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="d87eb-163">Type</span></span>

*   <span data-ttu-id="d87eb-164">String</span><span class="sxs-lookup"><span data-stu-id="d87eb-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d87eb-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d87eb-165">Requirements</span></span>

|<span data-ttu-id="d87eb-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="d87eb-166">Requirement</span></span>| <span data-ttu-id="d87eb-167">Valor</span><span class="sxs-lookup"><span data-stu-id="d87eb-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="d87eb-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d87eb-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d87eb-169">1.0</span><span class="sxs-lookup"><span data-stu-id="d87eb-169">1.0</span></span>|
|[<span data-ttu-id="d87eb-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d87eb-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d87eb-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d87eb-171">ReadItem</span></span>|
|[<span data-ttu-id="d87eb-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d87eb-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d87eb-173">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d87eb-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d87eb-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d87eb-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

####  <a name="timezone-string"></a><span data-ttu-id="d87eb-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="d87eb-175">timeZone :String</span></span>

<span data-ttu-id="d87eb-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="d87eb-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="d87eb-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="d87eb-177">Type</span></span>

*   <span data-ttu-id="d87eb-178">String</span><span class="sxs-lookup"><span data-stu-id="d87eb-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d87eb-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d87eb-179">Requirements</span></span>

|<span data-ttu-id="d87eb-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="d87eb-180">Requirement</span></span>| <span data-ttu-id="d87eb-181">Valor</span><span class="sxs-lookup"><span data-stu-id="d87eb-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="d87eb-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d87eb-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d87eb-183">1.0</span><span class="sxs-lookup"><span data-stu-id="d87eb-183">1.0</span></span>|
|[<span data-ttu-id="d87eb-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="d87eb-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d87eb-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d87eb-185">ReadItem</span></span>|
|[<span data-ttu-id="d87eb-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d87eb-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d87eb-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d87eb-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d87eb-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="d87eb-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
