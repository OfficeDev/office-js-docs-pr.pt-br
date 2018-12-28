---
title: 'Office.context.mailbox.userProfile: conjunto de requisitos da visualização'
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 061ee8367005f4af0795c4d9e1236d0b2443521a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432813"
---
# <a name="userprofile"></a><span data-ttu-id="dc2cc-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="dc2cc-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="dc2cc-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="dc2cc-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc2cc-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dc2cc-104">Requirements</span></span>

|<span data-ttu-id="dc2cc-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="dc2cc-105">Requirement</span></span>| <span data-ttu-id="dc2cc-106">Valor</span><span class="sxs-lookup"><span data-stu-id="dc2cc-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc2cc-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dc2cc-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc2cc-108">1.0</span><span class="sxs-lookup"><span data-stu-id="dc2cc-108">1.0</span></span>|
|[<span data-ttu-id="dc2cc-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="dc2cc-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc2cc-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc2cc-110">ReadItem</span></span>|
|[<span data-ttu-id="dc2cc-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dc2cc-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dc2cc-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="dc2cc-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="dc2cc-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="dc2cc-113">Members and methods</span></span>

| <span data-ttu-id="dc2cc-114">Membro</span><span class="sxs-lookup"><span data-stu-id="dc2cc-114">Member</span></span> | <span data-ttu-id="dc2cc-115">Type</span><span class="sxs-lookup"><span data-stu-id="dc2cc-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="dc2cc-116">accountType</span><span class="sxs-lookup"><span data-stu-id="dc2cc-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="dc2cc-117">Member</span><span class="sxs-lookup"><span data-stu-id="dc2cc-117">Member</span></span> |
| [<span data-ttu-id="dc2cc-118">displayName</span><span class="sxs-lookup"><span data-stu-id="dc2cc-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="dc2cc-119">Membro</span><span class="sxs-lookup"><span data-stu-id="dc2cc-119">Member</span></span> |
| [<span data-ttu-id="dc2cc-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="dc2cc-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="dc2cc-121">Membro</span><span class="sxs-lookup"><span data-stu-id="dc2cc-121">Member</span></span> |
| [<span data-ttu-id="dc2cc-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="dc2cc-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="dc2cc-123">Membro</span><span class="sxs-lookup"><span data-stu-id="dc2cc-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="dc2cc-124">Members</span><span class="sxs-lookup"><span data-stu-id="dc2cc-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="dc2cc-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="dc2cc-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="dc2cc-126">No momento, esse membro só tem suporte no Outlook 2016 ou posterior para Mac (build 16.9.1212 e posterior).</span><span class="sxs-lookup"><span data-stu-id="dc2cc-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="dc2cc-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="dc2cc-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="dc2cc-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="dc2cc-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="dc2cc-129">Value</span><span class="sxs-lookup"><span data-stu-id="dc2cc-129">Value</span></span> | <span data-ttu-id="dc2cc-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="dc2cc-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="dc2cc-131">A caixa de correio está em um servidor local do Exchange.</span><span class="sxs-lookup"><span data-stu-id="dc2cc-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="dc2cc-132">A caixa de correio está associada a uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="dc2cc-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="dc2cc-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="dc2cc-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="dc2cc-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="dc2cc-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="dc2cc-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="dc2cc-135">Type:</span></span>

*   <span data-ttu-id="dc2cc-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dc2cc-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc2cc-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dc2cc-137">Requirements</span></span>

|<span data-ttu-id="dc2cc-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="dc2cc-138">Requirement</span></span>| <span data-ttu-id="dc2cc-139">Valor</span><span class="sxs-lookup"><span data-stu-id="dc2cc-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc2cc-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dc2cc-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc2cc-141">1.6</span><span class="sxs-lookup"><span data-stu-id="dc2cc-141">1.6</span></span> |
|[<span data-ttu-id="dc2cc-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="dc2cc-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc2cc-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc2cc-143">ReadItem</span></span>|
|[<span data-ttu-id="dc2cc-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dc2cc-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dc2cc-145">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="dc2cc-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="dc2cc-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dc2cc-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="dc2cc-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="dc2cc-147">displayName :String</span></span>

<span data-ttu-id="dc2cc-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="dc2cc-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="dc2cc-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="dc2cc-149">Type:</span></span>

*   <span data-ttu-id="dc2cc-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dc2cc-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc2cc-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dc2cc-151">Requirements</span></span>

|<span data-ttu-id="dc2cc-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="dc2cc-152">Requirement</span></span>| <span data-ttu-id="dc2cc-153">Valor</span><span class="sxs-lookup"><span data-stu-id="dc2cc-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc2cc-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dc2cc-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc2cc-155">1.0</span><span class="sxs-lookup"><span data-stu-id="dc2cc-155">1.0</span></span>|
|[<span data-ttu-id="dc2cc-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="dc2cc-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc2cc-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc2cc-157">ReadItem</span></span>|
|[<span data-ttu-id="dc2cc-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dc2cc-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dc2cc-159">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="dc2cc-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="dc2cc-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dc2cc-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="dc2cc-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="dc2cc-161">emailAddress :String</span></span>

<span data-ttu-id="dc2cc-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="dc2cc-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="dc2cc-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="dc2cc-163">Type:</span></span>

*   <span data-ttu-id="dc2cc-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dc2cc-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc2cc-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dc2cc-165">Requirements</span></span>

|<span data-ttu-id="dc2cc-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="dc2cc-166">Requirement</span></span>| <span data-ttu-id="dc2cc-167">Valor</span><span class="sxs-lookup"><span data-stu-id="dc2cc-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc2cc-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dc2cc-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc2cc-169">1.0</span><span class="sxs-lookup"><span data-stu-id="dc2cc-169">1.0</span></span>|
|[<span data-ttu-id="dc2cc-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="dc2cc-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc2cc-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc2cc-171">ReadItem</span></span>|
|[<span data-ttu-id="dc2cc-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dc2cc-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dc2cc-173">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="dc2cc-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="dc2cc-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dc2cc-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="dc2cc-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="dc2cc-175">timeZone :String</span></span>

<span data-ttu-id="dc2cc-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="dc2cc-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="dc2cc-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="dc2cc-177">Type:</span></span>

*   <span data-ttu-id="dc2cc-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="dc2cc-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc2cc-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="dc2cc-179">Requirements</span></span>

|<span data-ttu-id="dc2cc-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="dc2cc-180">Requirement</span></span>| <span data-ttu-id="dc2cc-181">Valor</span><span class="sxs-lookup"><span data-stu-id="dc2cc-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc2cc-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="dc2cc-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dc2cc-183">1.0</span><span class="sxs-lookup"><span data-stu-id="dc2cc-183">1.0</span></span>|
|[<span data-ttu-id="dc2cc-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="dc2cc-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dc2cc-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dc2cc-185">ReadItem</span></span>|
|[<span data-ttu-id="dc2cc-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="dc2cc-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="dc2cc-187">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="dc2cc-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="dc2cc-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="dc2cc-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```