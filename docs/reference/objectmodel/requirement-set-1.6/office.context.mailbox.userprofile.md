---
title: 'Office.context.mailbox.userProfile: conjunto de requisitos da versão 1.6'
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: fe30a390583dc646e9c8792710c580d02c373a1a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432889"
---
# <a name="userprofile"></a><span data-ttu-id="5824e-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="5824e-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="5824e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="5824e-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="5824e-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5824e-104">Requirements</span></span>

|<span data-ttu-id="5824e-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="5824e-105">Requirement</span></span>| <span data-ttu-id="5824e-106">Valor</span><span class="sxs-lookup"><span data-stu-id="5824e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="5824e-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5824e-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5824e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="5824e-108">1.0</span></span>|
|[<span data-ttu-id="5824e-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="5824e-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5824e-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5824e-110">ReadItem</span></span>|
|[<span data-ttu-id="5824e-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5824e-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5824e-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="5824e-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5824e-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="5824e-113">Members and methods</span></span>

| <span data-ttu-id="5824e-114">Membro</span><span class="sxs-lookup"><span data-stu-id="5824e-114">Member</span></span> | <span data-ttu-id="5824e-115">Type</span><span class="sxs-lookup"><span data-stu-id="5824e-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5824e-116">accountType</span><span class="sxs-lookup"><span data-stu-id="5824e-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="5824e-117">Member</span><span class="sxs-lookup"><span data-stu-id="5824e-117">Member</span></span> |
| [<span data-ttu-id="5824e-118">displayName</span><span class="sxs-lookup"><span data-stu-id="5824e-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="5824e-119">Membro</span><span class="sxs-lookup"><span data-stu-id="5824e-119">Member</span></span> |
| [<span data-ttu-id="5824e-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="5824e-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="5824e-121">Membro</span><span class="sxs-lookup"><span data-stu-id="5824e-121">Member</span></span> |
| [<span data-ttu-id="5824e-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="5824e-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="5824e-123">Membro</span><span class="sxs-lookup"><span data-stu-id="5824e-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="5824e-124">Members</span><span class="sxs-lookup"><span data-stu-id="5824e-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="5824e-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="5824e-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="5824e-126">No momento, esse membro só tem suporte no Outlook 2016 ou posterior para Mac (build 16.9.1212 e posterior).</span><span class="sxs-lookup"><span data-stu-id="5824e-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="5824e-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="5824e-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="5824e-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="5824e-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="5824e-129">Value</span><span class="sxs-lookup"><span data-stu-id="5824e-129">Value</span></span> | <span data-ttu-id="5824e-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="5824e-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="5824e-131">A caixa de correio está em um servidor local do Exchange.</span><span class="sxs-lookup"><span data-stu-id="5824e-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="5824e-132">A caixa de correio está associada a uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="5824e-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="5824e-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="5824e-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="5824e-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="5824e-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="5824e-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="5824e-135">Type:</span></span>

*   <span data-ttu-id="5824e-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5824e-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5824e-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5824e-137">Requirements</span></span>

|<span data-ttu-id="5824e-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="5824e-138">Requirement</span></span>| <span data-ttu-id="5824e-139">Valor</span><span class="sxs-lookup"><span data-stu-id="5824e-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="5824e-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5824e-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5824e-141">1.6</span><span class="sxs-lookup"><span data-stu-id="5824e-141">1.6</span></span> |
|[<span data-ttu-id="5824e-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="5824e-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5824e-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5824e-143">ReadItem</span></span>|
|[<span data-ttu-id="5824e-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5824e-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5824e-145">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="5824e-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5824e-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="5824e-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="5824e-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="5824e-147">displayName :String</span></span>

<span data-ttu-id="5824e-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="5824e-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="5824e-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="5824e-149">Type:</span></span>

*   <span data-ttu-id="5824e-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5824e-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5824e-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5824e-151">Requirements</span></span>

|<span data-ttu-id="5824e-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="5824e-152">Requirement</span></span>| <span data-ttu-id="5824e-153">Valor</span><span class="sxs-lookup"><span data-stu-id="5824e-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="5824e-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5824e-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5824e-155">1.0</span><span class="sxs-lookup"><span data-stu-id="5824e-155">1.0</span></span>|
|[<span data-ttu-id="5824e-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="5824e-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5824e-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5824e-157">ReadItem</span></span>|
|[<span data-ttu-id="5824e-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5824e-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5824e-159">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="5824e-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5824e-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="5824e-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="5824e-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="5824e-161">emailAddress :String</span></span>

<span data-ttu-id="5824e-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="5824e-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="5824e-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="5824e-163">Type:</span></span>

*   <span data-ttu-id="5824e-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5824e-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5824e-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5824e-165">Requirements</span></span>

|<span data-ttu-id="5824e-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="5824e-166">Requirement</span></span>| <span data-ttu-id="5824e-167">Valor</span><span class="sxs-lookup"><span data-stu-id="5824e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="5824e-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5824e-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5824e-169">1.0</span><span class="sxs-lookup"><span data-stu-id="5824e-169">1.0</span></span>|
|[<span data-ttu-id="5824e-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="5824e-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5824e-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5824e-171">ReadItem</span></span>|
|[<span data-ttu-id="5824e-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5824e-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5824e-173">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="5824e-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5824e-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="5824e-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="5824e-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="5824e-175">timeZone :String</span></span>

<span data-ttu-id="5824e-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="5824e-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="5824e-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="5824e-177">Type:</span></span>

*   <span data-ttu-id="5824e-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5824e-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5824e-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5824e-179">Requirements</span></span>

|<span data-ttu-id="5824e-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="5824e-180">Requirement</span></span>| <span data-ttu-id="5824e-181">Valor</span><span class="sxs-lookup"><span data-stu-id="5824e-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="5824e-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5824e-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5824e-183">1.0</span><span class="sxs-lookup"><span data-stu-id="5824e-183">1.0</span></span>|
|[<span data-ttu-id="5824e-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="5824e-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5824e-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="5824e-185">ReadItem</span></span>|
|[<span data-ttu-id="5824e-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5824e-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5824e-187">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="5824e-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="5824e-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="5824e-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```