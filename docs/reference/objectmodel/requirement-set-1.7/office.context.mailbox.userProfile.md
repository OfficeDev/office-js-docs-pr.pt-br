---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.7
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 866bf063cf4ad8bf040753714986a7b2db05b6d6
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433856"
---
# <a name="userprofile"></a><span data-ttu-id="9413a-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="9413a-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="9413a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="9413a-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="9413a-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9413a-104">Requirements</span></span>

|<span data-ttu-id="9413a-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="9413a-105">Requirement</span></span>| <span data-ttu-id="9413a-106">Valor</span><span class="sxs-lookup"><span data-stu-id="9413a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9413a-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9413a-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9413a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9413a-108">1.0</span></span>|
|[<span data-ttu-id="9413a-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9413a-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9413a-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9413a-110">ReadItem</span></span>|
|[<span data-ttu-id="9413a-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9413a-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9413a-112">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9413a-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9413a-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="9413a-113">Members and methods</span></span>

| <span data-ttu-id="9413a-114">Membro</span><span class="sxs-lookup"><span data-stu-id="9413a-114">Member</span></span> | <span data-ttu-id="9413a-115">Type</span><span class="sxs-lookup"><span data-stu-id="9413a-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9413a-116">accountType</span><span class="sxs-lookup"><span data-stu-id="9413a-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="9413a-117">Member</span><span class="sxs-lookup"><span data-stu-id="9413a-117">Member</span></span> |
| [<span data-ttu-id="9413a-118">displayName</span><span class="sxs-lookup"><span data-stu-id="9413a-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="9413a-119">Membro</span><span class="sxs-lookup"><span data-stu-id="9413a-119">Member</span></span> |
| [<span data-ttu-id="9413a-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="9413a-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="9413a-121">Membro</span><span class="sxs-lookup"><span data-stu-id="9413a-121">Member</span></span> |
| [<span data-ttu-id="9413a-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="9413a-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="9413a-123">Membro</span><span class="sxs-lookup"><span data-stu-id="9413a-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="9413a-124">Members</span><span class="sxs-lookup"><span data-stu-id="9413a-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="9413a-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="9413a-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="9413a-126">No momento, esse membro só tem suporte no Outlook 2016 para Mac, build 16.9.1212 e superior.</span><span class="sxs-lookup"><span data-stu-id="9413a-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="9413a-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="9413a-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="9413a-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="9413a-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="9413a-129">Value</span><span class="sxs-lookup"><span data-stu-id="9413a-129">Value</span></span> | <span data-ttu-id="9413a-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="9413a-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="9413a-131">A caixa de correio está em um servidor local do Exchange.</span><span class="sxs-lookup"><span data-stu-id="9413a-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="9413a-132">A caixa de correio está associada a uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="9413a-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="9413a-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="9413a-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="9413a-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="9413a-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="9413a-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9413a-135">Type:</span></span>

*   <span data-ttu-id="9413a-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9413a-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9413a-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9413a-137">Requirements</span></span>

|<span data-ttu-id="9413a-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="9413a-138">Requirement</span></span>| <span data-ttu-id="9413a-139">Valor</span><span class="sxs-lookup"><span data-stu-id="9413a-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="9413a-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9413a-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9413a-141">1.6</span><span class="sxs-lookup"><span data-stu-id="9413a-141">1.6</span></span> |
|[<span data-ttu-id="9413a-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9413a-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9413a-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9413a-143">ReadItem</span></span>|
|[<span data-ttu-id="9413a-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9413a-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9413a-145">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9413a-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9413a-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9413a-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="9413a-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="9413a-147">displayName :String</span></span>

<span data-ttu-id="9413a-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="9413a-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="9413a-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9413a-149">Type:</span></span>

*   <span data-ttu-id="9413a-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9413a-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9413a-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9413a-151">Requirements</span></span>

|<span data-ttu-id="9413a-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="9413a-152">Requirement</span></span>| <span data-ttu-id="9413a-153">Valor</span><span class="sxs-lookup"><span data-stu-id="9413a-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="9413a-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9413a-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9413a-155">1.0</span><span class="sxs-lookup"><span data-stu-id="9413a-155">1.0</span></span>|
|[<span data-ttu-id="9413a-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9413a-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9413a-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9413a-157">ReadItem</span></span>|
|[<span data-ttu-id="9413a-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9413a-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9413a-159">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9413a-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9413a-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9413a-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="9413a-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="9413a-161">emailAddress :String</span></span>

<span data-ttu-id="9413a-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="9413a-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="9413a-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9413a-163">Type:</span></span>

*   <span data-ttu-id="9413a-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9413a-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9413a-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9413a-165">Requirements</span></span>

|<span data-ttu-id="9413a-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="9413a-166">Requirement</span></span>| <span data-ttu-id="9413a-167">Valor</span><span class="sxs-lookup"><span data-stu-id="9413a-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="9413a-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9413a-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9413a-169">1.0</span><span class="sxs-lookup"><span data-stu-id="9413a-169">1.0</span></span>|
|[<span data-ttu-id="9413a-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9413a-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9413a-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9413a-171">ReadItem</span></span>|
|[<span data-ttu-id="9413a-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9413a-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9413a-173">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9413a-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9413a-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9413a-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="9413a-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="9413a-175">timeZone :String</span></span>

<span data-ttu-id="9413a-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="9413a-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="9413a-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="9413a-177">Type:</span></span>

*   <span data-ttu-id="9413a-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="9413a-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9413a-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="9413a-179">Requirements</span></span>

|<span data-ttu-id="9413a-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="9413a-180">Requirement</span></span>| <span data-ttu-id="9413a-181">Valor</span><span class="sxs-lookup"><span data-stu-id="9413a-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="9413a-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="9413a-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9413a-183">1.0</span><span class="sxs-lookup"><span data-stu-id="9413a-183">1.0</span></span>|
|[<span data-ttu-id="9413a-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="9413a-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9413a-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9413a-185">ReadItem</span></span>|
|[<span data-ttu-id="9413a-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="9413a-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9413a-187">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="9413a-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="9413a-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="9413a-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```