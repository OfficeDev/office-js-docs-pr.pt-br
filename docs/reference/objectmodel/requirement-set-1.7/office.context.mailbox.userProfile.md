---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.7
description: ''
ms.date: 10/31/2018
localization_priority: Normal
ms.openlocfilehash: b07ff5bee3adc18cc1006bb574e373182b29f5fe
ms.sourcegitcommit: 2e4b97f0252ff3dd908a3aa7a9720f0cb50b855d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/30/2019
ms.locfileid: "29635899"
---
# <a name="userprofile"></a><span data-ttu-id="a31d7-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="a31d7-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="a31d7-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="a31d7-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a31d7-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a31d7-104">Requirements</span></span>

|<span data-ttu-id="a31d7-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="a31d7-105">Requirement</span></span>| <span data-ttu-id="a31d7-106">Valor</span><span class="sxs-lookup"><span data-stu-id="a31d7-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a31d7-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a31d7-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a31d7-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a31d7-108">1.0</span></span>|
|[<span data-ttu-id="a31d7-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a31d7-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a31d7-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a31d7-110">ReadItem</span></span>|
|[<span data-ttu-id="a31d7-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a31d7-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a31d7-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a31d7-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a31d7-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="a31d7-113">Members and methods</span></span>

| <span data-ttu-id="a31d7-114">Membro</span><span class="sxs-lookup"><span data-stu-id="a31d7-114">Member</span></span> | <span data-ttu-id="a31d7-115">Type</span><span class="sxs-lookup"><span data-stu-id="a31d7-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a31d7-116">accountType</span><span class="sxs-lookup"><span data-stu-id="a31d7-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="a31d7-117">Member</span><span class="sxs-lookup"><span data-stu-id="a31d7-117">Member</span></span> |
| [<span data-ttu-id="a31d7-118">displayName</span><span class="sxs-lookup"><span data-stu-id="a31d7-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="a31d7-119">Membro</span><span class="sxs-lookup"><span data-stu-id="a31d7-119">Member</span></span> |
| [<span data-ttu-id="a31d7-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="a31d7-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="a31d7-121">Membro</span><span class="sxs-lookup"><span data-stu-id="a31d7-121">Member</span></span> |
| [<span data-ttu-id="a31d7-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="a31d7-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="a31d7-123">Membro</span><span class="sxs-lookup"><span data-stu-id="a31d7-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="a31d7-124">Members</span><span class="sxs-lookup"><span data-stu-id="a31d7-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="a31d7-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="a31d7-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="a31d7-126">Este membro é atualmente com suporte apenas 2016 do Outlook para Mac (build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="a31d7-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="a31d7-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="a31d7-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="a31d7-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="a31d7-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="a31d7-129">Value</span><span class="sxs-lookup"><span data-stu-id="a31d7-129">Value</span></span> | <span data-ttu-id="a31d7-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="a31d7-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="a31d7-131">A caixa de correio está em um servidor local do Exchange.</span><span class="sxs-lookup"><span data-stu-id="a31d7-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="a31d7-132">A caixa de correio está associada a uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="a31d7-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="a31d7-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="a31d7-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="a31d7-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="a31d7-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="a31d7-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a31d7-135">Type:</span></span>

*   <span data-ttu-id="a31d7-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a31d7-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a31d7-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a31d7-137">Requirements</span></span>

|<span data-ttu-id="a31d7-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="a31d7-138">Requirement</span></span>| <span data-ttu-id="a31d7-139">Valor</span><span class="sxs-lookup"><span data-stu-id="a31d7-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="a31d7-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a31d7-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a31d7-141">1.6</span><span class="sxs-lookup"><span data-stu-id="a31d7-141">1.6</span></span> |
|[<span data-ttu-id="a31d7-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a31d7-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a31d7-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a31d7-143">ReadItem</span></span>|
|[<span data-ttu-id="a31d7-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a31d7-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a31d7-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a31d7-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a31d7-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a31d7-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="a31d7-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="a31d7-147">displayName :String</span></span>

<span data-ttu-id="a31d7-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="a31d7-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a31d7-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a31d7-149">Type:</span></span>

*   <span data-ttu-id="a31d7-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a31d7-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a31d7-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a31d7-151">Requirements</span></span>

|<span data-ttu-id="a31d7-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="a31d7-152">Requirement</span></span>| <span data-ttu-id="a31d7-153">Valor</span><span class="sxs-lookup"><span data-stu-id="a31d7-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="a31d7-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a31d7-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a31d7-155">1.0</span><span class="sxs-lookup"><span data-stu-id="a31d7-155">1.0</span></span>|
|[<span data-ttu-id="a31d7-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a31d7-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a31d7-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a31d7-157">ReadItem</span></span>|
|[<span data-ttu-id="a31d7-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a31d7-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a31d7-159">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a31d7-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a31d7-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a31d7-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a31d7-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="a31d7-161">emailAddress :String</span></span>

<span data-ttu-id="a31d7-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="a31d7-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a31d7-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a31d7-163">Type:</span></span>

*   <span data-ttu-id="a31d7-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a31d7-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a31d7-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a31d7-165">Requirements</span></span>

|<span data-ttu-id="a31d7-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="a31d7-166">Requirement</span></span>| <span data-ttu-id="a31d7-167">Valor</span><span class="sxs-lookup"><span data-stu-id="a31d7-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="a31d7-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a31d7-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a31d7-169">1.0</span><span class="sxs-lookup"><span data-stu-id="a31d7-169">1.0</span></span>|
|[<span data-ttu-id="a31d7-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a31d7-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a31d7-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a31d7-171">ReadItem</span></span>|
|[<span data-ttu-id="a31d7-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a31d7-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a31d7-173">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a31d7-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a31d7-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a31d7-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a31d7-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="a31d7-175">timeZone :String</span></span>

<span data-ttu-id="a31d7-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="a31d7-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a31d7-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="a31d7-177">Type:</span></span>

*   <span data-ttu-id="a31d7-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a31d7-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a31d7-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a31d7-179">Requirements</span></span>

|<span data-ttu-id="a31d7-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="a31d7-180">Requirement</span></span>| <span data-ttu-id="a31d7-181">Valor</span><span class="sxs-lookup"><span data-stu-id="a31d7-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="a31d7-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a31d7-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a31d7-183">1.0</span><span class="sxs-lookup"><span data-stu-id="a31d7-183">1.0</span></span>|
|[<span data-ttu-id="a31d7-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="a31d7-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a31d7-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a31d7-185">ReadItem</span></span>|
|[<span data-ttu-id="a31d7-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a31d7-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a31d7-187">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="a31d7-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a31d7-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="a31d7-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
