---
title: Office.context.mailbox.userProfile – conjunto de requisitos 1.7
description: ''
ms.date: 10/31/2018
localization_priority: Normal
ms.openlocfilehash: b07ff5bee3adc18cc1006bb574e373182b29f5fe
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/13/2019
ms.locfileid: "29635899"
---
# <a name="userprofile"></a><span data-ttu-id="ee5ca-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ee5ca-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ee5ca-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ee5ca-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5ca-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee5ca-104">Requirements</span></span>

|<span data-ttu-id="ee5ca-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee5ca-105">Requirement</span></span>| <span data-ttu-id="ee5ca-106">Valor</span><span class="sxs-lookup"><span data-stu-id="ee5ca-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5ca-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee5ca-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee5ca-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ee5ca-108">1.0</span></span>|
|[<span data-ttu-id="ee5ca-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee5ca-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee5ca-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee5ca-110">ReadItem</span></span>|
|[<span data-ttu-id="ee5ca-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee5ca-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee5ca-112">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee5ca-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ee5ca-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="ee5ca-113">Members and methods</span></span>

| <span data-ttu-id="ee5ca-114">Membro</span><span class="sxs-lookup"><span data-stu-id="ee5ca-114">Member</span></span> | <span data-ttu-id="ee5ca-115">Type</span><span class="sxs-lookup"><span data-stu-id="ee5ca-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ee5ca-116">accountType</span><span class="sxs-lookup"><span data-stu-id="ee5ca-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="ee5ca-117">Member</span><span class="sxs-lookup"><span data-stu-id="ee5ca-117">Member</span></span> |
| [<span data-ttu-id="ee5ca-118">displayName</span><span class="sxs-lookup"><span data-stu-id="ee5ca-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="ee5ca-119">Member</span><span class="sxs-lookup"><span data-stu-id="ee5ca-119">Member</span></span> |
| [<span data-ttu-id="ee5ca-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="ee5ca-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="ee5ca-121">Membro</span><span class="sxs-lookup"><span data-stu-id="ee5ca-121">Member</span></span> |
| [<span data-ttu-id="ee5ca-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="ee5ca-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="ee5ca-123">Membro</span><span class="sxs-lookup"><span data-stu-id="ee5ca-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="ee5ca-124">Members</span><span class="sxs-lookup"><span data-stu-id="ee5ca-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="ee5ca-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="ee5ca-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="ee5ca-126">Este membro é atualmente com suporte apenas 2016 do Outlook para Mac (build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="ee5ca-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="ee5ca-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="ee5ca-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="ee5ca-128">Os valores possíveis são listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee5ca-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="ee5ca-129">Value</span><span class="sxs-lookup"><span data-stu-id="ee5ca-129">Value</span></span> | <span data-ttu-id="ee5ca-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee5ca-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="ee5ca-131">A caixa de correio está em um servidor local do Exchange.</span><span class="sxs-lookup"><span data-stu-id="ee5ca-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="ee5ca-132">A caixa de correio está associada a uma conta do Gmail.</span><span class="sxs-lookup"><span data-stu-id="ee5ca-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="ee5ca-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="ee5ca-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="ee5ca-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="ee5ca-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="ee5ca-135">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee5ca-135">Type:</span></span>

*   <span data-ttu-id="ee5ca-136">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ee5ca-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5ca-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee5ca-137">Requirements</span></span>

|<span data-ttu-id="ee5ca-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee5ca-138">Requirement</span></span>| <span data-ttu-id="ee5ca-139">Valor</span><span class="sxs-lookup"><span data-stu-id="ee5ca-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5ca-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee5ca-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee5ca-141">1.6</span><span class="sxs-lookup"><span data-stu-id="ee5ca-141">1.6</span></span> |
|[<span data-ttu-id="ee5ca-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee5ca-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee5ca-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee5ca-143">ReadItem</span></span>|
|[<span data-ttu-id="ee5ca-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee5ca-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee5ca-145">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee5ca-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5ca-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee5ca-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="ee5ca-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ee5ca-147">displayName :String</span></span>

<span data-ttu-id="ee5ca-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="ee5ca-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5ca-149">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee5ca-149">Type:</span></span>

*   <span data-ttu-id="ee5ca-150">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ee5ca-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5ca-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee5ca-151">Requirements</span></span>

|<span data-ttu-id="ee5ca-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee5ca-152">Requirement</span></span>| <span data-ttu-id="ee5ca-153">Valor</span><span class="sxs-lookup"><span data-stu-id="ee5ca-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5ca-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee5ca-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee5ca-155">1.0</span><span class="sxs-lookup"><span data-stu-id="ee5ca-155">1.0</span></span>|
|[<span data-ttu-id="ee5ca-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee5ca-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee5ca-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee5ca-157">ReadItem</span></span>|
|[<span data-ttu-id="ee5ca-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee5ca-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee5ca-159">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee5ca-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5ca-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee5ca-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ee5ca-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ee5ca-161">emailAddress :String</span></span>

<span data-ttu-id="ee5ca-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="ee5ca-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5ca-163">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee5ca-163">Type:</span></span>

*   <span data-ttu-id="ee5ca-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ee5ca-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5ca-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee5ca-165">Requirements</span></span>

|<span data-ttu-id="ee5ca-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee5ca-166">Requirement</span></span>| <span data-ttu-id="ee5ca-167">Valor</span><span class="sxs-lookup"><span data-stu-id="ee5ca-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5ca-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee5ca-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee5ca-169">1.0</span><span class="sxs-lookup"><span data-stu-id="ee5ca-169">1.0</span></span>|
|[<span data-ttu-id="ee5ca-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee5ca-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee5ca-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee5ca-171">ReadItem</span></span>|
|[<span data-ttu-id="ee5ca-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee5ca-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee5ca-173">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee5ca-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5ca-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee5ca-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ee5ca-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ee5ca-175">timeZone :String</span></span>

<span data-ttu-id="ee5ca-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="ee5ca-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5ca-177">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ee5ca-177">Type:</span></span>

*   <span data-ttu-id="ee5ca-178">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ee5ca-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5ca-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ee5ca-179">Requirements</span></span>

|<span data-ttu-id="ee5ca-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="ee5ca-180">Requirement</span></span>| <span data-ttu-id="ee5ca-181">Valor</span><span class="sxs-lookup"><span data-stu-id="ee5ca-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5ca-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ee5ca-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee5ca-183">1.0</span><span class="sxs-lookup"><span data-stu-id="ee5ca-183">1.0</span></span>|
|[<span data-ttu-id="ee5ca-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="ee5ca-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee5ca-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee5ca-185">ReadItem</span></span>|
|[<span data-ttu-id="ee5ca-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ee5ca-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee5ca-187">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="ee5ca-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5ca-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee5ca-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
