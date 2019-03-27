---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos de visualização
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 204097497c958c26a6e67fc01d6dbd5142d8dced
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871868"
---
# <a name="userprofile"></a><span data-ttu-id="51c58-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="51c58-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="51c58-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="51c58-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="51c58-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="51c58-104">Requirements</span></span>

|<span data-ttu-id="51c58-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="51c58-105">Requirement</span></span>| <span data-ttu-id="51c58-106">Valor</span><span class="sxs-lookup"><span data-stu-id="51c58-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="51c58-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="51c58-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="51c58-108">1.0</span><span class="sxs-lookup"><span data-stu-id="51c58-108">1.0</span></span>|
|[<span data-ttu-id="51c58-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="51c58-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="51c58-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="51c58-110">ReadItem</span></span>|
|[<span data-ttu-id="51c58-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="51c58-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="51c58-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="51c58-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="51c58-113">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="51c58-113">Members and methods</span></span>

| <span data-ttu-id="51c58-114">Membro</span><span class="sxs-lookup"><span data-stu-id="51c58-114">Member</span></span> | <span data-ttu-id="51c58-115">Tipo</span><span class="sxs-lookup"><span data-stu-id="51c58-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="51c58-116">accountType</span><span class="sxs-lookup"><span data-stu-id="51c58-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="51c58-117">Member</span><span class="sxs-lookup"><span data-stu-id="51c58-117">Member</span></span> |
| [<span data-ttu-id="51c58-118">displayName</span><span class="sxs-lookup"><span data-stu-id="51c58-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="51c58-119">Member</span><span class="sxs-lookup"><span data-stu-id="51c58-119">Member</span></span> |
| [<span data-ttu-id="51c58-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="51c58-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="51c58-121">Member</span><span class="sxs-lookup"><span data-stu-id="51c58-121">Member</span></span> |
| [<span data-ttu-id="51c58-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="51c58-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="51c58-123">Membro</span><span class="sxs-lookup"><span data-stu-id="51c58-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="51c58-124">Membros</span><span class="sxs-lookup"><span data-stu-id="51c58-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="51c58-125">AccountType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="51c58-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="51c58-126">Atualmente, esse membro só tem suporte no Outlook 2016 ou posterior para Mac (Build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="51c58-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="51c58-127">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="51c58-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="51c58-128">Os valores possíveis estão listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="51c58-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="51c58-129">Valor</span><span class="sxs-lookup"><span data-stu-id="51c58-129">Value</span></span> | <span data-ttu-id="51c58-130">Descrição</span><span class="sxs-lookup"><span data-stu-id="51c58-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="51c58-131">A caixa de correio está em um servidor Exchange local.</span><span class="sxs-lookup"><span data-stu-id="51c58-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="51c58-132">A caixa de correio está associada a uma conta do gmail.</span><span class="sxs-lookup"><span data-stu-id="51c58-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="51c58-133">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="51c58-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="51c58-134">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="51c58-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="51c58-135">Tipo</span><span class="sxs-lookup"><span data-stu-id="51c58-135">Type</span></span>

*   <span data-ttu-id="51c58-136">String</span><span class="sxs-lookup"><span data-stu-id="51c58-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="51c58-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="51c58-137">Requirements</span></span>

|<span data-ttu-id="51c58-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="51c58-138">Requirement</span></span>| <span data-ttu-id="51c58-139">Valor</span><span class="sxs-lookup"><span data-stu-id="51c58-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="51c58-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="51c58-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="51c58-141">1.6</span><span class="sxs-lookup"><span data-stu-id="51c58-141">1.6</span></span> |
|[<span data-ttu-id="51c58-142">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="51c58-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="51c58-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="51c58-143">ReadItem</span></span>|
|[<span data-ttu-id="51c58-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="51c58-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="51c58-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="51c58-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="51c58-146">Exemplo</span><span class="sxs-lookup"><span data-stu-id="51c58-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="51c58-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="51c58-147">displayName :String</span></span>

<span data-ttu-id="51c58-148">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="51c58-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="51c58-149">Tipo</span><span class="sxs-lookup"><span data-stu-id="51c58-149">Type</span></span>

*   <span data-ttu-id="51c58-150">String</span><span class="sxs-lookup"><span data-stu-id="51c58-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="51c58-151">Requisitos</span><span class="sxs-lookup"><span data-stu-id="51c58-151">Requirements</span></span>

|<span data-ttu-id="51c58-152">Requisito</span><span class="sxs-lookup"><span data-stu-id="51c58-152">Requirement</span></span>| <span data-ttu-id="51c58-153">Valor</span><span class="sxs-lookup"><span data-stu-id="51c58-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="51c58-154">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="51c58-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="51c58-155">1.0</span><span class="sxs-lookup"><span data-stu-id="51c58-155">1.0</span></span>|
|[<span data-ttu-id="51c58-156">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="51c58-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="51c58-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="51c58-157">ReadItem</span></span>|
|[<span data-ttu-id="51c58-158">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="51c58-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="51c58-159">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="51c58-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="51c58-160">Exemplo</span><span class="sxs-lookup"><span data-stu-id="51c58-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="51c58-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="51c58-161">emailAddress :String</span></span>

<span data-ttu-id="51c58-162">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="51c58-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="51c58-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="51c58-163">Type</span></span>

*   <span data-ttu-id="51c58-164">String</span><span class="sxs-lookup"><span data-stu-id="51c58-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="51c58-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="51c58-165">Requirements</span></span>

|<span data-ttu-id="51c58-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="51c58-166">Requirement</span></span>| <span data-ttu-id="51c58-167">Valor</span><span class="sxs-lookup"><span data-stu-id="51c58-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="51c58-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="51c58-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="51c58-169">1.0</span><span class="sxs-lookup"><span data-stu-id="51c58-169">1.0</span></span>|
|[<span data-ttu-id="51c58-170">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="51c58-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="51c58-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="51c58-171">ReadItem</span></span>|
|[<span data-ttu-id="51c58-172">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="51c58-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="51c58-173">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="51c58-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="51c58-174">Exemplo</span><span class="sxs-lookup"><span data-stu-id="51c58-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="51c58-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="51c58-175">timeZone :String</span></span>

<span data-ttu-id="51c58-176">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="51c58-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="51c58-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="51c58-177">Type</span></span>

*   <span data-ttu-id="51c58-178">String</span><span class="sxs-lookup"><span data-stu-id="51c58-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="51c58-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="51c58-179">Requirements</span></span>

|<span data-ttu-id="51c58-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="51c58-180">Requirement</span></span>| <span data-ttu-id="51c58-181">Valor</span><span class="sxs-lookup"><span data-stu-id="51c58-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="51c58-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="51c58-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="51c58-183">1.0</span><span class="sxs-lookup"><span data-stu-id="51c58-183">1.0</span></span>|
|[<span data-ttu-id="51c58-184">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="51c58-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="51c58-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="51c58-185">ReadItem</span></span>|
|[<span data-ttu-id="51c58-186">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="51c58-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="51c58-187">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="51c58-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="51c58-188">Exemplo</span><span class="sxs-lookup"><span data-stu-id="51c58-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
