---
title: Office. Context. Mailbox. userProfile – conjunto de requisitos de visualização
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 4afc64f247155576ab3f0024d1929a29a0f7dc0c
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629255"
---
# <a name="userprofile"></a><span data-ttu-id="bebf8-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="bebf8-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="bebf8-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="bebf8-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="bebf8-104">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bebf8-104">Requirements</span></span>

|<span data-ttu-id="bebf8-105">Requisito</span><span class="sxs-lookup"><span data-stu-id="bebf8-105">Requirement</span></span>| <span data-ttu-id="bebf8-106">Valor</span><span class="sxs-lookup"><span data-stu-id="bebf8-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="bebf8-107">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bebf8-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bebf8-108">1.0</span><span class="sxs-lookup"><span data-stu-id="bebf8-108">1.0</span></span>|
|[<span data-ttu-id="bebf8-109">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bebf8-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bebf8-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bebf8-110">ReadItem</span></span>|
|[<span data-ttu-id="bebf8-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bebf8-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bebf8-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bebf8-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="bebf8-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="bebf8-113">Properties</span></span>

| <span data-ttu-id="bebf8-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="bebf8-114">Property</span></span> | <span data-ttu-id="bebf8-115">Mínimo</span><span class="sxs-lookup"><span data-stu-id="bebf8-115">Minimum</span></span><br><span data-ttu-id="bebf8-116">nível de permissão</span><span class="sxs-lookup"><span data-stu-id="bebf8-116">permission level</span></span> | <span data-ttu-id="bebf8-117">Modelos</span><span class="sxs-lookup"><span data-stu-id="bebf8-117">Modes</span></span> | <span data-ttu-id="bebf8-118">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="bebf8-118">Return type</span></span> | <span data-ttu-id="bebf8-119">Mínimo</span><span class="sxs-lookup"><span data-stu-id="bebf8-119">Minimum</span></span><br><span data-ttu-id="bebf8-120">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="bebf8-120">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="bebf8-121">accountType</span><span class="sxs-lookup"><span data-stu-id="bebf8-121">accountType</span></span>](#accounttype-string) | <span data-ttu-id="bebf8-122">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bebf8-122">ReadItem</span></span> | <span data-ttu-id="bebf8-123">Escrever</span><span class="sxs-lookup"><span data-stu-id="bebf8-123">Compose</span></span><br><span data-ttu-id="bebf8-124">Ler</span><span class="sxs-lookup"><span data-stu-id="bebf8-124">Read</span></span> | <span data-ttu-id="bebf8-125">String</span><span class="sxs-lookup"><span data-stu-id="bebf8-125">String</span></span> | <span data-ttu-id="bebf8-126">1.6</span><span class="sxs-lookup"><span data-stu-id="bebf8-126">1.6</span></span> |
| [<span data-ttu-id="bebf8-127">displayName</span><span class="sxs-lookup"><span data-stu-id="bebf8-127">displayName</span></span>](#displayname-string) | <span data-ttu-id="bebf8-128">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bebf8-128">ReadItem</span></span> | <span data-ttu-id="bebf8-129">Escrever</span><span class="sxs-lookup"><span data-stu-id="bebf8-129">Compose</span></span><br><span data-ttu-id="bebf8-130">Ler</span><span class="sxs-lookup"><span data-stu-id="bebf8-130">Read</span></span> | <span data-ttu-id="bebf8-131">String</span><span class="sxs-lookup"><span data-stu-id="bebf8-131">String</span></span> | <span data-ttu-id="bebf8-132">1.0</span><span class="sxs-lookup"><span data-stu-id="bebf8-132">1.0</span></span> |
| [<span data-ttu-id="bebf8-133">emailAddress</span><span class="sxs-lookup"><span data-stu-id="bebf8-133">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="bebf8-134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bebf8-134">ReadItem</span></span> | <span data-ttu-id="bebf8-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="bebf8-135">Compose</span></span><br><span data-ttu-id="bebf8-136">Ler</span><span class="sxs-lookup"><span data-stu-id="bebf8-136">Read</span></span> | <span data-ttu-id="bebf8-137">String</span><span class="sxs-lookup"><span data-stu-id="bebf8-137">String</span></span> | <span data-ttu-id="bebf8-138">1.0</span><span class="sxs-lookup"><span data-stu-id="bebf8-138">1.0</span></span> |
| [<span data-ttu-id="bebf8-139">timeZone</span><span class="sxs-lookup"><span data-stu-id="bebf8-139">timeZone</span></span>](#timezone-string) | <span data-ttu-id="bebf8-140">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bebf8-140">ReadItem</span></span> | <span data-ttu-id="bebf8-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="bebf8-141">Compose</span></span><br><span data-ttu-id="bebf8-142">Ler</span><span class="sxs-lookup"><span data-stu-id="bebf8-142">Read</span></span> | <span data-ttu-id="bebf8-143">String</span><span class="sxs-lookup"><span data-stu-id="bebf8-143">String</span></span> | <span data-ttu-id="bebf8-144">1.0</span><span class="sxs-lookup"><span data-stu-id="bebf8-144">1.0</span></span> |

## <a name="property-details"></a><span data-ttu-id="bebf8-145">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="bebf8-145">Property details</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="bebf8-146">AccountType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bebf8-146">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="bebf8-147">Atualmente, esse membro só tem suporte no Outlook 2016 ou posterior no Mac (Build 16.9.1212 ou posterior).</span><span class="sxs-lookup"><span data-stu-id="bebf8-147">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="bebf8-148">Obtém o tipo de conta do usuário associado à caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="bebf8-148">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="bebf8-149">Os valores possíveis estão listados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="bebf8-149">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="bebf8-150">Valor</span><span class="sxs-lookup"><span data-stu-id="bebf8-150">Value</span></span> | <span data-ttu-id="bebf8-151">Descrição</span><span class="sxs-lookup"><span data-stu-id="bebf8-151">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="bebf8-152">A caixa de correio está em um servidor Exchange local.</span><span class="sxs-lookup"><span data-stu-id="bebf8-152">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="bebf8-153">A caixa de correio está associada a uma conta do gmail.</span><span class="sxs-lookup"><span data-stu-id="bebf8-153">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="bebf8-154">A caixa de correio está associada a uma conta corporativa ou de estudante do Office 365.</span><span class="sxs-lookup"><span data-stu-id="bebf8-154">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="bebf8-155">A caixa de correio está associada a uma conta pessoal do Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="bebf8-155">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="bebf8-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="bebf8-156">Type</span></span>

*   <span data-ttu-id="bebf8-157">String</span><span class="sxs-lookup"><span data-stu-id="bebf8-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bebf8-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bebf8-158">Requirements</span></span>

|<span data-ttu-id="bebf8-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="bebf8-159">Requirement</span></span>| <span data-ttu-id="bebf8-160">Valor</span><span class="sxs-lookup"><span data-stu-id="bebf8-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="bebf8-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bebf8-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bebf8-162">1.6</span><span class="sxs-lookup"><span data-stu-id="bebf8-162">1.6</span></span> |
|[<span data-ttu-id="bebf8-163">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bebf8-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bebf8-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bebf8-164">ReadItem</span></span>|
|[<span data-ttu-id="bebf8-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bebf8-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bebf8-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bebf8-166">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bebf8-167">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bebf8-167">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="bebf8-168">displayName: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bebf8-168">displayName: String</span></span>

<span data-ttu-id="bebf8-169">Obtém o nome de exibição do usuário.</span><span class="sxs-lookup"><span data-stu-id="bebf8-169">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="bebf8-170">Tipo</span><span class="sxs-lookup"><span data-stu-id="bebf8-170">Type</span></span>

*   <span data-ttu-id="bebf8-171">String</span><span class="sxs-lookup"><span data-stu-id="bebf8-171">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bebf8-172">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bebf8-172">Requirements</span></span>

|<span data-ttu-id="bebf8-173">Requisito</span><span class="sxs-lookup"><span data-stu-id="bebf8-173">Requirement</span></span>| <span data-ttu-id="bebf8-174">Valor</span><span class="sxs-lookup"><span data-stu-id="bebf8-174">Value</span></span>|
|---|---|
|[<span data-ttu-id="bebf8-175">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bebf8-175">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bebf8-176">1.0</span><span class="sxs-lookup"><span data-stu-id="bebf8-176">1.0</span></span>|
|[<span data-ttu-id="bebf8-177">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bebf8-177">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bebf8-178">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bebf8-178">ReadItem</span></span>|
|[<span data-ttu-id="bebf8-179">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bebf8-179">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bebf8-180">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bebf8-180">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bebf8-181">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bebf8-181">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="bebf8-182">emailAddress: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bebf8-182">emailAddress: String</span></span>

<span data-ttu-id="bebf8-183">Obtém o endereço de email SMTP do usuário.</span><span class="sxs-lookup"><span data-stu-id="bebf8-183">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="bebf8-184">Tipo</span><span class="sxs-lookup"><span data-stu-id="bebf8-184">Type</span></span>

*   <span data-ttu-id="bebf8-185">String</span><span class="sxs-lookup"><span data-stu-id="bebf8-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bebf8-186">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bebf8-186">Requirements</span></span>

|<span data-ttu-id="bebf8-187">Requisito</span><span class="sxs-lookup"><span data-stu-id="bebf8-187">Requirement</span></span>| <span data-ttu-id="bebf8-188">Valor</span><span class="sxs-lookup"><span data-stu-id="bebf8-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="bebf8-189">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bebf8-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bebf8-190">1.0</span><span class="sxs-lookup"><span data-stu-id="bebf8-190">1.0</span></span>|
|[<span data-ttu-id="bebf8-191">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bebf8-191">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bebf8-192">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bebf8-192">ReadItem</span></span>|
|[<span data-ttu-id="bebf8-193">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bebf8-193">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bebf8-194">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bebf8-194">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bebf8-195">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bebf8-195">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="bebf8-196">timeZone: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="bebf8-196">timeZone: String</span></span>

<span data-ttu-id="bebf8-197">Obtém o fuso horário padrão do usuário.</span><span class="sxs-lookup"><span data-stu-id="bebf8-197">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="bebf8-198">Tipo</span><span class="sxs-lookup"><span data-stu-id="bebf8-198">Type</span></span>

*   <span data-ttu-id="bebf8-199">String</span><span class="sxs-lookup"><span data-stu-id="bebf8-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bebf8-200">Requisitos</span><span class="sxs-lookup"><span data-stu-id="bebf8-200">Requirements</span></span>

|<span data-ttu-id="bebf8-201">Requisito</span><span class="sxs-lookup"><span data-stu-id="bebf8-201">Requirement</span></span>| <span data-ttu-id="bebf8-202">Valor</span><span class="sxs-lookup"><span data-stu-id="bebf8-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="bebf8-203">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="bebf8-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bebf8-204">1.0</span><span class="sxs-lookup"><span data-stu-id="bebf8-204">1.0</span></span>|
|[<span data-ttu-id="bebf8-205">Nível de permissão mínimo</span><span class="sxs-lookup"><span data-stu-id="bebf8-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bebf8-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bebf8-206">ReadItem</span></span>|
|[<span data-ttu-id="bebf8-207">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="bebf8-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bebf8-208">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="bebf8-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bebf8-209">Exemplo</span><span class="sxs-lookup"><span data-stu-id="bebf8-209">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
