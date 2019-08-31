---
title: Namespace do Office – conjunto de requisitos 1.2
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 25322842d82c1db7228a67c268facb6945168222
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696292"
---
# <a name="office"></a><span data-ttu-id="e7a92-102">Office</span><span class="sxs-lookup"><span data-stu-id="e7a92-102">Office</span></span>

<span data-ttu-id="e7a92-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="e7a92-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7a92-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e7a92-105">Requirements</span></span>

|<span data-ttu-id="e7a92-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="e7a92-106">Requirement</span></span>| <span data-ttu-id="e7a92-107">Valor</span><span class="sxs-lookup"><span data-stu-id="e7a92-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7a92-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e7a92-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7a92-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e7a92-109">1.0</span></span>|
|[<span data-ttu-id="e7a92-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e7a92-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7a92-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e7a92-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e7a92-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="e7a92-112">Members and methods</span></span>

| <span data-ttu-id="e7a92-113">Membro</span><span class="sxs-lookup"><span data-stu-id="e7a92-113">Member</span></span> | <span data-ttu-id="e7a92-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="e7a92-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e7a92-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="e7a92-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="e7a92-116">Membro</span><span class="sxs-lookup"><span data-stu-id="e7a92-116">Member</span></span> |
| [<span data-ttu-id="e7a92-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="e7a92-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="e7a92-118">Membro</span><span class="sxs-lookup"><span data-stu-id="e7a92-118">Member</span></span> |
| [<span data-ttu-id="e7a92-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="e7a92-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="e7a92-120">Membro</span><span class="sxs-lookup"><span data-stu-id="e7a92-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="e7a92-121">Namespaces</span><span class="sxs-lookup"><span data-stu-id="e7a92-121">Namespaces</span></span>

<span data-ttu-id="e7a92-122">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="e7a92-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="e7a92-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): inclui um número de enumerações, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="e7a92-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="e7a92-124">Members</span><span class="sxs-lookup"><span data-stu-id="e7a92-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="e7a92-125">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e7a92-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="e7a92-126">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="e7a92-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e7a92-127">Tipo</span><span class="sxs-lookup"><span data-stu-id="e7a92-127">Type</span></span>

*   <span data-ttu-id="e7a92-128">String</span><span class="sxs-lookup"><span data-stu-id="e7a92-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e7a92-129">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="e7a92-129">Properties:</span></span>

|<span data-ttu-id="e7a92-130">Nome</span><span class="sxs-lookup"><span data-stu-id="e7a92-130">Name</span></span>| <span data-ttu-id="e7a92-131">Tipo</span><span class="sxs-lookup"><span data-stu-id="e7a92-131">Type</span></span>| <span data-ttu-id="e7a92-132">Descrição</span><span class="sxs-lookup"><span data-stu-id="e7a92-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e7a92-133">String</span><span class="sxs-lookup"><span data-stu-id="e7a92-133">String</span></span>|<span data-ttu-id="e7a92-134">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="e7a92-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e7a92-135">String</span><span class="sxs-lookup"><span data-stu-id="e7a92-135">String</span></span>|<span data-ttu-id="e7a92-136">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="e7a92-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7a92-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e7a92-137">Requirements</span></span>

|<span data-ttu-id="e7a92-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="e7a92-138">Requirement</span></span>| <span data-ttu-id="e7a92-139">Valor</span><span class="sxs-lookup"><span data-stu-id="e7a92-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7a92-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e7a92-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7a92-141">1.0</span><span class="sxs-lookup"><span data-stu-id="e7a92-141">1.0</span></span>|
|[<span data-ttu-id="e7a92-142">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e7a92-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7a92-143">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e7a92-143">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="e7a92-144">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e7a92-144">CoercionType: String</span></span>

<span data-ttu-id="e7a92-145">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="e7a92-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e7a92-146">Tipo</span><span class="sxs-lookup"><span data-stu-id="e7a92-146">Type</span></span>

*   <span data-ttu-id="e7a92-147">String</span><span class="sxs-lookup"><span data-stu-id="e7a92-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e7a92-148">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="e7a92-148">Properties:</span></span>

|<span data-ttu-id="e7a92-149">Nome</span><span class="sxs-lookup"><span data-stu-id="e7a92-149">Name</span></span>| <span data-ttu-id="e7a92-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="e7a92-150">Type</span></span>| <span data-ttu-id="e7a92-151">Descrição</span><span class="sxs-lookup"><span data-stu-id="e7a92-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e7a92-152">String</span><span class="sxs-lookup"><span data-stu-id="e7a92-152">String</span></span>|<span data-ttu-id="e7a92-153">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="e7a92-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e7a92-154">String</span><span class="sxs-lookup"><span data-stu-id="e7a92-154">String</span></span>|<span data-ttu-id="e7a92-155">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="e7a92-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7a92-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e7a92-156">Requirements</span></span>

|<span data-ttu-id="e7a92-157">Requisito</span><span class="sxs-lookup"><span data-stu-id="e7a92-157">Requirement</span></span>| <span data-ttu-id="e7a92-158">Valor</span><span class="sxs-lookup"><span data-stu-id="e7a92-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7a92-159">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e7a92-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7a92-160">1.0</span><span class="sxs-lookup"><span data-stu-id="e7a92-160">1.0</span></span>|
|[<span data-ttu-id="e7a92-161">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e7a92-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7a92-162">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e7a92-162">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="e7a92-163">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e7a92-163">SourceProperty: String</span></span>

<span data-ttu-id="e7a92-164">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="e7a92-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e7a92-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="e7a92-165">Type</span></span>

*   <span data-ttu-id="e7a92-166">String</span><span class="sxs-lookup"><span data-stu-id="e7a92-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e7a92-167">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="e7a92-167">Properties:</span></span>

|<span data-ttu-id="e7a92-168">Nome</span><span class="sxs-lookup"><span data-stu-id="e7a92-168">Name</span></span>| <span data-ttu-id="e7a92-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="e7a92-169">Type</span></span>| <span data-ttu-id="e7a92-170">Descrição</span><span class="sxs-lookup"><span data-stu-id="e7a92-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e7a92-171">String</span><span class="sxs-lookup"><span data-stu-id="e7a92-171">String</span></span>|<span data-ttu-id="e7a92-172">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e7a92-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e7a92-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="e7a92-173">String</span></span>|<span data-ttu-id="e7a92-174">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="e7a92-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7a92-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e7a92-175">Requirements</span></span>

|<span data-ttu-id="e7a92-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="e7a92-176">Requirement</span></span>| <span data-ttu-id="e7a92-177">Valor</span><span class="sxs-lookup"><span data-stu-id="e7a92-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7a92-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="e7a92-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7a92-179">1.0</span><span class="sxs-lookup"><span data-stu-id="e7a92-179">1.0</span></span>|
|[<span data-ttu-id="e7a92-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="e7a92-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e7a92-181">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="e7a92-181">Compose or Read</span></span>|
