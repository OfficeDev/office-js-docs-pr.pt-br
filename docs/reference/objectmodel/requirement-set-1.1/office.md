---
title: Namespace do Office – conjunto de requisitos 1,1
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 70413bdfc01378bb5b1814fd938ab94a7e5101ba
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696390"
---
# <a name="office"></a><span data-ttu-id="90976-102">Office</span><span class="sxs-lookup"><span data-stu-id="90976-102">Office</span></span>

<span data-ttu-id="90976-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="90976-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="90976-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="90976-105">Requirements</span></span>

|<span data-ttu-id="90976-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="90976-106">Requirement</span></span>| <span data-ttu-id="90976-107">Valor</span><span class="sxs-lookup"><span data-stu-id="90976-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="90976-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="90976-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="90976-109">1.0</span><span class="sxs-lookup"><span data-stu-id="90976-109">1.0</span></span>|
|[<span data-ttu-id="90976-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="90976-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="90976-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="90976-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="90976-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="90976-112">Members and methods</span></span>

| <span data-ttu-id="90976-113">Membro</span><span class="sxs-lookup"><span data-stu-id="90976-113">Member</span></span> | <span data-ttu-id="90976-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="90976-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="90976-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="90976-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="90976-116">Membro</span><span class="sxs-lookup"><span data-stu-id="90976-116">Member</span></span> |
| [<span data-ttu-id="90976-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="90976-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="90976-118">Membro</span><span class="sxs-lookup"><span data-stu-id="90976-118">Member</span></span> |
| [<span data-ttu-id="90976-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="90976-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="90976-120">Membro</span><span class="sxs-lookup"><span data-stu-id="90976-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="90976-121">Namespaces</span><span class="sxs-lookup"><span data-stu-id="90976-121">Namespaces</span></span>

<span data-ttu-id="90976-122">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="90976-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="90976-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.1): inclui um número de enumerações, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="90976-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.1): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="90976-124">Members</span><span class="sxs-lookup"><span data-stu-id="90976-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="90976-125">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="90976-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="90976-126">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="90976-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="90976-127">Tipo</span><span class="sxs-lookup"><span data-stu-id="90976-127">Type</span></span>

*   <span data-ttu-id="90976-128">String</span><span class="sxs-lookup"><span data-stu-id="90976-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="90976-129">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="90976-129">Properties:</span></span>

|<span data-ttu-id="90976-130">Nome</span><span class="sxs-lookup"><span data-stu-id="90976-130">Name</span></span>| <span data-ttu-id="90976-131">Tipo</span><span class="sxs-lookup"><span data-stu-id="90976-131">Type</span></span>| <span data-ttu-id="90976-132">Descrição</span><span class="sxs-lookup"><span data-stu-id="90976-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="90976-133">String</span><span class="sxs-lookup"><span data-stu-id="90976-133">String</span></span>|<span data-ttu-id="90976-134">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="90976-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="90976-135">String</span><span class="sxs-lookup"><span data-stu-id="90976-135">String</span></span>|<span data-ttu-id="90976-136">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="90976-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="90976-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="90976-137">Requirements</span></span>

|<span data-ttu-id="90976-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="90976-138">Requirement</span></span>| <span data-ttu-id="90976-139">Valor</span><span class="sxs-lookup"><span data-stu-id="90976-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="90976-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="90976-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="90976-141">1.0</span><span class="sxs-lookup"><span data-stu-id="90976-141">1.0</span></span>|
|[<span data-ttu-id="90976-142">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="90976-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="90976-143">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="90976-143">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="90976-144">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="90976-144">CoercionType: String</span></span>

<span data-ttu-id="90976-145">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="90976-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="90976-146">Tipo</span><span class="sxs-lookup"><span data-stu-id="90976-146">Type</span></span>

*   <span data-ttu-id="90976-147">String</span><span class="sxs-lookup"><span data-stu-id="90976-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="90976-148">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="90976-148">Properties:</span></span>

|<span data-ttu-id="90976-149">Nome</span><span class="sxs-lookup"><span data-stu-id="90976-149">Name</span></span>| <span data-ttu-id="90976-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="90976-150">Type</span></span>| <span data-ttu-id="90976-151">Descrição</span><span class="sxs-lookup"><span data-stu-id="90976-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="90976-152">String</span><span class="sxs-lookup"><span data-stu-id="90976-152">String</span></span>|<span data-ttu-id="90976-153">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="90976-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="90976-154">String</span><span class="sxs-lookup"><span data-stu-id="90976-154">String</span></span>|<span data-ttu-id="90976-155">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="90976-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="90976-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="90976-156">Requirements</span></span>

|<span data-ttu-id="90976-157">Requisito</span><span class="sxs-lookup"><span data-stu-id="90976-157">Requirement</span></span>| <span data-ttu-id="90976-158">Valor</span><span class="sxs-lookup"><span data-stu-id="90976-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="90976-159">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="90976-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="90976-160">1.0</span><span class="sxs-lookup"><span data-stu-id="90976-160">1.0</span></span>|
|[<span data-ttu-id="90976-161">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="90976-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="90976-162">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="90976-162">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="90976-163">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="90976-163">SourceProperty: String</span></span>

<span data-ttu-id="90976-164">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="90976-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="90976-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="90976-165">Type</span></span>

*   <span data-ttu-id="90976-166">String</span><span class="sxs-lookup"><span data-stu-id="90976-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="90976-167">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="90976-167">Properties:</span></span>

|<span data-ttu-id="90976-168">Nome</span><span class="sxs-lookup"><span data-stu-id="90976-168">Name</span></span>| <span data-ttu-id="90976-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="90976-169">Type</span></span>| <span data-ttu-id="90976-170">Descrição</span><span class="sxs-lookup"><span data-stu-id="90976-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="90976-171">String</span><span class="sxs-lookup"><span data-stu-id="90976-171">String</span></span>|<span data-ttu-id="90976-172">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="90976-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="90976-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="90976-173">String</span></span>|<span data-ttu-id="90976-174">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="90976-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="90976-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="90976-175">Requirements</span></span>

|<span data-ttu-id="90976-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="90976-176">Requirement</span></span>| <span data-ttu-id="90976-177">Valor</span><span class="sxs-lookup"><span data-stu-id="90976-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="90976-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="90976-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="90976-179">1.0</span><span class="sxs-lookup"><span data-stu-id="90976-179">1.0</span></span>|
|[<span data-ttu-id="90976-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="90976-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="90976-181">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="90976-181">Compose or Read</span></span>|
