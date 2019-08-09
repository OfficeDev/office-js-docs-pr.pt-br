---
title: Namespace do Office – conjunto de requisitos 1,3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 0b22574693fb129be6a08a89b58beceb746fa283
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268394"
---
# <a name="office"></a><span data-ttu-id="4eb36-102">Office</span><span class="sxs-lookup"><span data-stu-id="4eb36-102">Office</span></span>

<span data-ttu-id="4eb36-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="4eb36-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4eb36-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4eb36-105">Requirements</span></span>

|<span data-ttu-id="4eb36-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="4eb36-106">Requirement</span></span>| <span data-ttu-id="4eb36-107">Valor</span><span class="sxs-lookup"><span data-stu-id="4eb36-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4eb36-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4eb36-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4eb36-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4eb36-109">1.0</span></span>|
|[<span data-ttu-id="4eb36-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4eb36-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4eb36-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4eb36-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4eb36-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="4eb36-112">Members and methods</span></span>

| <span data-ttu-id="4eb36-113">Membro</span><span class="sxs-lookup"><span data-stu-id="4eb36-113">Member</span></span> | <span data-ttu-id="4eb36-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="4eb36-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4eb36-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4eb36-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4eb36-116">Membro</span><span class="sxs-lookup"><span data-stu-id="4eb36-116">Member</span></span> |
| [<span data-ttu-id="4eb36-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4eb36-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4eb36-118">Membro</span><span class="sxs-lookup"><span data-stu-id="4eb36-118">Member</span></span> |
| [<span data-ttu-id="4eb36-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4eb36-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4eb36-120">Membro</span><span class="sxs-lookup"><span data-stu-id="4eb36-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4eb36-121">Namespaces</span><span class="sxs-lookup"><span data-stu-id="4eb36-121">Namespaces</span></span>

<span data-ttu-id="4eb36-122">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="4eb36-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="4eb36-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.3): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="4eb36-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.3): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="4eb36-124">Membros</span><span class="sxs-lookup"><span data-stu-id="4eb36-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="4eb36-125">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4eb36-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="4eb36-126">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="4eb36-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4eb36-127">Tipo</span><span class="sxs-lookup"><span data-stu-id="4eb36-127">Type</span></span>

*   <span data-ttu-id="4eb36-128">String</span><span class="sxs-lookup"><span data-stu-id="4eb36-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4eb36-129">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="4eb36-129">Properties:</span></span>

|<span data-ttu-id="4eb36-130">Nome</span><span class="sxs-lookup"><span data-stu-id="4eb36-130">Name</span></span>| <span data-ttu-id="4eb36-131">Tipo</span><span class="sxs-lookup"><span data-stu-id="4eb36-131">Type</span></span>| <span data-ttu-id="4eb36-132">Descrição</span><span class="sxs-lookup"><span data-stu-id="4eb36-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4eb36-133">String</span><span class="sxs-lookup"><span data-stu-id="4eb36-133">String</span></span>|<span data-ttu-id="4eb36-134">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="4eb36-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4eb36-135">String</span><span class="sxs-lookup"><span data-stu-id="4eb36-135">String</span></span>|<span data-ttu-id="4eb36-136">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="4eb36-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4eb36-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4eb36-137">Requirements</span></span>

|<span data-ttu-id="4eb36-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="4eb36-138">Requirement</span></span>| <span data-ttu-id="4eb36-139">Valor</span><span class="sxs-lookup"><span data-stu-id="4eb36-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="4eb36-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4eb36-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4eb36-141">1.0</span><span class="sxs-lookup"><span data-stu-id="4eb36-141">1.0</span></span>|
|[<span data-ttu-id="4eb36-142">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4eb36-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4eb36-143">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4eb36-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="4eb36-144">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4eb36-144">CoercionType: String</span></span>

<span data-ttu-id="4eb36-145">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="4eb36-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4eb36-146">Tipo</span><span class="sxs-lookup"><span data-stu-id="4eb36-146">Type</span></span>

*   <span data-ttu-id="4eb36-147">String</span><span class="sxs-lookup"><span data-stu-id="4eb36-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4eb36-148">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="4eb36-148">Properties:</span></span>

|<span data-ttu-id="4eb36-149">Nome</span><span class="sxs-lookup"><span data-stu-id="4eb36-149">Name</span></span>| <span data-ttu-id="4eb36-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="4eb36-150">Type</span></span>| <span data-ttu-id="4eb36-151">Descrição</span><span class="sxs-lookup"><span data-stu-id="4eb36-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4eb36-152">String</span><span class="sxs-lookup"><span data-stu-id="4eb36-152">String</span></span>|<span data-ttu-id="4eb36-153">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="4eb36-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4eb36-154">String</span><span class="sxs-lookup"><span data-stu-id="4eb36-154">String</span></span>|<span data-ttu-id="4eb36-155">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="4eb36-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4eb36-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4eb36-156">Requirements</span></span>

|<span data-ttu-id="4eb36-157">Requisito</span><span class="sxs-lookup"><span data-stu-id="4eb36-157">Requirement</span></span>| <span data-ttu-id="4eb36-158">Valor</span><span class="sxs-lookup"><span data-stu-id="4eb36-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="4eb36-159">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4eb36-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4eb36-160">1.0</span><span class="sxs-lookup"><span data-stu-id="4eb36-160">1.0</span></span>|
|[<span data-ttu-id="4eb36-161">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4eb36-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4eb36-162">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4eb36-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="4eb36-163">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4eb36-163">SourceProperty: String</span></span>

<span data-ttu-id="4eb36-164">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="4eb36-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4eb36-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="4eb36-165">Type</span></span>

*   <span data-ttu-id="4eb36-166">String</span><span class="sxs-lookup"><span data-stu-id="4eb36-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4eb36-167">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="4eb36-167">Properties:</span></span>

|<span data-ttu-id="4eb36-168">Nome</span><span class="sxs-lookup"><span data-stu-id="4eb36-168">Name</span></span>| <span data-ttu-id="4eb36-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="4eb36-169">Type</span></span>| <span data-ttu-id="4eb36-170">Descrição</span><span class="sxs-lookup"><span data-stu-id="4eb36-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4eb36-171">String</span><span class="sxs-lookup"><span data-stu-id="4eb36-171">String</span></span>|<span data-ttu-id="4eb36-172">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4eb36-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4eb36-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4eb36-173">String</span></span>|<span data-ttu-id="4eb36-174">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4eb36-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4eb36-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4eb36-175">Requirements</span></span>

|<span data-ttu-id="4eb36-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="4eb36-176">Requirement</span></span>| <span data-ttu-id="4eb36-177">Valor</span><span class="sxs-lookup"><span data-stu-id="4eb36-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="4eb36-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4eb36-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4eb36-179">1.0</span><span class="sxs-lookup"><span data-stu-id="4eb36-179">1.0</span></span>|
|[<span data-ttu-id="4eb36-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4eb36-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4eb36-181">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4eb36-181">Compose or Read</span></span>|
