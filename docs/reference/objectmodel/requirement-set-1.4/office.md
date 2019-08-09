---
title: Namespace do Office – conjunto de requisitos 1,4
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: e6c4614af74a665805c400c407e4a7785efe9f96
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268646"
---
# <a name="office"></a><span data-ttu-id="ec1d2-102">Office</span><span class="sxs-lookup"><span data-stu-id="ec1d2-102">Office</span></span>

<span data-ttu-id="ec1d2-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ec1d2-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ec1d2-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ec1d2-105">Requirements</span></span>

|<span data-ttu-id="ec1d2-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="ec1d2-106">Requirement</span></span>| <span data-ttu-id="ec1d2-107">Valor</span><span class="sxs-lookup"><span data-stu-id="ec1d2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec1d2-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ec1d2-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec1d2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ec1d2-109">1.0</span></span>|
|[<span data-ttu-id="ec1d2-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ec1d2-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ec1d2-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ec1d2-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ec1d2-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="ec1d2-112">Members and methods</span></span>

| <span data-ttu-id="ec1d2-113">Membro</span><span class="sxs-lookup"><span data-stu-id="ec1d2-113">Member</span></span> | <span data-ttu-id="ec1d2-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="ec1d2-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ec1d2-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ec1d2-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ec1d2-116">Membro</span><span class="sxs-lookup"><span data-stu-id="ec1d2-116">Member</span></span> |
| [<span data-ttu-id="ec1d2-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ec1d2-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ec1d2-118">Membro</span><span class="sxs-lookup"><span data-stu-id="ec1d2-118">Member</span></span> |
| [<span data-ttu-id="ec1d2-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ec1d2-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ec1d2-120">Membro</span><span class="sxs-lookup"><span data-stu-id="ec1d2-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ec1d2-121">Namespaces</span><span class="sxs-lookup"><span data-stu-id="ec1d2-121">Namespaces</span></span>

<span data-ttu-id="ec1d2-122">[context](Office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="ec1d2-122">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="ec1d2-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="ec1d2-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="ec1d2-124">Membros</span><span class="sxs-lookup"><span data-stu-id="ec1d2-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="ec1d2-125">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec1d2-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="ec1d2-126">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="ec1d2-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ec1d2-127">Tipo</span><span class="sxs-lookup"><span data-stu-id="ec1d2-127">Type</span></span>

*   <span data-ttu-id="ec1d2-128">String</span><span class="sxs-lookup"><span data-stu-id="ec1d2-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ec1d2-129">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ec1d2-129">Properties:</span></span>

|<span data-ttu-id="ec1d2-130">Nome</span><span class="sxs-lookup"><span data-stu-id="ec1d2-130">Name</span></span>| <span data-ttu-id="ec1d2-131">Tipo</span><span class="sxs-lookup"><span data-stu-id="ec1d2-131">Type</span></span>| <span data-ttu-id="ec1d2-132">Descrição</span><span class="sxs-lookup"><span data-stu-id="ec1d2-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ec1d2-133">String</span><span class="sxs-lookup"><span data-stu-id="ec1d2-133">String</span></span>|<span data-ttu-id="ec1d2-134">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="ec1d2-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ec1d2-135">String</span><span class="sxs-lookup"><span data-stu-id="ec1d2-135">String</span></span>|<span data-ttu-id="ec1d2-136">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="ec1d2-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ec1d2-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ec1d2-137">Requirements</span></span>

|<span data-ttu-id="ec1d2-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="ec1d2-138">Requirement</span></span>| <span data-ttu-id="ec1d2-139">Valor</span><span class="sxs-lookup"><span data-stu-id="ec1d2-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec1d2-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ec1d2-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec1d2-141">1.0</span><span class="sxs-lookup"><span data-stu-id="ec1d2-141">1.0</span></span>|
|[<span data-ttu-id="ec1d2-142">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ec1d2-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ec1d2-143">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ec1d2-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="ec1d2-144">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec1d2-144">CoercionType: String</span></span>

<span data-ttu-id="ec1d2-145">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="ec1d2-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ec1d2-146">Tipo</span><span class="sxs-lookup"><span data-stu-id="ec1d2-146">Type</span></span>

*   <span data-ttu-id="ec1d2-147">String</span><span class="sxs-lookup"><span data-stu-id="ec1d2-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ec1d2-148">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ec1d2-148">Properties:</span></span>

|<span data-ttu-id="ec1d2-149">Nome</span><span class="sxs-lookup"><span data-stu-id="ec1d2-149">Name</span></span>| <span data-ttu-id="ec1d2-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="ec1d2-150">Type</span></span>| <span data-ttu-id="ec1d2-151">Descrição</span><span class="sxs-lookup"><span data-stu-id="ec1d2-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ec1d2-152">String</span><span class="sxs-lookup"><span data-stu-id="ec1d2-152">String</span></span>|<span data-ttu-id="ec1d2-153">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="ec1d2-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ec1d2-154">String</span><span class="sxs-lookup"><span data-stu-id="ec1d2-154">String</span></span>|<span data-ttu-id="ec1d2-155">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="ec1d2-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ec1d2-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ec1d2-156">Requirements</span></span>

|<span data-ttu-id="ec1d2-157">Requisito</span><span class="sxs-lookup"><span data-stu-id="ec1d2-157">Requirement</span></span>| <span data-ttu-id="ec1d2-158">Valor</span><span class="sxs-lookup"><span data-stu-id="ec1d2-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec1d2-159">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ec1d2-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec1d2-160">1.0</span><span class="sxs-lookup"><span data-stu-id="ec1d2-160">1.0</span></span>|
|[<span data-ttu-id="ec1d2-161">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ec1d2-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ec1d2-162">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ec1d2-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="ec1d2-163">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec1d2-163">SourceProperty: String</span></span>

<span data-ttu-id="ec1d2-164">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="ec1d2-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ec1d2-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="ec1d2-165">Type</span></span>

*   <span data-ttu-id="ec1d2-166">String</span><span class="sxs-lookup"><span data-stu-id="ec1d2-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ec1d2-167">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ec1d2-167">Properties:</span></span>

|<span data-ttu-id="ec1d2-168">Nome</span><span class="sxs-lookup"><span data-stu-id="ec1d2-168">Name</span></span>| <span data-ttu-id="ec1d2-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="ec1d2-169">Type</span></span>| <span data-ttu-id="ec1d2-170">Descrição</span><span class="sxs-lookup"><span data-stu-id="ec1d2-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ec1d2-171">String</span><span class="sxs-lookup"><span data-stu-id="ec1d2-171">String</span></span>|<span data-ttu-id="ec1d2-172">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ec1d2-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ec1d2-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ec1d2-173">String</span></span>|<span data-ttu-id="ec1d2-174">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ec1d2-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ec1d2-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ec1d2-175">Requirements</span></span>

|<span data-ttu-id="ec1d2-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="ec1d2-176">Requirement</span></span>| <span data-ttu-id="ec1d2-177">Valor</span><span class="sxs-lookup"><span data-stu-id="ec1d2-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="ec1d2-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ec1d2-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ec1d2-179">1.0</span><span class="sxs-lookup"><span data-stu-id="ec1d2-179">1.0</span></span>|
|[<span data-ttu-id="ec1d2-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ec1d2-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ec1d2-181">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ec1d2-181">Compose or Read</span></span>|
