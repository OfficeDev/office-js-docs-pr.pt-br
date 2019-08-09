---
title: Namespace do Office – conjunto de requisitos 1,1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 393804c6e76e74ffdb8f860cddc9b454593f9982
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268471"
---
# <a name="office"></a><span data-ttu-id="44f27-102">Office</span><span class="sxs-lookup"><span data-stu-id="44f27-102">Office</span></span>

<span data-ttu-id="44f27-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="44f27-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="44f27-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44f27-105">Requirements</span></span>

|<span data-ttu-id="44f27-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="44f27-106">Requirement</span></span>| <span data-ttu-id="44f27-107">Valor</span><span class="sxs-lookup"><span data-stu-id="44f27-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="44f27-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44f27-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="44f27-109">1.0</span><span class="sxs-lookup"><span data-stu-id="44f27-109">1.0</span></span>|
|[<span data-ttu-id="44f27-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44f27-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="44f27-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44f27-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="44f27-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="44f27-112">Members and methods</span></span>

| <span data-ttu-id="44f27-113">Membro</span><span class="sxs-lookup"><span data-stu-id="44f27-113">Member</span></span> | <span data-ttu-id="44f27-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="44f27-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="44f27-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="44f27-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="44f27-116">Membro</span><span class="sxs-lookup"><span data-stu-id="44f27-116">Member</span></span> |
| [<span data-ttu-id="44f27-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="44f27-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="44f27-118">Membro</span><span class="sxs-lookup"><span data-stu-id="44f27-118">Member</span></span> |
| [<span data-ttu-id="44f27-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="44f27-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="44f27-120">Membro</span><span class="sxs-lookup"><span data-stu-id="44f27-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="44f27-121">Namespaces</span><span class="sxs-lookup"><span data-stu-id="44f27-121">Namespaces</span></span>

<span data-ttu-id="44f27-122">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="44f27-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="44f27-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.1): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="44f27-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.1): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="44f27-124">Membros</span><span class="sxs-lookup"><span data-stu-id="44f27-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="44f27-125">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="44f27-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="44f27-126">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="44f27-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="44f27-127">Tipo</span><span class="sxs-lookup"><span data-stu-id="44f27-127">Type</span></span>

*   <span data-ttu-id="44f27-128">String</span><span class="sxs-lookup"><span data-stu-id="44f27-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="44f27-129">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="44f27-129">Properties:</span></span>

|<span data-ttu-id="44f27-130">Nome</span><span class="sxs-lookup"><span data-stu-id="44f27-130">Name</span></span>| <span data-ttu-id="44f27-131">Tipo</span><span class="sxs-lookup"><span data-stu-id="44f27-131">Type</span></span>| <span data-ttu-id="44f27-132">Descrição</span><span class="sxs-lookup"><span data-stu-id="44f27-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="44f27-133">String</span><span class="sxs-lookup"><span data-stu-id="44f27-133">String</span></span>|<span data-ttu-id="44f27-134">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="44f27-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="44f27-135">String</span><span class="sxs-lookup"><span data-stu-id="44f27-135">String</span></span>|<span data-ttu-id="44f27-136">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="44f27-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44f27-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44f27-137">Requirements</span></span>

|<span data-ttu-id="44f27-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="44f27-138">Requirement</span></span>| <span data-ttu-id="44f27-139">Valor</span><span class="sxs-lookup"><span data-stu-id="44f27-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="44f27-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44f27-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="44f27-141">1.0</span><span class="sxs-lookup"><span data-stu-id="44f27-141">1.0</span></span>|
|[<span data-ttu-id="44f27-142">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44f27-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="44f27-143">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44f27-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="44f27-144">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="44f27-144">CoercionType: String</span></span>

<span data-ttu-id="44f27-145">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="44f27-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="44f27-146">Tipo</span><span class="sxs-lookup"><span data-stu-id="44f27-146">Type</span></span>

*   <span data-ttu-id="44f27-147">String</span><span class="sxs-lookup"><span data-stu-id="44f27-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="44f27-148">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="44f27-148">Properties:</span></span>

|<span data-ttu-id="44f27-149">Nome</span><span class="sxs-lookup"><span data-stu-id="44f27-149">Name</span></span>| <span data-ttu-id="44f27-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="44f27-150">Type</span></span>| <span data-ttu-id="44f27-151">Descrição</span><span class="sxs-lookup"><span data-stu-id="44f27-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="44f27-152">String</span><span class="sxs-lookup"><span data-stu-id="44f27-152">String</span></span>|<span data-ttu-id="44f27-153">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="44f27-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="44f27-154">String</span><span class="sxs-lookup"><span data-stu-id="44f27-154">String</span></span>|<span data-ttu-id="44f27-155">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="44f27-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44f27-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44f27-156">Requirements</span></span>

|<span data-ttu-id="44f27-157">Requisito</span><span class="sxs-lookup"><span data-stu-id="44f27-157">Requirement</span></span>| <span data-ttu-id="44f27-158">Valor</span><span class="sxs-lookup"><span data-stu-id="44f27-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="44f27-159">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44f27-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="44f27-160">1.0</span><span class="sxs-lookup"><span data-stu-id="44f27-160">1.0</span></span>|
|[<span data-ttu-id="44f27-161">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44f27-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="44f27-162">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44f27-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="44f27-163">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="44f27-163">SourceProperty: String</span></span>

<span data-ttu-id="44f27-164">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="44f27-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="44f27-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="44f27-165">Type</span></span>

*   <span data-ttu-id="44f27-166">String</span><span class="sxs-lookup"><span data-stu-id="44f27-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="44f27-167">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="44f27-167">Properties:</span></span>

|<span data-ttu-id="44f27-168">Nome</span><span class="sxs-lookup"><span data-stu-id="44f27-168">Name</span></span>| <span data-ttu-id="44f27-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="44f27-169">Type</span></span>| <span data-ttu-id="44f27-170">Descrição</span><span class="sxs-lookup"><span data-stu-id="44f27-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="44f27-171">String</span><span class="sxs-lookup"><span data-stu-id="44f27-171">String</span></span>|<span data-ttu-id="44f27-172">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="44f27-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="44f27-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="44f27-173">String</span></span>|<span data-ttu-id="44f27-174">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="44f27-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44f27-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="44f27-175">Requirements</span></span>

|<span data-ttu-id="44f27-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="44f27-176">Requirement</span></span>| <span data-ttu-id="44f27-177">Valor</span><span class="sxs-lookup"><span data-stu-id="44f27-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="44f27-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="44f27-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="44f27-179">1.0</span><span class="sxs-lookup"><span data-stu-id="44f27-179">1.0</span></span>|
|[<span data-ttu-id="44f27-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="44f27-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="44f27-181">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="44f27-181">Compose or Read</span></span>|
