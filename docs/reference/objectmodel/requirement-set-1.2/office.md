---
title: Namespace do Office – conjunto de requisitos 1.2
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 9dd492046df6325c5c2cdb04dbd1c8bc331b3471
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064393"
---
# <a name="office"></a><span data-ttu-id="7f6d3-102">Office</span><span class="sxs-lookup"><span data-stu-id="7f6d3-102">Office</span></span>

<span data-ttu-id="7f6d3-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="7f6d3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7f6d3-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7f6d3-105">Requirements</span></span>

|<span data-ttu-id="7f6d3-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="7f6d3-106">Requirement</span></span>| <span data-ttu-id="7f6d3-107">Valor</span><span class="sxs-lookup"><span data-stu-id="7f6d3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f6d3-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7f6d3-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f6d3-109">1.0</span><span class="sxs-lookup"><span data-stu-id="7f6d3-109">1.0</span></span>|
|[<span data-ttu-id="7f6d3-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7f6d3-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f6d3-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7f6d3-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="7f6d3-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="7f6d3-112">Namespaces</span></span>

<span data-ttu-id="7f6d3-113">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="7f6d3-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="7f6d3-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="7f6d3-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="7f6d3-115">Membros</span><span class="sxs-lookup"><span data-stu-id="7f6d3-115">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="7f6d3-116">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7f6d3-116">AsyncResultStatus: String</span></span>

<span data-ttu-id="7f6d3-117">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="7f6d3-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="7f6d3-118">Tipo</span><span class="sxs-lookup"><span data-stu-id="7f6d3-118">Type</span></span>

*   <span data-ttu-id="7f6d3-119">String</span><span class="sxs-lookup"><span data-stu-id="7f6d3-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7f6d3-120">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7f6d3-120">Properties:</span></span>

|<span data-ttu-id="7f6d3-121">Nome</span><span class="sxs-lookup"><span data-stu-id="7f6d3-121">Name</span></span>| <span data-ttu-id="7f6d3-122">Tipo</span><span class="sxs-lookup"><span data-stu-id="7f6d3-122">Type</span></span>| <span data-ttu-id="7f6d3-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="7f6d3-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="7f6d3-124">String</span><span class="sxs-lookup"><span data-stu-id="7f6d3-124">String</span></span>|<span data-ttu-id="7f6d3-125">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="7f6d3-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="7f6d3-126">String</span><span class="sxs-lookup"><span data-stu-id="7f6d3-126">String</span></span>|<span data-ttu-id="7f6d3-127">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="7f6d3-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f6d3-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7f6d3-128">Requirements</span></span>

|<span data-ttu-id="7f6d3-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="7f6d3-129">Requirement</span></span>| <span data-ttu-id="7f6d3-130">Valor</span><span class="sxs-lookup"><span data-stu-id="7f6d3-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f6d3-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7f6d3-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f6d3-132">1.0</span><span class="sxs-lookup"><span data-stu-id="7f6d3-132">1.0</span></span>|
|[<span data-ttu-id="7f6d3-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7f6d3-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f6d3-134">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7f6d3-134">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="7f6d3-135">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7f6d3-135">CoercionType: String</span></span>

<span data-ttu-id="7f6d3-136">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="7f6d3-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7f6d3-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="7f6d3-137">Type</span></span>

*   <span data-ttu-id="7f6d3-138">String</span><span class="sxs-lookup"><span data-stu-id="7f6d3-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7f6d3-139">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7f6d3-139">Properties:</span></span>

|<span data-ttu-id="7f6d3-140">Nome</span><span class="sxs-lookup"><span data-stu-id="7f6d3-140">Name</span></span>| <span data-ttu-id="7f6d3-141">Tipo</span><span class="sxs-lookup"><span data-stu-id="7f6d3-141">Type</span></span>| <span data-ttu-id="7f6d3-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="7f6d3-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="7f6d3-143">String</span><span class="sxs-lookup"><span data-stu-id="7f6d3-143">String</span></span>|<span data-ttu-id="7f6d3-144">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="7f6d3-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="7f6d3-145">String</span><span class="sxs-lookup"><span data-stu-id="7f6d3-145">String</span></span>|<span data-ttu-id="7f6d3-146">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="7f6d3-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f6d3-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7f6d3-147">Requirements</span></span>

|<span data-ttu-id="7f6d3-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="7f6d3-148">Requirement</span></span>| <span data-ttu-id="7f6d3-149">Valor</span><span class="sxs-lookup"><span data-stu-id="7f6d3-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f6d3-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7f6d3-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f6d3-151">1.0</span><span class="sxs-lookup"><span data-stu-id="7f6d3-151">1.0</span></span>|
|[<span data-ttu-id="7f6d3-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7f6d3-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f6d3-153">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7f6d3-153">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="7f6d3-154">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7f6d3-154">SourceProperty: String</span></span>

<span data-ttu-id="7f6d3-155">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="7f6d3-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7f6d3-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="7f6d3-156">Type</span></span>

*   <span data-ttu-id="7f6d3-157">String</span><span class="sxs-lookup"><span data-stu-id="7f6d3-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7f6d3-158">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7f6d3-158">Properties:</span></span>

|<span data-ttu-id="7f6d3-159">Nome</span><span class="sxs-lookup"><span data-stu-id="7f6d3-159">Name</span></span>| <span data-ttu-id="7f6d3-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="7f6d3-160">Type</span></span>| <span data-ttu-id="7f6d3-161">Descrição</span><span class="sxs-lookup"><span data-stu-id="7f6d3-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="7f6d3-162">String</span><span class="sxs-lookup"><span data-stu-id="7f6d3-162">String</span></span>|<span data-ttu-id="7f6d3-163">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7f6d3-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="7f6d3-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7f6d3-164">String</span></span>|<span data-ttu-id="7f6d3-165">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7f6d3-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7f6d3-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7f6d3-166">Requirements</span></span>

|<span data-ttu-id="7f6d3-167">Requisito</span><span class="sxs-lookup"><span data-stu-id="7f6d3-167">Requirement</span></span>| <span data-ttu-id="7f6d3-168">Valor</span><span class="sxs-lookup"><span data-stu-id="7f6d3-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="7f6d3-169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7f6d3-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7f6d3-170">1.0</span><span class="sxs-lookup"><span data-stu-id="7f6d3-170">1.0</span></span>|
|[<span data-ttu-id="7f6d3-171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7f6d3-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7f6d3-172">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7f6d3-172">Compose or Read</span></span>|
