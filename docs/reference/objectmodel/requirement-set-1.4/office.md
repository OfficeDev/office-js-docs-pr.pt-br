---
title: Namespace do Office – conjunto de requisitos 1,4
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 2617ba3c80f44b1cddab568f94044a95a3061065
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064674"
---
# <a name="office"></a><span data-ttu-id="931ee-102">Office</span><span class="sxs-lookup"><span data-stu-id="931ee-102">Office</span></span>

<span data-ttu-id="931ee-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="931ee-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="931ee-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="931ee-105">Requirements</span></span>

|<span data-ttu-id="931ee-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="931ee-106">Requirement</span></span>| <span data-ttu-id="931ee-107">Valor</span><span class="sxs-lookup"><span data-stu-id="931ee-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="931ee-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="931ee-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="931ee-109">1.0</span><span class="sxs-lookup"><span data-stu-id="931ee-109">1.0</span></span>|
|[<span data-ttu-id="931ee-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="931ee-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="931ee-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="931ee-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="931ee-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="931ee-112">Namespaces</span></span>

<span data-ttu-id="931ee-113">[context](Office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="931ee-113">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="931ee-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="931ee-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="931ee-115">Membros</span><span class="sxs-lookup"><span data-stu-id="931ee-115">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="931ee-116">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="931ee-116">AsyncResultStatus: String</span></span>

<span data-ttu-id="931ee-117">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="931ee-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="931ee-118">Tipo</span><span class="sxs-lookup"><span data-stu-id="931ee-118">Type</span></span>

*   <span data-ttu-id="931ee-119">String</span><span class="sxs-lookup"><span data-stu-id="931ee-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="931ee-120">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="931ee-120">Properties:</span></span>

|<span data-ttu-id="931ee-121">Nome</span><span class="sxs-lookup"><span data-stu-id="931ee-121">Name</span></span>| <span data-ttu-id="931ee-122">Tipo</span><span class="sxs-lookup"><span data-stu-id="931ee-122">Type</span></span>| <span data-ttu-id="931ee-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="931ee-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="931ee-124">String</span><span class="sxs-lookup"><span data-stu-id="931ee-124">String</span></span>|<span data-ttu-id="931ee-125">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="931ee-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="931ee-126">String</span><span class="sxs-lookup"><span data-stu-id="931ee-126">String</span></span>|<span data-ttu-id="931ee-127">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="931ee-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="931ee-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="931ee-128">Requirements</span></span>

|<span data-ttu-id="931ee-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="931ee-129">Requirement</span></span>| <span data-ttu-id="931ee-130">Valor</span><span class="sxs-lookup"><span data-stu-id="931ee-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="931ee-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="931ee-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="931ee-132">1.0</span><span class="sxs-lookup"><span data-stu-id="931ee-132">1.0</span></span>|
|[<span data-ttu-id="931ee-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="931ee-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="931ee-134">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="931ee-134">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="931ee-135">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="931ee-135">CoercionType: String</span></span>

<span data-ttu-id="931ee-136">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="931ee-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="931ee-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="931ee-137">Type</span></span>

*   <span data-ttu-id="931ee-138">String</span><span class="sxs-lookup"><span data-stu-id="931ee-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="931ee-139">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="931ee-139">Properties:</span></span>

|<span data-ttu-id="931ee-140">Nome</span><span class="sxs-lookup"><span data-stu-id="931ee-140">Name</span></span>| <span data-ttu-id="931ee-141">Tipo</span><span class="sxs-lookup"><span data-stu-id="931ee-141">Type</span></span>| <span data-ttu-id="931ee-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="931ee-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="931ee-143">String</span><span class="sxs-lookup"><span data-stu-id="931ee-143">String</span></span>|<span data-ttu-id="931ee-144">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="931ee-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="931ee-145">String</span><span class="sxs-lookup"><span data-stu-id="931ee-145">String</span></span>|<span data-ttu-id="931ee-146">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="931ee-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="931ee-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="931ee-147">Requirements</span></span>

|<span data-ttu-id="931ee-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="931ee-148">Requirement</span></span>| <span data-ttu-id="931ee-149">Valor</span><span class="sxs-lookup"><span data-stu-id="931ee-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="931ee-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="931ee-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="931ee-151">1.0</span><span class="sxs-lookup"><span data-stu-id="931ee-151">1.0</span></span>|
|[<span data-ttu-id="931ee-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="931ee-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="931ee-153">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="931ee-153">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="931ee-154">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="931ee-154">SourceProperty: String</span></span>

<span data-ttu-id="931ee-155">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="931ee-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="931ee-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="931ee-156">Type</span></span>

*   <span data-ttu-id="931ee-157">String</span><span class="sxs-lookup"><span data-stu-id="931ee-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="931ee-158">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="931ee-158">Properties:</span></span>

|<span data-ttu-id="931ee-159">Nome</span><span class="sxs-lookup"><span data-stu-id="931ee-159">Name</span></span>| <span data-ttu-id="931ee-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="931ee-160">Type</span></span>| <span data-ttu-id="931ee-161">Descrição</span><span class="sxs-lookup"><span data-stu-id="931ee-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="931ee-162">String</span><span class="sxs-lookup"><span data-stu-id="931ee-162">String</span></span>|<span data-ttu-id="931ee-163">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="931ee-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="931ee-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="931ee-164">String</span></span>|<span data-ttu-id="931ee-165">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="931ee-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="931ee-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="931ee-166">Requirements</span></span>

|<span data-ttu-id="931ee-167">Requisito</span><span class="sxs-lookup"><span data-stu-id="931ee-167">Requirement</span></span>| <span data-ttu-id="931ee-168">Valor</span><span class="sxs-lookup"><span data-stu-id="931ee-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="931ee-169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="931ee-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="931ee-170">1.0</span><span class="sxs-lookup"><span data-stu-id="931ee-170">1.0</span></span>|
|[<span data-ttu-id="931ee-171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="931ee-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="931ee-172">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="931ee-172">Compose or Read</span></span>|
