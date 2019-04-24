---
title: Namespace do Office – conjunto de requisitos 1,3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: ef01b7da3d447af852a5558853e0902eab815dd3
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451896"
---
# <a name="office"></a><span data-ttu-id="55953-102">Office</span><span class="sxs-lookup"><span data-stu-id="55953-102">Office</span></span>

<span data-ttu-id="55953-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="55953-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="55953-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="55953-105">Requirements</span></span>

|<span data-ttu-id="55953-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="55953-106">Requirement</span></span>| <span data-ttu-id="55953-107">Valor</span><span class="sxs-lookup"><span data-stu-id="55953-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="55953-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="55953-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55953-109">1.0</span><span class="sxs-lookup"><span data-stu-id="55953-109">1.0</span></span>|
|[<span data-ttu-id="55953-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="55953-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55953-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="55953-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="55953-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="55953-112">Namespaces</span></span>

<span data-ttu-id="55953-113">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="55953-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="55953-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="55953-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="55953-115">Membros</span><span class="sxs-lookup"><span data-stu-id="55953-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="55953-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="55953-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="55953-117">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="55953-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="55953-118">Tipo</span><span class="sxs-lookup"><span data-stu-id="55953-118">Type</span></span>

*   <span data-ttu-id="55953-119">String</span><span class="sxs-lookup"><span data-stu-id="55953-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="55953-120">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="55953-120">Properties:</span></span>

|<span data-ttu-id="55953-121">Name</span><span class="sxs-lookup"><span data-stu-id="55953-121">Name</span></span>| <span data-ttu-id="55953-122">Tipo</span><span class="sxs-lookup"><span data-stu-id="55953-122">Type</span></span>| <span data-ttu-id="55953-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="55953-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="55953-124">String</span><span class="sxs-lookup"><span data-stu-id="55953-124">String</span></span>|<span data-ttu-id="55953-125">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="55953-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="55953-126">String</span><span class="sxs-lookup"><span data-stu-id="55953-126">String</span></span>|<span data-ttu-id="55953-127">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="55953-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55953-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="55953-128">Requirements</span></span>

|<span data-ttu-id="55953-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="55953-129">Requirement</span></span>| <span data-ttu-id="55953-130">Valor</span><span class="sxs-lookup"><span data-stu-id="55953-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="55953-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="55953-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55953-132">1.0</span><span class="sxs-lookup"><span data-stu-id="55953-132">1.0</span></span>|
|[<span data-ttu-id="55953-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="55953-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55953-134">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="55953-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="55953-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="55953-135">CoercionType :String</span></span>

<span data-ttu-id="55953-136">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="55953-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="55953-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="55953-137">Type</span></span>

*   <span data-ttu-id="55953-138">String</span><span class="sxs-lookup"><span data-stu-id="55953-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="55953-139">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="55953-139">Properties:</span></span>

|<span data-ttu-id="55953-140">Name</span><span class="sxs-lookup"><span data-stu-id="55953-140">Name</span></span>| <span data-ttu-id="55953-141">Tipo</span><span class="sxs-lookup"><span data-stu-id="55953-141">Type</span></span>| <span data-ttu-id="55953-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="55953-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="55953-143">String</span><span class="sxs-lookup"><span data-stu-id="55953-143">String</span></span>|<span data-ttu-id="55953-144">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="55953-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="55953-145">String</span><span class="sxs-lookup"><span data-stu-id="55953-145">String</span></span>|<span data-ttu-id="55953-146">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="55953-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55953-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="55953-147">Requirements</span></span>

|<span data-ttu-id="55953-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="55953-148">Requirement</span></span>| <span data-ttu-id="55953-149">Valor</span><span class="sxs-lookup"><span data-stu-id="55953-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="55953-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="55953-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55953-151">1.0</span><span class="sxs-lookup"><span data-stu-id="55953-151">1.0</span></span>|
|[<span data-ttu-id="55953-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="55953-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55953-153">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="55953-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="55953-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="55953-154">SourceProperty :String</span></span>

<span data-ttu-id="55953-155">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="55953-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="55953-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="55953-156">Type</span></span>

*   <span data-ttu-id="55953-157">String</span><span class="sxs-lookup"><span data-stu-id="55953-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="55953-158">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="55953-158">Properties:</span></span>

|<span data-ttu-id="55953-159">Name</span><span class="sxs-lookup"><span data-stu-id="55953-159">Name</span></span>| <span data-ttu-id="55953-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="55953-160">Type</span></span>| <span data-ttu-id="55953-161">Descrição</span><span class="sxs-lookup"><span data-stu-id="55953-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="55953-162">String</span><span class="sxs-lookup"><span data-stu-id="55953-162">String</span></span>|<span data-ttu-id="55953-163">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="55953-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="55953-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="55953-164">String</span></span>|<span data-ttu-id="55953-165">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="55953-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="55953-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="55953-166">Requirements</span></span>

|<span data-ttu-id="55953-167">Requisito</span><span class="sxs-lookup"><span data-stu-id="55953-167">Requirement</span></span>| <span data-ttu-id="55953-168">Valor</span><span class="sxs-lookup"><span data-stu-id="55953-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="55953-169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="55953-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="55953-170">1.0</span><span class="sxs-lookup"><span data-stu-id="55953-170">1.0</span></span>|
|[<span data-ttu-id="55953-171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="55953-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="55953-172">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="55953-172">Compose or Read</span></span>|
