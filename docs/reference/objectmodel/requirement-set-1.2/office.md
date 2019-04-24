---
title: Namespace do Office – conjunto de requisitos 1.2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: dc98d4c2da6e8f9ca294a6c686cf081478e1bb24
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450307"
---
# <a name="office"></a><span data-ttu-id="d5cb0-102">Office</span><span class="sxs-lookup"><span data-stu-id="d5cb0-102">Office</span></span>

<span data-ttu-id="d5cb0-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d5cb0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d5cb0-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d5cb0-105">Requirements</span></span>

|<span data-ttu-id="d5cb0-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="d5cb0-106">Requirement</span></span>| <span data-ttu-id="d5cb0-107">Valor</span><span class="sxs-lookup"><span data-stu-id="d5cb0-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5cb0-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d5cb0-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5cb0-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d5cb0-109">1.0</span></span>|
|[<span data-ttu-id="d5cb0-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d5cb0-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5cb0-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d5cb0-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="d5cb0-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="d5cb0-112">Namespaces</span></span>

<span data-ttu-id="d5cb0-113">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d5cb0-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d5cb0-114">[MailboxEnums](/javascript/api/outlook_1_2/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="d5cb0-114">[MailboxEnums](/javascript/api/outlook_1_2/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="d5cb0-115">Membros</span><span class="sxs-lookup"><span data-stu-id="d5cb0-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="d5cb0-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="d5cb0-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="d5cb0-117">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="d5cb0-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d5cb0-118">Tipo</span><span class="sxs-lookup"><span data-stu-id="d5cb0-118">Type</span></span>

*   <span data-ttu-id="d5cb0-119">String</span><span class="sxs-lookup"><span data-stu-id="d5cb0-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d5cb0-120">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="d5cb0-120">Properties:</span></span>

|<span data-ttu-id="d5cb0-121">Name</span><span class="sxs-lookup"><span data-stu-id="d5cb0-121">Name</span></span>| <span data-ttu-id="d5cb0-122">Tipo</span><span class="sxs-lookup"><span data-stu-id="d5cb0-122">Type</span></span>| <span data-ttu-id="d5cb0-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="d5cb0-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d5cb0-124">String</span><span class="sxs-lookup"><span data-stu-id="d5cb0-124">String</span></span>|<span data-ttu-id="d5cb0-125">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="d5cb0-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d5cb0-126">String</span><span class="sxs-lookup"><span data-stu-id="d5cb0-126">String</span></span>|<span data-ttu-id="d5cb0-127">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="d5cb0-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5cb0-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d5cb0-128">Requirements</span></span>

|<span data-ttu-id="d5cb0-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="d5cb0-129">Requirement</span></span>| <span data-ttu-id="d5cb0-130">Valor</span><span class="sxs-lookup"><span data-stu-id="d5cb0-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5cb0-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d5cb0-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5cb0-132">1.0</span><span class="sxs-lookup"><span data-stu-id="d5cb0-132">1.0</span></span>|
|[<span data-ttu-id="d5cb0-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d5cb0-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5cb0-134">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d5cb0-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="d5cb0-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="d5cb0-135">CoercionType :String</span></span>

<span data-ttu-id="d5cb0-136">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="d5cb0-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d5cb0-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="d5cb0-137">Type</span></span>

*   <span data-ttu-id="d5cb0-138">String</span><span class="sxs-lookup"><span data-stu-id="d5cb0-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d5cb0-139">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="d5cb0-139">Properties:</span></span>

|<span data-ttu-id="d5cb0-140">Name</span><span class="sxs-lookup"><span data-stu-id="d5cb0-140">Name</span></span>| <span data-ttu-id="d5cb0-141">Tipo</span><span class="sxs-lookup"><span data-stu-id="d5cb0-141">Type</span></span>| <span data-ttu-id="d5cb0-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="d5cb0-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d5cb0-143">String</span><span class="sxs-lookup"><span data-stu-id="d5cb0-143">String</span></span>|<span data-ttu-id="d5cb0-144">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="d5cb0-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d5cb0-145">String</span><span class="sxs-lookup"><span data-stu-id="d5cb0-145">String</span></span>|<span data-ttu-id="d5cb0-146">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="d5cb0-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5cb0-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d5cb0-147">Requirements</span></span>

|<span data-ttu-id="d5cb0-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="d5cb0-148">Requirement</span></span>| <span data-ttu-id="d5cb0-149">Valor</span><span class="sxs-lookup"><span data-stu-id="d5cb0-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5cb0-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d5cb0-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5cb0-151">1.0</span><span class="sxs-lookup"><span data-stu-id="d5cb0-151">1.0</span></span>|
|[<span data-ttu-id="d5cb0-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d5cb0-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5cb0-153">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d5cb0-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="d5cb0-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="d5cb0-154">SourceProperty :String</span></span>

<span data-ttu-id="d5cb0-155">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="d5cb0-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d5cb0-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="d5cb0-156">Type</span></span>

*   <span data-ttu-id="d5cb0-157">String</span><span class="sxs-lookup"><span data-stu-id="d5cb0-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d5cb0-158">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="d5cb0-158">Properties:</span></span>

|<span data-ttu-id="d5cb0-159">Name</span><span class="sxs-lookup"><span data-stu-id="d5cb0-159">Name</span></span>| <span data-ttu-id="d5cb0-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="d5cb0-160">Type</span></span>| <span data-ttu-id="d5cb0-161">Descrição</span><span class="sxs-lookup"><span data-stu-id="d5cb0-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d5cb0-162">String</span><span class="sxs-lookup"><span data-stu-id="d5cb0-162">String</span></span>|<span data-ttu-id="d5cb0-163">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="d5cb0-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d5cb0-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d5cb0-164">String</span></span>|<span data-ttu-id="d5cb0-165">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="d5cb0-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5cb0-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d5cb0-166">Requirements</span></span>

|<span data-ttu-id="d5cb0-167">Requisito</span><span class="sxs-lookup"><span data-stu-id="d5cb0-167">Requirement</span></span>| <span data-ttu-id="d5cb0-168">Valor</span><span class="sxs-lookup"><span data-stu-id="d5cb0-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5cb0-169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d5cb0-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5cb0-170">1.0</span><span class="sxs-lookup"><span data-stu-id="d5cb0-170">1.0</span></span>|
|[<span data-ttu-id="d5cb0-171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d5cb0-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5cb0-172">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d5cb0-172">Compose or Read</span></span>|
