---
title: Namespace do Office – conjunto de requisitos 1,1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: eda5e1fb5f2c11ae91e4a1479892c36ec23e1897
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451910"
---
# <a name="office"></a><span data-ttu-id="5501f-102">Office</span><span class="sxs-lookup"><span data-stu-id="5501f-102">Office</span></span>

<span data-ttu-id="5501f-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="5501f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5501f-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5501f-105">Requirements</span></span>

|<span data-ttu-id="5501f-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="5501f-106">Requirement</span></span>| <span data-ttu-id="5501f-107">Valor</span><span class="sxs-lookup"><span data-stu-id="5501f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="5501f-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5501f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5501f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="5501f-109">1.0</span></span>|
|[<span data-ttu-id="5501f-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5501f-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5501f-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5501f-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="5501f-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="5501f-112">Namespaces</span></span>

<span data-ttu-id="5501f-113">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="5501f-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="5501f-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="5501f-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="5501f-115">Membros</span><span class="sxs-lookup"><span data-stu-id="5501f-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="5501f-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="5501f-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="5501f-117">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="5501f-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5501f-118">Tipo</span><span class="sxs-lookup"><span data-stu-id="5501f-118">Type</span></span>

*   <span data-ttu-id="5501f-119">String</span><span class="sxs-lookup"><span data-stu-id="5501f-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5501f-120">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5501f-120">Properties:</span></span>

|<span data-ttu-id="5501f-121">Name</span><span class="sxs-lookup"><span data-stu-id="5501f-121">Name</span></span>| <span data-ttu-id="5501f-122">Tipo</span><span class="sxs-lookup"><span data-stu-id="5501f-122">Type</span></span>| <span data-ttu-id="5501f-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="5501f-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5501f-124">String</span><span class="sxs-lookup"><span data-stu-id="5501f-124">String</span></span>|<span data-ttu-id="5501f-125">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="5501f-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5501f-126">String</span><span class="sxs-lookup"><span data-stu-id="5501f-126">String</span></span>|<span data-ttu-id="5501f-127">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="5501f-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5501f-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5501f-128">Requirements</span></span>

|<span data-ttu-id="5501f-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="5501f-129">Requirement</span></span>| <span data-ttu-id="5501f-130">Valor</span><span class="sxs-lookup"><span data-stu-id="5501f-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="5501f-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5501f-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5501f-132">1.0</span><span class="sxs-lookup"><span data-stu-id="5501f-132">1.0</span></span>|
|[<span data-ttu-id="5501f-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5501f-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5501f-134">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5501f-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="5501f-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="5501f-135">CoercionType :String</span></span>

<span data-ttu-id="5501f-136">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="5501f-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5501f-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="5501f-137">Type</span></span>

*   <span data-ttu-id="5501f-138">String</span><span class="sxs-lookup"><span data-stu-id="5501f-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5501f-139">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5501f-139">Properties:</span></span>

|<span data-ttu-id="5501f-140">Name</span><span class="sxs-lookup"><span data-stu-id="5501f-140">Name</span></span>| <span data-ttu-id="5501f-141">Tipo</span><span class="sxs-lookup"><span data-stu-id="5501f-141">Type</span></span>| <span data-ttu-id="5501f-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="5501f-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5501f-143">String</span><span class="sxs-lookup"><span data-stu-id="5501f-143">String</span></span>|<span data-ttu-id="5501f-144">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="5501f-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5501f-145">String</span><span class="sxs-lookup"><span data-stu-id="5501f-145">String</span></span>|<span data-ttu-id="5501f-146">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="5501f-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5501f-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5501f-147">Requirements</span></span>

|<span data-ttu-id="5501f-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="5501f-148">Requirement</span></span>| <span data-ttu-id="5501f-149">Valor</span><span class="sxs-lookup"><span data-stu-id="5501f-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="5501f-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5501f-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5501f-151">1.0</span><span class="sxs-lookup"><span data-stu-id="5501f-151">1.0</span></span>|
|[<span data-ttu-id="5501f-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5501f-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5501f-153">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5501f-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="5501f-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="5501f-154">SourceProperty :String</span></span>

<span data-ttu-id="5501f-155">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="5501f-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5501f-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="5501f-156">Type</span></span>

*   <span data-ttu-id="5501f-157">String</span><span class="sxs-lookup"><span data-stu-id="5501f-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5501f-158">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5501f-158">Properties:</span></span>

|<span data-ttu-id="5501f-159">Name</span><span class="sxs-lookup"><span data-stu-id="5501f-159">Name</span></span>| <span data-ttu-id="5501f-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="5501f-160">Type</span></span>| <span data-ttu-id="5501f-161">Descrição</span><span class="sxs-lookup"><span data-stu-id="5501f-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5501f-162">String</span><span class="sxs-lookup"><span data-stu-id="5501f-162">String</span></span>|<span data-ttu-id="5501f-163">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5501f-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5501f-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5501f-164">String</span></span>|<span data-ttu-id="5501f-165">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5501f-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5501f-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5501f-166">Requirements</span></span>

|<span data-ttu-id="5501f-167">Requisito</span><span class="sxs-lookup"><span data-stu-id="5501f-167">Requirement</span></span>| <span data-ttu-id="5501f-168">Valor</span><span class="sxs-lookup"><span data-stu-id="5501f-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="5501f-169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5501f-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5501f-170">1.0</span><span class="sxs-lookup"><span data-stu-id="5501f-170">1.0</span></span>|
|[<span data-ttu-id="5501f-171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5501f-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5501f-172">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5501f-172">Compose or Read</span></span>|
