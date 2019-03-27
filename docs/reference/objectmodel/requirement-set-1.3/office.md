---
title: Namespace do Office – conjunto de requisitos 1,3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: ef01b7da3d447af852a5558853e0902eab815dd3
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871217"
---
# <a name="office"></a><span data-ttu-id="ad969-102">Office</span><span class="sxs-lookup"><span data-stu-id="ad969-102">Office</span></span>

<span data-ttu-id="ad969-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ad969-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad969-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ad969-105">Requirements</span></span>

|<span data-ttu-id="ad969-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="ad969-106">Requirement</span></span>| <span data-ttu-id="ad969-107">Valor</span><span class="sxs-lookup"><span data-stu-id="ad969-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad969-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ad969-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad969-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ad969-109">1.0</span></span>|
|[<span data-ttu-id="ad969-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ad969-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad969-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ad969-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="ad969-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="ad969-112">Namespaces</span></span>

<span data-ttu-id="ad969-113">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="ad969-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="ad969-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="ad969-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="ad969-115">Membros</span><span class="sxs-lookup"><span data-stu-id="ad969-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="ad969-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="ad969-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="ad969-117">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="ad969-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ad969-118">Tipo</span><span class="sxs-lookup"><span data-stu-id="ad969-118">Type</span></span>

*   <span data-ttu-id="ad969-119">String</span><span class="sxs-lookup"><span data-stu-id="ad969-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ad969-120">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ad969-120">Properties:</span></span>

|<span data-ttu-id="ad969-121">Nome</span><span class="sxs-lookup"><span data-stu-id="ad969-121">Name</span></span>| <span data-ttu-id="ad969-122">Tipo</span><span class="sxs-lookup"><span data-stu-id="ad969-122">Type</span></span>| <span data-ttu-id="ad969-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="ad969-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ad969-124">String</span><span class="sxs-lookup"><span data-stu-id="ad969-124">String</span></span>|<span data-ttu-id="ad969-125">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="ad969-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ad969-126">String</span><span class="sxs-lookup"><span data-stu-id="ad969-126">String</span></span>|<span data-ttu-id="ad969-127">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="ad969-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad969-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ad969-128">Requirements</span></span>

|<span data-ttu-id="ad969-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="ad969-129">Requirement</span></span>| <span data-ttu-id="ad969-130">Valor</span><span class="sxs-lookup"><span data-stu-id="ad969-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad969-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ad969-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad969-132">1.0</span><span class="sxs-lookup"><span data-stu-id="ad969-132">1.0</span></span>|
|[<span data-ttu-id="ad969-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ad969-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad969-134">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ad969-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="ad969-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="ad969-135">CoercionType :String</span></span>

<span data-ttu-id="ad969-136">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="ad969-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ad969-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="ad969-137">Type</span></span>

*   <span data-ttu-id="ad969-138">String</span><span class="sxs-lookup"><span data-stu-id="ad969-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ad969-139">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ad969-139">Properties:</span></span>

|<span data-ttu-id="ad969-140">Nome</span><span class="sxs-lookup"><span data-stu-id="ad969-140">Name</span></span>| <span data-ttu-id="ad969-141">Tipo</span><span class="sxs-lookup"><span data-stu-id="ad969-141">Type</span></span>| <span data-ttu-id="ad969-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="ad969-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ad969-143">String</span><span class="sxs-lookup"><span data-stu-id="ad969-143">String</span></span>|<span data-ttu-id="ad969-144">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="ad969-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ad969-145">String</span><span class="sxs-lookup"><span data-stu-id="ad969-145">String</span></span>|<span data-ttu-id="ad969-146">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="ad969-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad969-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ad969-147">Requirements</span></span>

|<span data-ttu-id="ad969-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="ad969-148">Requirement</span></span>| <span data-ttu-id="ad969-149">Valor</span><span class="sxs-lookup"><span data-stu-id="ad969-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad969-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ad969-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad969-151">1.0</span><span class="sxs-lookup"><span data-stu-id="ad969-151">1.0</span></span>|
|[<span data-ttu-id="ad969-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ad969-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad969-153">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ad969-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="ad969-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="ad969-154">SourceProperty :String</span></span>

<span data-ttu-id="ad969-155">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="ad969-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ad969-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="ad969-156">Type</span></span>

*   <span data-ttu-id="ad969-157">String</span><span class="sxs-lookup"><span data-stu-id="ad969-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ad969-158">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ad969-158">Properties:</span></span>

|<span data-ttu-id="ad969-159">Nome</span><span class="sxs-lookup"><span data-stu-id="ad969-159">Name</span></span>| <span data-ttu-id="ad969-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="ad969-160">Type</span></span>| <span data-ttu-id="ad969-161">Descrição</span><span class="sxs-lookup"><span data-stu-id="ad969-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ad969-162">String</span><span class="sxs-lookup"><span data-stu-id="ad969-162">String</span></span>|<span data-ttu-id="ad969-163">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ad969-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ad969-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ad969-164">String</span></span>|<span data-ttu-id="ad969-165">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ad969-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad969-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ad969-166">Requirements</span></span>

|<span data-ttu-id="ad969-167">Requisito</span><span class="sxs-lookup"><span data-stu-id="ad969-167">Requirement</span></span>| <span data-ttu-id="ad969-168">Valor</span><span class="sxs-lookup"><span data-stu-id="ad969-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad969-169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ad969-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad969-170">1.0</span><span class="sxs-lookup"><span data-stu-id="ad969-170">1.0</span></span>|
|[<span data-ttu-id="ad969-171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ad969-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad969-172">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ad969-172">Compose or Read</span></span>|
