---
title: Namespace do Office – conjunto de requisitos 1,1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: eda5e1fb5f2c11ae91e4a1479892c36ec23e1897
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871994"
---
# <a name="office"></a><span data-ttu-id="70c54-102">Office</span><span class="sxs-lookup"><span data-stu-id="70c54-102">Office</span></span>

<span data-ttu-id="70c54-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="70c54-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="70c54-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="70c54-105">Requirements</span></span>

|<span data-ttu-id="70c54-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="70c54-106">Requirement</span></span>| <span data-ttu-id="70c54-107">Valor</span><span class="sxs-lookup"><span data-stu-id="70c54-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="70c54-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="70c54-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="70c54-109">1.0</span><span class="sxs-lookup"><span data-stu-id="70c54-109">1.0</span></span>|
|[<span data-ttu-id="70c54-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="70c54-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="70c54-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="70c54-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="70c54-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="70c54-112">Namespaces</span></span>

<span data-ttu-id="70c54-113">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="70c54-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="70c54-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="70c54-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="70c54-115">Membros</span><span class="sxs-lookup"><span data-stu-id="70c54-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="70c54-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="70c54-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="70c54-117">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="70c54-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="70c54-118">Tipo</span><span class="sxs-lookup"><span data-stu-id="70c54-118">Type</span></span>

*   <span data-ttu-id="70c54-119">String</span><span class="sxs-lookup"><span data-stu-id="70c54-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="70c54-120">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="70c54-120">Properties:</span></span>

|<span data-ttu-id="70c54-121">Nome</span><span class="sxs-lookup"><span data-stu-id="70c54-121">Name</span></span>| <span data-ttu-id="70c54-122">Tipo</span><span class="sxs-lookup"><span data-stu-id="70c54-122">Type</span></span>| <span data-ttu-id="70c54-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="70c54-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="70c54-124">String</span><span class="sxs-lookup"><span data-stu-id="70c54-124">String</span></span>|<span data-ttu-id="70c54-125">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="70c54-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="70c54-126">String</span><span class="sxs-lookup"><span data-stu-id="70c54-126">String</span></span>|<span data-ttu-id="70c54-127">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="70c54-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="70c54-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="70c54-128">Requirements</span></span>

|<span data-ttu-id="70c54-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="70c54-129">Requirement</span></span>| <span data-ttu-id="70c54-130">Valor</span><span class="sxs-lookup"><span data-stu-id="70c54-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="70c54-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="70c54-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="70c54-132">1.0</span><span class="sxs-lookup"><span data-stu-id="70c54-132">1.0</span></span>|
|[<span data-ttu-id="70c54-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="70c54-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="70c54-134">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="70c54-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="70c54-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="70c54-135">CoercionType :String</span></span>

<span data-ttu-id="70c54-136">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="70c54-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="70c54-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="70c54-137">Type</span></span>

*   <span data-ttu-id="70c54-138">String</span><span class="sxs-lookup"><span data-stu-id="70c54-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="70c54-139">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="70c54-139">Properties:</span></span>

|<span data-ttu-id="70c54-140">Nome</span><span class="sxs-lookup"><span data-stu-id="70c54-140">Name</span></span>| <span data-ttu-id="70c54-141">Tipo</span><span class="sxs-lookup"><span data-stu-id="70c54-141">Type</span></span>| <span data-ttu-id="70c54-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="70c54-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="70c54-143">String</span><span class="sxs-lookup"><span data-stu-id="70c54-143">String</span></span>|<span data-ttu-id="70c54-144">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="70c54-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="70c54-145">String</span><span class="sxs-lookup"><span data-stu-id="70c54-145">String</span></span>|<span data-ttu-id="70c54-146">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="70c54-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="70c54-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="70c54-147">Requirements</span></span>

|<span data-ttu-id="70c54-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="70c54-148">Requirement</span></span>| <span data-ttu-id="70c54-149">Valor</span><span class="sxs-lookup"><span data-stu-id="70c54-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="70c54-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="70c54-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="70c54-151">1.0</span><span class="sxs-lookup"><span data-stu-id="70c54-151">1.0</span></span>|
|[<span data-ttu-id="70c54-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="70c54-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="70c54-153">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="70c54-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="70c54-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="70c54-154">SourceProperty :String</span></span>

<span data-ttu-id="70c54-155">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="70c54-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="70c54-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="70c54-156">Type</span></span>

*   <span data-ttu-id="70c54-157">String</span><span class="sxs-lookup"><span data-stu-id="70c54-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="70c54-158">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="70c54-158">Properties:</span></span>

|<span data-ttu-id="70c54-159">Nome</span><span class="sxs-lookup"><span data-stu-id="70c54-159">Name</span></span>| <span data-ttu-id="70c54-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="70c54-160">Type</span></span>| <span data-ttu-id="70c54-161">Descrição</span><span class="sxs-lookup"><span data-stu-id="70c54-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="70c54-162">String</span><span class="sxs-lookup"><span data-stu-id="70c54-162">String</span></span>|<span data-ttu-id="70c54-163">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="70c54-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="70c54-164">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="70c54-164">String</span></span>|<span data-ttu-id="70c54-165">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="70c54-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="70c54-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="70c54-166">Requirements</span></span>

|<span data-ttu-id="70c54-167">Requisito</span><span class="sxs-lookup"><span data-stu-id="70c54-167">Requirement</span></span>| <span data-ttu-id="70c54-168">Valor</span><span class="sxs-lookup"><span data-stu-id="70c54-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="70c54-169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="70c54-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="70c54-170">1.0</span><span class="sxs-lookup"><span data-stu-id="70c54-170">1.0</span></span>|
|[<span data-ttu-id="70c54-171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="70c54-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="70c54-172">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="70c54-172">Compose or Read</span></span>|
