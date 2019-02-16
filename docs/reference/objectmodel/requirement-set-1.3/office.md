---
title: Namespace do Office – conjunto de requisitos versão 1.3
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: ad08895719d809436216d2f0bb455260dbca3b1e
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067885"
---
# <a name="office"></a><span data-ttu-id="6ef5c-102">Office</span><span class="sxs-lookup"><span data-stu-id="6ef5c-102">Office</span></span>

<span data-ttu-id="6ef5c-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="6ef5c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6ef5c-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6ef5c-105">Requirements</span></span>

|<span data-ttu-id="6ef5c-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="6ef5c-106">Requirement</span></span>| <span data-ttu-id="6ef5c-107">Valor</span><span class="sxs-lookup"><span data-stu-id="6ef5c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6ef5c-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6ef5c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6ef5c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="6ef5c-109">1.0</span></span>|
|[<span data-ttu-id="6ef5c-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6ef5c-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6ef5c-111">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6ef5c-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="6ef5c-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="6ef5c-112">Namespaces</span></span>

<span data-ttu-id="6ef5c-113">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="6ef5c-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="6ef5c-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="6ef5c-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="6ef5c-115">Membros</span><span class="sxs-lookup"><span data-stu-id="6ef5c-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="6ef5c-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="6ef5c-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="6ef5c-117">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="6ef5c-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="6ef5c-118">Tipo</span><span class="sxs-lookup"><span data-stu-id="6ef5c-118">Type</span></span>

*   <span data-ttu-id="6ef5c-119">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6ef5c-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6ef5c-120">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="6ef5c-120">Properties:</span></span>

|<span data-ttu-id="6ef5c-121">Nome</span><span class="sxs-lookup"><span data-stu-id="6ef5c-121">Name</span></span>| <span data-ttu-id="6ef5c-122">Tipo</span><span class="sxs-lookup"><span data-stu-id="6ef5c-122">Type</span></span>| <span data-ttu-id="6ef5c-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="6ef5c-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="6ef5c-124">String</span><span class="sxs-lookup"><span data-stu-id="6ef5c-124">String</span></span>|<span data-ttu-id="6ef5c-125">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="6ef5c-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="6ef5c-126">String</span><span class="sxs-lookup"><span data-stu-id="6ef5c-126">String</span></span>|<span data-ttu-id="6ef5c-127">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="6ef5c-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6ef5c-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6ef5c-128">Requirements</span></span>

|<span data-ttu-id="6ef5c-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="6ef5c-129">Requirement</span></span>| <span data-ttu-id="6ef5c-130">Valor</span><span class="sxs-lookup"><span data-stu-id="6ef5c-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="6ef5c-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6ef5c-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6ef5c-132">1.0</span><span class="sxs-lookup"><span data-stu-id="6ef5c-132">1.0</span></span>|
|[<span data-ttu-id="6ef5c-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6ef5c-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6ef5c-134">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6ef5c-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="6ef5c-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="6ef5c-135">CoercionType :String</span></span>

<span data-ttu-id="6ef5c-136">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="6ef5c-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6ef5c-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="6ef5c-137">Type</span></span>

*   <span data-ttu-id="6ef5c-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6ef5c-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6ef5c-139">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="6ef5c-139">Properties:</span></span>

|<span data-ttu-id="6ef5c-140">Nome</span><span class="sxs-lookup"><span data-stu-id="6ef5c-140">Name</span></span>| <span data-ttu-id="6ef5c-141">Tipo</span><span class="sxs-lookup"><span data-stu-id="6ef5c-141">Type</span></span>| <span data-ttu-id="6ef5c-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="6ef5c-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="6ef5c-143">String</span><span class="sxs-lookup"><span data-stu-id="6ef5c-143">String</span></span>|<span data-ttu-id="6ef5c-144">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="6ef5c-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="6ef5c-145">String</span><span class="sxs-lookup"><span data-stu-id="6ef5c-145">String</span></span>|<span data-ttu-id="6ef5c-146">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="6ef5c-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6ef5c-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6ef5c-147">Requirements</span></span>

|<span data-ttu-id="6ef5c-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="6ef5c-148">Requirement</span></span>| <span data-ttu-id="6ef5c-149">Valor</span><span class="sxs-lookup"><span data-stu-id="6ef5c-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="6ef5c-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6ef5c-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6ef5c-151">1.0</span><span class="sxs-lookup"><span data-stu-id="6ef5c-151">1.0</span></span>|
|[<span data-ttu-id="6ef5c-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6ef5c-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6ef5c-153">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6ef5c-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="6ef5c-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="6ef5c-154">SourceProperty :String</span></span>

<span data-ttu-id="6ef5c-155">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="6ef5c-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6ef5c-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="6ef5c-156">Type</span></span>

*   <span data-ttu-id="6ef5c-157">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6ef5c-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6ef5c-158">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="6ef5c-158">Properties:</span></span>

|<span data-ttu-id="6ef5c-159">Nome</span><span class="sxs-lookup"><span data-stu-id="6ef5c-159">Name</span></span>| <span data-ttu-id="6ef5c-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="6ef5c-160">Type</span></span>| <span data-ttu-id="6ef5c-161">Descrição</span><span class="sxs-lookup"><span data-stu-id="6ef5c-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="6ef5c-162">String</span><span class="sxs-lookup"><span data-stu-id="6ef5c-162">String</span></span>|<span data-ttu-id="6ef5c-163">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="6ef5c-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="6ef5c-164">String</span><span class="sxs-lookup"><span data-stu-id="6ef5c-164">String</span></span>|<span data-ttu-id="6ef5c-165">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="6ef5c-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6ef5c-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6ef5c-166">Requirements</span></span>

|<span data-ttu-id="6ef5c-167">Requisito</span><span class="sxs-lookup"><span data-stu-id="6ef5c-167">Requirement</span></span>| <span data-ttu-id="6ef5c-168">Valor</span><span class="sxs-lookup"><span data-stu-id="6ef5c-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="6ef5c-169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6ef5c-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6ef5c-170">1.0</span><span class="sxs-lookup"><span data-stu-id="6ef5c-170">1.0</span></span>|
|[<span data-ttu-id="6ef5c-171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6ef5c-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6ef5c-172">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="6ef5c-172">Compose or Read</span></span>|
