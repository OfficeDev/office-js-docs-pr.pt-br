---
title: Namespace do Office – conjunto de requisitos 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: eff7896214866e71b92a1c8a0c72a16e622873f3
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067857"
---
# <a name="office"></a><span data-ttu-id="5e1c2-102">Office</span><span class="sxs-lookup"><span data-stu-id="5e1c2-102">Office</span></span>

<span data-ttu-id="5e1c2-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="5e1c2-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e1c2-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5e1c2-105">Requirements</span></span>

|<span data-ttu-id="5e1c2-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="5e1c2-106">Requirement</span></span>| <span data-ttu-id="5e1c2-107">Valor</span><span class="sxs-lookup"><span data-stu-id="5e1c2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e1c2-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5e1c2-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e1c2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="5e1c2-109">1.0</span></span>|
|[<span data-ttu-id="5e1c2-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5e1c2-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5e1c2-111">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5e1c2-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="5e1c2-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="5e1c2-112">Namespaces</span></span>

<span data-ttu-id="5e1c2-113">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="5e1c2-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="5e1c2-114">[MailboxEnums](/javascript/api/outlook_1_2/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="5e1c2-114">[MailboxEnums](/javascript/api/outlook_1_2/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="5e1c2-115">Membros</span><span class="sxs-lookup"><span data-stu-id="5e1c2-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="5e1c2-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="5e1c2-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="5e1c2-117">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="5e1c2-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5e1c2-118">Tipo</span><span class="sxs-lookup"><span data-stu-id="5e1c2-118">Type</span></span>

*   <span data-ttu-id="5e1c2-119">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5e1c2-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5e1c2-120">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5e1c2-120">Properties:</span></span>

|<span data-ttu-id="5e1c2-121">Nome</span><span class="sxs-lookup"><span data-stu-id="5e1c2-121">Name</span></span>| <span data-ttu-id="5e1c2-122">Tipo</span><span class="sxs-lookup"><span data-stu-id="5e1c2-122">Type</span></span>| <span data-ttu-id="5e1c2-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="5e1c2-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5e1c2-124">String</span><span class="sxs-lookup"><span data-stu-id="5e1c2-124">String</span></span>|<span data-ttu-id="5e1c2-125">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="5e1c2-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5e1c2-126">String</span><span class="sxs-lookup"><span data-stu-id="5e1c2-126">String</span></span>|<span data-ttu-id="5e1c2-127">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="5e1c2-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5e1c2-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5e1c2-128">Requirements</span></span>

|<span data-ttu-id="5e1c2-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="5e1c2-129">Requirement</span></span>| <span data-ttu-id="5e1c2-130">Valor</span><span class="sxs-lookup"><span data-stu-id="5e1c2-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e1c2-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5e1c2-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e1c2-132">1.0</span><span class="sxs-lookup"><span data-stu-id="5e1c2-132">1.0</span></span>|
|[<span data-ttu-id="5e1c2-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5e1c2-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5e1c2-134">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5e1c2-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="5e1c2-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="5e1c2-135">CoercionType :String</span></span>

<span data-ttu-id="5e1c2-136">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="5e1c2-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5e1c2-137">Tipo</span><span class="sxs-lookup"><span data-stu-id="5e1c2-137">Type</span></span>

*   <span data-ttu-id="5e1c2-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5e1c2-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5e1c2-139">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5e1c2-139">Properties:</span></span>

|<span data-ttu-id="5e1c2-140">Nome</span><span class="sxs-lookup"><span data-stu-id="5e1c2-140">Name</span></span>| <span data-ttu-id="5e1c2-141">Tipo</span><span class="sxs-lookup"><span data-stu-id="5e1c2-141">Type</span></span>| <span data-ttu-id="5e1c2-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="5e1c2-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5e1c2-143">String</span><span class="sxs-lookup"><span data-stu-id="5e1c2-143">String</span></span>|<span data-ttu-id="5e1c2-144">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="5e1c2-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5e1c2-145">String</span><span class="sxs-lookup"><span data-stu-id="5e1c2-145">String</span></span>|<span data-ttu-id="5e1c2-146">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="5e1c2-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5e1c2-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5e1c2-147">Requirements</span></span>

|<span data-ttu-id="5e1c2-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="5e1c2-148">Requirement</span></span>| <span data-ttu-id="5e1c2-149">Valor</span><span class="sxs-lookup"><span data-stu-id="5e1c2-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e1c2-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5e1c2-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e1c2-151">1.0</span><span class="sxs-lookup"><span data-stu-id="5e1c2-151">1.0</span></span>|
|[<span data-ttu-id="5e1c2-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5e1c2-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5e1c2-153">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5e1c2-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="5e1c2-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="5e1c2-154">SourceProperty :String</span></span>

<span data-ttu-id="5e1c2-155">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="5e1c2-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5e1c2-156">Tipo</span><span class="sxs-lookup"><span data-stu-id="5e1c2-156">Type</span></span>

*   <span data-ttu-id="5e1c2-157">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5e1c2-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5e1c2-158">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5e1c2-158">Properties:</span></span>

|<span data-ttu-id="5e1c2-159">Nome</span><span class="sxs-lookup"><span data-stu-id="5e1c2-159">Name</span></span>| <span data-ttu-id="5e1c2-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="5e1c2-160">Type</span></span>| <span data-ttu-id="5e1c2-161">Descrição</span><span class="sxs-lookup"><span data-stu-id="5e1c2-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5e1c2-162">String</span><span class="sxs-lookup"><span data-stu-id="5e1c2-162">String</span></span>|<span data-ttu-id="5e1c2-163">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5e1c2-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5e1c2-164">String</span><span class="sxs-lookup"><span data-stu-id="5e1c2-164">String</span></span>|<span data-ttu-id="5e1c2-165">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5e1c2-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5e1c2-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5e1c2-166">Requirements</span></span>

|<span data-ttu-id="5e1c2-167">Requisito</span><span class="sxs-lookup"><span data-stu-id="5e1c2-167">Requirement</span></span>| <span data-ttu-id="5e1c2-168">Valor</span><span class="sxs-lookup"><span data-stu-id="5e1c2-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e1c2-169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5e1c2-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5e1c2-170">1.0</span><span class="sxs-lookup"><span data-stu-id="5e1c2-170">1.0</span></span>|
|[<span data-ttu-id="5e1c2-171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5e1c2-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5e1c2-172">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5e1c2-172">Compose or Read</span></span>|
