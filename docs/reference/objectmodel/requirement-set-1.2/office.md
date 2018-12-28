---
title: Namespace do Office – conjunto de requisitos 1.2
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 0dfd20d8c45f7630f388cb7ecde9ffb2a0aa12fb
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433233"
---
# <a name="office"></a><span data-ttu-id="db843-102">Office</span><span class="sxs-lookup"><span data-stu-id="db843-102">Office</span></span>

<span data-ttu-id="db843-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="db843-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="db843-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="db843-105">Requirements</span></span>

|<span data-ttu-id="db843-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="db843-106">Requirement</span></span>| <span data-ttu-id="db843-107">Valor</span><span class="sxs-lookup"><span data-stu-id="db843-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="db843-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="db843-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="db843-109">1.0</span><span class="sxs-lookup"><span data-stu-id="db843-109">1.0</span></span>|
|[<span data-ttu-id="db843-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="db843-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="db843-111">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="db843-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="db843-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="db843-112">Namespaces</span></span>

<span data-ttu-id="db843-113">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="db843-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="db843-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="db843-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="db843-115">Membros</span><span class="sxs-lookup"><span data-stu-id="db843-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="db843-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="db843-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="db843-117">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="db843-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="db843-118">Tipo:</span><span class="sxs-lookup"><span data-stu-id="db843-118">Type:</span></span>

*   <span data-ttu-id="db843-119">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="db843-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="db843-120">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="db843-120">Properties:</span></span>

|<span data-ttu-id="db843-121">Nome</span><span class="sxs-lookup"><span data-stu-id="db843-121">Name</span></span>| <span data-ttu-id="db843-122">Tipo</span><span class="sxs-lookup"><span data-stu-id="db843-122">Type</span></span>| <span data-ttu-id="db843-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="db843-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="db843-124">String</span><span class="sxs-lookup"><span data-stu-id="db843-124">String</span></span>|<span data-ttu-id="db843-125">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="db843-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="db843-126">String</span><span class="sxs-lookup"><span data-stu-id="db843-126">String</span></span>|<span data-ttu-id="db843-127">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="db843-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="db843-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="db843-128">Requirements</span></span>

|<span data-ttu-id="db843-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="db843-129">Requirement</span></span>| <span data-ttu-id="db843-130">Valor</span><span class="sxs-lookup"><span data-stu-id="db843-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="db843-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="db843-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="db843-132">1.0</span><span class="sxs-lookup"><span data-stu-id="db843-132">1.0</span></span>|
|[<span data-ttu-id="db843-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="db843-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="db843-134">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="db843-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="db843-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="db843-135">CoercionType :String</span></span>

<span data-ttu-id="db843-136">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="db843-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="db843-137">Tipo:</span><span class="sxs-lookup"><span data-stu-id="db843-137">Type:</span></span>

*   <span data-ttu-id="db843-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="db843-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="db843-139">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="db843-139">Properties:</span></span>

|<span data-ttu-id="db843-140">Nome</span><span class="sxs-lookup"><span data-stu-id="db843-140">Name</span></span>| <span data-ttu-id="db843-141">Tipo</span><span class="sxs-lookup"><span data-stu-id="db843-141">Type</span></span>| <span data-ttu-id="db843-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="db843-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="db843-143">String</span><span class="sxs-lookup"><span data-stu-id="db843-143">String</span></span>|<span data-ttu-id="db843-144">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="db843-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="db843-145">String</span><span class="sxs-lookup"><span data-stu-id="db843-145">String</span></span>|<span data-ttu-id="db843-146">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="db843-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="db843-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="db843-147">Requirements</span></span>

|<span data-ttu-id="db843-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="db843-148">Requirement</span></span>| <span data-ttu-id="db843-149">Valor</span><span class="sxs-lookup"><span data-stu-id="db843-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="db843-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="db843-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="db843-151">1.0</span><span class="sxs-lookup"><span data-stu-id="db843-151">1.0</span></span>|
|[<span data-ttu-id="db843-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="db843-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="db843-153">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="db843-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="db843-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="db843-154">SourceProperty :String</span></span>

<span data-ttu-id="db843-155">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="db843-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="db843-156">Tipo:</span><span class="sxs-lookup"><span data-stu-id="db843-156">Type:</span></span>

*   <span data-ttu-id="db843-157">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="db843-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="db843-158">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="db843-158">Properties:</span></span>

|<span data-ttu-id="db843-159">Nome</span><span class="sxs-lookup"><span data-stu-id="db843-159">Name</span></span>| <span data-ttu-id="db843-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="db843-160">Type</span></span>| <span data-ttu-id="db843-161">Descrição</span><span class="sxs-lookup"><span data-stu-id="db843-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="db843-162">String</span><span class="sxs-lookup"><span data-stu-id="db843-162">String</span></span>|<span data-ttu-id="db843-163">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="db843-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="db843-164">String</span><span class="sxs-lookup"><span data-stu-id="db843-164">String</span></span>|<span data-ttu-id="db843-165">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="db843-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="db843-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="db843-166">Requirements</span></span>

|<span data-ttu-id="db843-167">Requisito</span><span class="sxs-lookup"><span data-stu-id="db843-167">Requirement</span></span>| <span data-ttu-id="db843-168">Valor</span><span class="sxs-lookup"><span data-stu-id="db843-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="db843-169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="db843-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="db843-170">1.0</span><span class="sxs-lookup"><span data-stu-id="db843-170">1.0</span></span>|
|[<span data-ttu-id="db843-171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="db843-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="db843-172">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="db843-172">Compose or read</span></span>|