---
title: Namespace do Office – conjunto de requisitos versão 1.3
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 9a0f06cbe286f6479ac9244d5ad5bde43ab6b5b6
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457793"
---
# <a name="office"></a><span data-ttu-id="5a486-102">Office</span><span class="sxs-lookup"><span data-stu-id="5a486-102">Office</span></span>

<span data-ttu-id="5a486-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="5a486-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5a486-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5a486-105">Requirements</span></span>

|<span data-ttu-id="5a486-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="5a486-106">Requirement</span></span>| <span data-ttu-id="5a486-107">Valor</span><span class="sxs-lookup"><span data-stu-id="5a486-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="5a486-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5a486-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5a486-109">1.0</span><span class="sxs-lookup"><span data-stu-id="5a486-109">1.0</span></span>|
|[<span data-ttu-id="5a486-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5a486-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5a486-111">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="5a486-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="5a486-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="5a486-112">Namespaces</span></span>

<span data-ttu-id="5a486-113">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="5a486-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="5a486-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="5a486-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="5a486-115">Membros</span><span class="sxs-lookup"><span data-stu-id="5a486-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="5a486-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="5a486-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="5a486-117">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="5a486-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5a486-118">Tipo:</span><span class="sxs-lookup"><span data-stu-id="5a486-118">Type:</span></span>

*   <span data-ttu-id="5a486-119">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5a486-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5a486-120">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5a486-120">Properties:</span></span>

|<span data-ttu-id="5a486-121">Nome</span><span class="sxs-lookup"><span data-stu-id="5a486-121">Name</span></span>| <span data-ttu-id="5a486-122">Tipo</span><span class="sxs-lookup"><span data-stu-id="5a486-122">Type</span></span>| <span data-ttu-id="5a486-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="5a486-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5a486-124">String</span><span class="sxs-lookup"><span data-stu-id="5a486-124">String</span></span>|<span data-ttu-id="5a486-125">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="5a486-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5a486-126">String</span><span class="sxs-lookup"><span data-stu-id="5a486-126">String</span></span>|<span data-ttu-id="5a486-127">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="5a486-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5a486-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5a486-128">Requirements</span></span>

|<span data-ttu-id="5a486-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="5a486-129">Requirement</span></span>| <span data-ttu-id="5a486-130">Valor</span><span class="sxs-lookup"><span data-stu-id="5a486-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="5a486-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5a486-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5a486-132">1.0</span><span class="sxs-lookup"><span data-stu-id="5a486-132">1.0</span></span>|
|[<span data-ttu-id="5a486-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5a486-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5a486-134">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5a486-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="5a486-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="5a486-135">CoercionType :String</span></span>

<span data-ttu-id="5a486-136">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="5a486-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5a486-137">Tipo:</span><span class="sxs-lookup"><span data-stu-id="5a486-137">Type:</span></span>

*   <span data-ttu-id="5a486-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5a486-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5a486-139">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5a486-139">Properties:</span></span>

|<span data-ttu-id="5a486-140">Nome</span><span class="sxs-lookup"><span data-stu-id="5a486-140">Name</span></span>| <span data-ttu-id="5a486-141">Tipo</span><span class="sxs-lookup"><span data-stu-id="5a486-141">Type</span></span>| <span data-ttu-id="5a486-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="5a486-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5a486-143">String</span><span class="sxs-lookup"><span data-stu-id="5a486-143">String</span></span>|<span data-ttu-id="5a486-144">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="5a486-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5a486-145">String</span><span class="sxs-lookup"><span data-stu-id="5a486-145">String</span></span>|<span data-ttu-id="5a486-146">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="5a486-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5a486-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5a486-147">Requirements</span></span>

|<span data-ttu-id="5a486-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="5a486-148">Requirement</span></span>| <span data-ttu-id="5a486-149">Valor</span><span class="sxs-lookup"><span data-stu-id="5a486-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="5a486-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5a486-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5a486-151">1.0</span><span class="sxs-lookup"><span data-stu-id="5a486-151">1.0</span></span>|
|[<span data-ttu-id="5a486-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5a486-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5a486-153">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5a486-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="5a486-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="5a486-154">SourceProperty :String</span></span>

<span data-ttu-id="5a486-155">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="5a486-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5a486-156">Tipo:</span><span class="sxs-lookup"><span data-stu-id="5a486-156">Type:</span></span>

*   <span data-ttu-id="5a486-157">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5a486-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5a486-158">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5a486-158">Properties:</span></span>

|<span data-ttu-id="5a486-159">Nome</span><span class="sxs-lookup"><span data-stu-id="5a486-159">Name</span></span>| <span data-ttu-id="5a486-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="5a486-160">Type</span></span>| <span data-ttu-id="5a486-161">Descrição</span><span class="sxs-lookup"><span data-stu-id="5a486-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5a486-162">String</span><span class="sxs-lookup"><span data-stu-id="5a486-162">String</span></span>|<span data-ttu-id="5a486-163">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5a486-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5a486-164">String</span><span class="sxs-lookup"><span data-stu-id="5a486-164">String</span></span>|<span data-ttu-id="5a486-165">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5a486-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5a486-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5a486-166">Requirements</span></span>

|<span data-ttu-id="5a486-167">Requisito</span><span class="sxs-lookup"><span data-stu-id="5a486-167">Requirement</span></span>| <span data-ttu-id="5a486-168">Valor</span><span class="sxs-lookup"><span data-stu-id="5a486-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="5a486-169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5a486-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5a486-170">1.0</span><span class="sxs-lookup"><span data-stu-id="5a486-170">1.0</span></span>|
|[<span data-ttu-id="5a486-171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5a486-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5a486-172">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5a486-172">Compose or read</span></span>|