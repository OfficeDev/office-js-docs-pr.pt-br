---
title: Namespace do Office – conjunto de requisitos 1.1
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: af2a48d5bc943d4f443c32777fefaf8ed4a30032
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457429"
---
# <a name="office"></a><span data-ttu-id="01765-102">Office</span><span class="sxs-lookup"><span data-stu-id="01765-102">Office</span></span>

<span data-ttu-id="01765-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="01765-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="01765-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="01765-105">Requirements</span></span>

|<span data-ttu-id="01765-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="01765-106">Requirement</span></span>| <span data-ttu-id="01765-107">Valor</span><span class="sxs-lookup"><span data-stu-id="01765-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="01765-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="01765-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="01765-109">1.0</span><span class="sxs-lookup"><span data-stu-id="01765-109">1.0</span></span>|
|[<span data-ttu-id="01765-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="01765-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="01765-111">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="01765-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="01765-112">Namespaces</span><span class="sxs-lookup"><span data-stu-id="01765-112">Namespaces</span></span>

<span data-ttu-id="01765-113">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="01765-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="01765-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="01765-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="01765-115">Membros</span><span class="sxs-lookup"><span data-stu-id="01765-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="01765-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="01765-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="01765-117">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="01765-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="01765-118">Tipo:</span><span class="sxs-lookup"><span data-stu-id="01765-118">Type:</span></span>

*   <span data-ttu-id="01765-119">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="01765-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="01765-120">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="01765-120">Properties:</span></span>

|<span data-ttu-id="01765-121">Nome</span><span class="sxs-lookup"><span data-stu-id="01765-121">Name</span></span>| <span data-ttu-id="01765-122">Tipo</span><span class="sxs-lookup"><span data-stu-id="01765-122">Type</span></span>| <span data-ttu-id="01765-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="01765-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="01765-124">String</span><span class="sxs-lookup"><span data-stu-id="01765-124">String</span></span>|<span data-ttu-id="01765-125">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="01765-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="01765-126">String</span><span class="sxs-lookup"><span data-stu-id="01765-126">String</span></span>|<span data-ttu-id="01765-127">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="01765-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="01765-128">Requisitos</span><span class="sxs-lookup"><span data-stu-id="01765-128">Requirements</span></span>

|<span data-ttu-id="01765-129">Requisito</span><span class="sxs-lookup"><span data-stu-id="01765-129">Requirement</span></span>| <span data-ttu-id="01765-130">Valor</span><span class="sxs-lookup"><span data-stu-id="01765-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="01765-131">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="01765-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="01765-132">1.0</span><span class="sxs-lookup"><span data-stu-id="01765-132">1.0</span></span>|
|[<span data-ttu-id="01765-133">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="01765-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="01765-134">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="01765-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="01765-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="01765-135">CoercionType :String</span></span>

<span data-ttu-id="01765-136">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="01765-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="01765-137">Tipo:</span><span class="sxs-lookup"><span data-stu-id="01765-137">Type:</span></span>

*   <span data-ttu-id="01765-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="01765-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="01765-139">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="01765-139">Properties:</span></span>

|<span data-ttu-id="01765-140">Nome</span><span class="sxs-lookup"><span data-stu-id="01765-140">Name</span></span>| <span data-ttu-id="01765-141">Tipo</span><span class="sxs-lookup"><span data-stu-id="01765-141">Type</span></span>| <span data-ttu-id="01765-142">Descrição</span><span class="sxs-lookup"><span data-stu-id="01765-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="01765-143">String</span><span class="sxs-lookup"><span data-stu-id="01765-143">String</span></span>|<span data-ttu-id="01765-144">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="01765-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="01765-145">String</span><span class="sxs-lookup"><span data-stu-id="01765-145">String</span></span>|<span data-ttu-id="01765-146">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="01765-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="01765-147">Requisitos</span><span class="sxs-lookup"><span data-stu-id="01765-147">Requirements</span></span>

|<span data-ttu-id="01765-148">Requisito</span><span class="sxs-lookup"><span data-stu-id="01765-148">Requirement</span></span>| <span data-ttu-id="01765-149">Valor</span><span class="sxs-lookup"><span data-stu-id="01765-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="01765-150">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="01765-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="01765-151">1.0</span><span class="sxs-lookup"><span data-stu-id="01765-151">1.0</span></span>|
|[<span data-ttu-id="01765-152">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="01765-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="01765-153">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="01765-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="01765-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="01765-154">SourceProperty :String</span></span>

<span data-ttu-id="01765-155">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="01765-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="01765-156">Tipo:</span><span class="sxs-lookup"><span data-stu-id="01765-156">Type:</span></span>

*   <span data-ttu-id="01765-157">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="01765-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="01765-158">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="01765-158">Properties:</span></span>

|<span data-ttu-id="01765-159">Nome</span><span class="sxs-lookup"><span data-stu-id="01765-159">Name</span></span>| <span data-ttu-id="01765-160">Tipo</span><span class="sxs-lookup"><span data-stu-id="01765-160">Type</span></span>| <span data-ttu-id="01765-161">Descrição</span><span class="sxs-lookup"><span data-stu-id="01765-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="01765-162">String</span><span class="sxs-lookup"><span data-stu-id="01765-162">String</span></span>|<span data-ttu-id="01765-163">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="01765-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="01765-164">String</span><span class="sxs-lookup"><span data-stu-id="01765-164">String</span></span>|<span data-ttu-id="01765-165">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="01765-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="01765-166">Requisitos</span><span class="sxs-lookup"><span data-stu-id="01765-166">Requirements</span></span>

|<span data-ttu-id="01765-167">Requisito</span><span class="sxs-lookup"><span data-stu-id="01765-167">Requirement</span></span>| <span data-ttu-id="01765-168">Valor</span><span class="sxs-lookup"><span data-stu-id="01765-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="01765-169">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="01765-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="01765-170">1.0</span><span class="sxs-lookup"><span data-stu-id="01765-170">1.0</span></span>|
|[<span data-ttu-id="01765-171">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="01765-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="01765-172">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="01765-172">Compose or read</span></span>|