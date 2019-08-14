---
title: Namespace do Office – conjunto de requisitos 1.2
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 5a8431580fce2a98f2076ef3df151f08d5435d54
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395642"
---
# <a name="office"></a><span data-ttu-id="c1737-102">Office</span><span class="sxs-lookup"><span data-stu-id="c1737-102">Office</span></span>

<span data-ttu-id="c1737-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c1737-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1737-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c1737-105">Requirements</span></span>

|<span data-ttu-id="c1737-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="c1737-106">Requirement</span></span>| <span data-ttu-id="c1737-107">Valor</span><span class="sxs-lookup"><span data-stu-id="c1737-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1737-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c1737-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1737-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c1737-109">1.0</span></span>|
|[<span data-ttu-id="c1737-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c1737-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c1737-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c1737-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c1737-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="c1737-112">Members and methods</span></span>

| <span data-ttu-id="c1737-113">Membro</span><span class="sxs-lookup"><span data-stu-id="c1737-113">Member</span></span> | <span data-ttu-id="c1737-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="c1737-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c1737-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c1737-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c1737-116">Membro</span><span class="sxs-lookup"><span data-stu-id="c1737-116">Member</span></span> |
| [<span data-ttu-id="c1737-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c1737-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c1737-118">Membro</span><span class="sxs-lookup"><span data-stu-id="c1737-118">Member</span></span> |
| [<span data-ttu-id="c1737-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c1737-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c1737-120">Membro</span><span class="sxs-lookup"><span data-stu-id="c1737-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c1737-121">Namespaces</span><span class="sxs-lookup"><span data-stu-id="c1737-121">Namespaces</span></span>

<span data-ttu-id="c1737-122">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c1737-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c1737-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): inclui um número de enumerações, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="c1737-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="c1737-124">Members</span><span class="sxs-lookup"><span data-stu-id="c1737-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="c1737-125">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c1737-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="c1737-126">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="c1737-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c1737-127">Tipo</span><span class="sxs-lookup"><span data-stu-id="c1737-127">Type</span></span>

*   <span data-ttu-id="c1737-128">String</span><span class="sxs-lookup"><span data-stu-id="c1737-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c1737-129">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="c1737-129">Properties:</span></span>

|<span data-ttu-id="c1737-130">Nome</span><span class="sxs-lookup"><span data-stu-id="c1737-130">Name</span></span>| <span data-ttu-id="c1737-131">Tipo</span><span class="sxs-lookup"><span data-stu-id="c1737-131">Type</span></span>| <span data-ttu-id="c1737-132">Descrição</span><span class="sxs-lookup"><span data-stu-id="c1737-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c1737-133">String</span><span class="sxs-lookup"><span data-stu-id="c1737-133">String</span></span>|<span data-ttu-id="c1737-134">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="c1737-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c1737-135">String</span><span class="sxs-lookup"><span data-stu-id="c1737-135">String</span></span>|<span data-ttu-id="c1737-136">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="c1737-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c1737-137">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c1737-137">Requirements</span></span>

|<span data-ttu-id="c1737-138">Requisito</span><span class="sxs-lookup"><span data-stu-id="c1737-138">Requirement</span></span>| <span data-ttu-id="c1737-139">Valor</span><span class="sxs-lookup"><span data-stu-id="c1737-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1737-140">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c1737-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1737-141">1.0</span><span class="sxs-lookup"><span data-stu-id="c1737-141">1.0</span></span>|
|[<span data-ttu-id="c1737-142">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c1737-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c1737-143">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c1737-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="c1737-144">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c1737-144">CoercionType: String</span></span>

<span data-ttu-id="c1737-145">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="c1737-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c1737-146">Tipo</span><span class="sxs-lookup"><span data-stu-id="c1737-146">Type</span></span>

*   <span data-ttu-id="c1737-147">String</span><span class="sxs-lookup"><span data-stu-id="c1737-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c1737-148">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="c1737-148">Properties:</span></span>

|<span data-ttu-id="c1737-149">Nome</span><span class="sxs-lookup"><span data-stu-id="c1737-149">Name</span></span>| <span data-ttu-id="c1737-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="c1737-150">Type</span></span>| <span data-ttu-id="c1737-151">Descrição</span><span class="sxs-lookup"><span data-stu-id="c1737-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c1737-152">String</span><span class="sxs-lookup"><span data-stu-id="c1737-152">String</span></span>|<span data-ttu-id="c1737-153">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="c1737-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c1737-154">String</span><span class="sxs-lookup"><span data-stu-id="c1737-154">String</span></span>|<span data-ttu-id="c1737-155">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="c1737-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c1737-156">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c1737-156">Requirements</span></span>

|<span data-ttu-id="c1737-157">Requisito</span><span class="sxs-lookup"><span data-stu-id="c1737-157">Requirement</span></span>| <span data-ttu-id="c1737-158">Valor</span><span class="sxs-lookup"><span data-stu-id="c1737-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1737-159">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c1737-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1737-160">1.0</span><span class="sxs-lookup"><span data-stu-id="c1737-160">1.0</span></span>|
|[<span data-ttu-id="c1737-161">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c1737-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c1737-162">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c1737-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="c1737-163">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c1737-163">SourceProperty: String</span></span>

<span data-ttu-id="c1737-164">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="c1737-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c1737-165">Tipo</span><span class="sxs-lookup"><span data-stu-id="c1737-165">Type</span></span>

*   <span data-ttu-id="c1737-166">String</span><span class="sxs-lookup"><span data-stu-id="c1737-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c1737-167">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="c1737-167">Properties:</span></span>

|<span data-ttu-id="c1737-168">Nome</span><span class="sxs-lookup"><span data-stu-id="c1737-168">Name</span></span>| <span data-ttu-id="c1737-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="c1737-169">Type</span></span>| <span data-ttu-id="c1737-170">Descrição</span><span class="sxs-lookup"><span data-stu-id="c1737-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c1737-171">String</span><span class="sxs-lookup"><span data-stu-id="c1737-171">String</span></span>|<span data-ttu-id="c1737-172">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="c1737-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c1737-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c1737-173">String</span></span>|<span data-ttu-id="c1737-174">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="c1737-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c1737-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c1737-175">Requirements</span></span>

|<span data-ttu-id="c1737-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="c1737-176">Requirement</span></span>| <span data-ttu-id="c1737-177">Valor</span><span class="sxs-lookup"><span data-stu-id="c1737-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1737-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c1737-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1737-179">1.0</span><span class="sxs-lookup"><span data-stu-id="c1737-179">1.0</span></span>|
|[<span data-ttu-id="c1737-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c1737-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c1737-181">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c1737-181">Compose or Read</span></span>|
