---
title: Namespace do Office – conjunto de requisitos 1.2
description: Office namespace disponíveis para os Outlook que usam o conjunto de requisitos da API de Caixa de Correio 1.2.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 4cd15d77d1c5d9b95152f038f3421c5838bfb84f
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590404"
---
# <a name="office-mailbox-requirement-set-12"></a><span data-ttu-id="43c8b-103">Office (conjunto de requisitos de caixa de correio 1.2)</span><span class="sxs-lookup"><span data-stu-id="43c8b-103">Office (Mailbox requirement set 1.2)</span></span>

<span data-ttu-id="43c8b-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="43c8b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="43c8b-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c8b-106">Requirements</span></span>

|<span data-ttu-id="43c8b-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c8b-107">Requirement</span></span>| <span data-ttu-id="43c8b-108">Valor</span><span class="sxs-lookup"><span data-stu-id="43c8b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c8b-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c8b-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="43c8b-110">1.1</span><span class="sxs-lookup"><span data-stu-id="43c8b-110">1.1</span></span>|
|[<span data-ttu-id="43c8b-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c8b-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="43c8b-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c8b-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="43c8b-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="43c8b-113">Properties</span></span>

| <span data-ttu-id="43c8b-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="43c8b-114">Property</span></span> | <span data-ttu-id="43c8b-115">Modos</span><span class="sxs-lookup"><span data-stu-id="43c8b-115">Modes</span></span> | <span data-ttu-id="43c8b-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="43c8b-116">Return type</span></span> | <span data-ttu-id="43c8b-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="43c8b-117">Minimum</span></span><br><span data-ttu-id="43c8b-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="43c8b-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="43c8b-119">context</span><span class="sxs-lookup"><span data-stu-id="43c8b-119">context</span></span>](office.context.md) | <span data-ttu-id="43c8b-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="43c8b-120">Compose</span></span><br><span data-ttu-id="43c8b-121">Ler</span><span class="sxs-lookup"><span data-stu-id="43c8b-121">Read</span></span> | [<span data-ttu-id="43c8b-122">Context</span><span class="sxs-lookup"><span data-stu-id="43c8b-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="43c8b-123">1.1</span><span class="sxs-lookup"><span data-stu-id="43c8b-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="43c8b-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="43c8b-124">Enumerations</span></span>

| <span data-ttu-id="43c8b-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="43c8b-125">Enumeration</span></span> | <span data-ttu-id="43c8b-126">Modos</span><span class="sxs-lookup"><span data-stu-id="43c8b-126">Modes</span></span> | <span data-ttu-id="43c8b-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="43c8b-127">Return type</span></span> | <span data-ttu-id="43c8b-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="43c8b-128">Minimum</span></span><br><span data-ttu-id="43c8b-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="43c8b-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="43c8b-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="43c8b-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="43c8b-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="43c8b-131">Compose</span></span><br><span data-ttu-id="43c8b-132">Ler</span><span class="sxs-lookup"><span data-stu-id="43c8b-132">Read</span></span> | <span data-ttu-id="43c8b-133">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="43c8b-133">String</span></span> | [<span data-ttu-id="43c8b-134">1.1</span><span class="sxs-lookup"><span data-stu-id="43c8b-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="43c8b-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="43c8b-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="43c8b-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="43c8b-136">Compose</span></span><br><span data-ttu-id="43c8b-137">Ler</span><span class="sxs-lookup"><span data-stu-id="43c8b-137">Read</span></span> | <span data-ttu-id="43c8b-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="43c8b-138">String</span></span> | [<span data-ttu-id="43c8b-139">1.1</span><span class="sxs-lookup"><span data-stu-id="43c8b-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="43c8b-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="43c8b-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="43c8b-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="43c8b-141">Compose</span></span><br><span data-ttu-id="43c8b-142">Ler</span><span class="sxs-lookup"><span data-stu-id="43c8b-142">Read</span></span> | <span data-ttu-id="43c8b-143">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="43c8b-143">String</span></span> | [<span data-ttu-id="43c8b-144">1.1</span><span class="sxs-lookup"><span data-stu-id="43c8b-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="43c8b-145">Namespaces</span><span class="sxs-lookup"><span data-stu-id="43c8b-145">Namespaces</span></span>

<span data-ttu-id="43c8b-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): inclui várias enumerações específicas Outlook, por exemplo, `ItemType` , , , , , e `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="43c8b-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="43c8b-147">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="43c8b-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="43c8b-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="43c8b-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="43c8b-149">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="43c8b-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="43c8b-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c8b-150">Type</span></span>

*   <span data-ttu-id="43c8b-151">String</span><span class="sxs-lookup"><span data-stu-id="43c8b-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="43c8b-152">Propriedades</span><span class="sxs-lookup"><span data-stu-id="43c8b-152">Properties</span></span>

|<span data-ttu-id="43c8b-153">Nome</span><span class="sxs-lookup"><span data-stu-id="43c8b-153">Name</span></span>| <span data-ttu-id="43c8b-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c8b-154">Type</span></span>| <span data-ttu-id="43c8b-155">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c8b-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="43c8b-156">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="43c8b-156">String</span></span>|<span data-ttu-id="43c8b-157">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="43c8b-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="43c8b-158">String</span><span class="sxs-lookup"><span data-stu-id="43c8b-158">String</span></span>|<span data-ttu-id="43c8b-159">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="43c8b-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43c8b-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c8b-160">Requirements</span></span>

|<span data-ttu-id="43c8b-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c8b-161">Requirement</span></span>| <span data-ttu-id="43c8b-162">Valor</span><span class="sxs-lookup"><span data-stu-id="43c8b-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c8b-163">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c8b-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="43c8b-164">1.1</span><span class="sxs-lookup"><span data-stu-id="43c8b-164">1.1</span></span>|
|[<span data-ttu-id="43c8b-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c8b-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="43c8b-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c8b-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="43c8b-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="43c8b-167">CoercionType: String</span></span>

<span data-ttu-id="43c8b-168">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="43c8b-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="43c8b-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c8b-169">Type</span></span>

*   <span data-ttu-id="43c8b-170">String</span><span class="sxs-lookup"><span data-stu-id="43c8b-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="43c8b-171">Propriedades</span><span class="sxs-lookup"><span data-stu-id="43c8b-171">Properties</span></span>

|<span data-ttu-id="43c8b-172">Nome</span><span class="sxs-lookup"><span data-stu-id="43c8b-172">Name</span></span>| <span data-ttu-id="43c8b-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c8b-173">Type</span></span>| <span data-ttu-id="43c8b-174">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c8b-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="43c8b-175">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="43c8b-175">String</span></span>|<span data-ttu-id="43c8b-176">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="43c8b-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="43c8b-177">String</span><span class="sxs-lookup"><span data-stu-id="43c8b-177">String</span></span>|<span data-ttu-id="43c8b-178">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="43c8b-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43c8b-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c8b-179">Requirements</span></span>

|<span data-ttu-id="43c8b-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c8b-180">Requirement</span></span>| <span data-ttu-id="43c8b-181">Valor</span><span class="sxs-lookup"><span data-stu-id="43c8b-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c8b-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c8b-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="43c8b-183">1.1</span><span class="sxs-lookup"><span data-stu-id="43c8b-183">1.1</span></span>|
|[<span data-ttu-id="43c8b-184">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c8b-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="43c8b-185">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c8b-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="43c8b-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="43c8b-186">SourceProperty: String</span></span>

<span data-ttu-id="43c8b-187">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="43c8b-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="43c8b-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c8b-188">Type</span></span>

*   <span data-ttu-id="43c8b-189">String</span><span class="sxs-lookup"><span data-stu-id="43c8b-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="43c8b-190">Propriedades</span><span class="sxs-lookup"><span data-stu-id="43c8b-190">Properties</span></span>

|<span data-ttu-id="43c8b-191">Nome</span><span class="sxs-lookup"><span data-stu-id="43c8b-191">Name</span></span>| <span data-ttu-id="43c8b-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="43c8b-192">Type</span></span>| <span data-ttu-id="43c8b-193">Descrição</span><span class="sxs-lookup"><span data-stu-id="43c8b-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="43c8b-194">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="43c8b-194">String</span></span>|<span data-ttu-id="43c8b-195">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="43c8b-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="43c8b-196">String</span><span class="sxs-lookup"><span data-stu-id="43c8b-196">String</span></span>|<span data-ttu-id="43c8b-197">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="43c8b-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="43c8b-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="43c8b-198">Requirements</span></span>

|<span data-ttu-id="43c8b-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="43c8b-199">Requirement</span></span>| <span data-ttu-id="43c8b-200">Valor</span><span class="sxs-lookup"><span data-stu-id="43c8b-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="43c8b-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="43c8b-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="43c8b-202">1.1</span><span class="sxs-lookup"><span data-stu-id="43c8b-202">1.1</span></span>|
|[<span data-ttu-id="43c8b-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="43c8b-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="43c8b-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="43c8b-204">Compose or Read</span></span>|
