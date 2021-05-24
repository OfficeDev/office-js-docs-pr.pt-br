---
title: Office namespace - conjunto de requisitos 1.4
description: Office namespace disponíveis para Outlook que usam conjunto de requisitos da API de Caixa de Correio 1.4.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 0221ab09048719317c131f0204e2fc60c4f8f7d4
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591020"
---
# <a name="office-mailbox-requirement-set-14"></a><span data-ttu-id="6fe4e-103">Office (conjunto de requisitos de caixa de correio 1.4)</span><span class="sxs-lookup"><span data-stu-id="6fe4e-103">Office (Mailbox requirement set 1.4)</span></span>

<span data-ttu-id="6fe4e-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="6fe4e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fe4e-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6fe4e-106">Requirements</span></span>

|<span data-ttu-id="6fe4e-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="6fe4e-107">Requirement</span></span>| <span data-ttu-id="6fe4e-108">Valor</span><span class="sxs-lookup"><span data-stu-id="6fe4e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fe4e-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6fe4e-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fe4e-110">1.1</span><span class="sxs-lookup"><span data-stu-id="6fe4e-110">1.1</span></span>|
|[<span data-ttu-id="6fe4e-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6fe4e-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fe4e-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6fe4e-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="6fe4e-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="6fe4e-113">Properties</span></span>

| <span data-ttu-id="6fe4e-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="6fe4e-114">Property</span></span> | <span data-ttu-id="6fe4e-115">Modos</span><span class="sxs-lookup"><span data-stu-id="6fe4e-115">Modes</span></span> | <span data-ttu-id="6fe4e-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="6fe4e-116">Return type</span></span> | <span data-ttu-id="6fe4e-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="6fe4e-117">Minimum</span></span><br><span data-ttu-id="6fe4e-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="6fe4e-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6fe4e-119">context</span><span class="sxs-lookup"><span data-stu-id="6fe4e-119">context</span></span>](office.context.md) | <span data-ttu-id="6fe4e-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="6fe4e-120">Compose</span></span><br><span data-ttu-id="6fe4e-121">Ler</span><span class="sxs-lookup"><span data-stu-id="6fe4e-121">Read</span></span> | [<span data-ttu-id="6fe4e-122">Context</span><span class="sxs-lookup"><span data-stu-id="6fe4e-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="6fe4e-123">1.1</span><span class="sxs-lookup"><span data-stu-id="6fe4e-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="6fe4e-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="6fe4e-124">Enumerations</span></span>

| <span data-ttu-id="6fe4e-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="6fe4e-125">Enumeration</span></span> | <span data-ttu-id="6fe4e-126">Modos</span><span class="sxs-lookup"><span data-stu-id="6fe4e-126">Modes</span></span> | <span data-ttu-id="6fe4e-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="6fe4e-127">Return type</span></span> | <span data-ttu-id="6fe4e-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="6fe4e-128">Minimum</span></span><br><span data-ttu-id="6fe4e-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="6fe4e-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6fe4e-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="6fe4e-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="6fe4e-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="6fe4e-131">Compose</span></span><br><span data-ttu-id="6fe4e-132">Ler</span><span class="sxs-lookup"><span data-stu-id="6fe4e-132">Read</span></span> | <span data-ttu-id="6fe4e-133">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6fe4e-133">String</span></span> | [<span data-ttu-id="6fe4e-134">1.1</span><span class="sxs-lookup"><span data-stu-id="6fe4e-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6fe4e-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="6fe4e-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="6fe4e-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="6fe4e-136">Compose</span></span><br><span data-ttu-id="6fe4e-137">Ler</span><span class="sxs-lookup"><span data-stu-id="6fe4e-137">Read</span></span> | <span data-ttu-id="6fe4e-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6fe4e-138">String</span></span> | [<span data-ttu-id="6fe4e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="6fe4e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6fe4e-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="6fe4e-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="6fe4e-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="6fe4e-141">Compose</span></span><br><span data-ttu-id="6fe4e-142">Ler</span><span class="sxs-lookup"><span data-stu-id="6fe4e-142">Read</span></span> | <span data-ttu-id="6fe4e-143">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6fe4e-143">String</span></span> | [<span data-ttu-id="6fe4e-144">1.1</span><span class="sxs-lookup"><span data-stu-id="6fe4e-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="6fe4e-145">Namespaces</span><span class="sxs-lookup"><span data-stu-id="6fe4e-145">Namespaces</span></span>

<span data-ttu-id="6fe4e-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true): inclui várias enumerações específicas Outlook, por exemplo, `ItemType` , , , , , e `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="6fe4e-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="6fe4e-147">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="6fe4e-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="6fe4e-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="6fe4e-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="6fe4e-149">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="6fe4e-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="6fe4e-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="6fe4e-150">Type</span></span>

*   <span data-ttu-id="6fe4e-151">String</span><span class="sxs-lookup"><span data-stu-id="6fe4e-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6fe4e-152">Propriedades</span><span class="sxs-lookup"><span data-stu-id="6fe4e-152">Properties</span></span>

|<span data-ttu-id="6fe4e-153">Nome</span><span class="sxs-lookup"><span data-stu-id="6fe4e-153">Name</span></span>| <span data-ttu-id="6fe4e-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="6fe4e-154">Type</span></span>| <span data-ttu-id="6fe4e-155">Descrição</span><span class="sxs-lookup"><span data-stu-id="6fe4e-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="6fe4e-156">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6fe4e-156">String</span></span>|<span data-ttu-id="6fe4e-157">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="6fe4e-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="6fe4e-158">String</span><span class="sxs-lookup"><span data-stu-id="6fe4e-158">String</span></span>|<span data-ttu-id="6fe4e-159">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="6fe4e-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6fe4e-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6fe4e-160">Requirements</span></span>

|<span data-ttu-id="6fe4e-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="6fe4e-161">Requirement</span></span>| <span data-ttu-id="6fe4e-162">Valor</span><span class="sxs-lookup"><span data-stu-id="6fe4e-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fe4e-163">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6fe4e-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fe4e-164">1.1</span><span class="sxs-lookup"><span data-stu-id="6fe4e-164">1.1</span></span>|
|[<span data-ttu-id="6fe4e-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6fe4e-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fe4e-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6fe4e-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="6fe4e-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="6fe4e-167">CoercionType: String</span></span>

<span data-ttu-id="6fe4e-168">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="6fe4e-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6fe4e-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="6fe4e-169">Type</span></span>

*   <span data-ttu-id="6fe4e-170">String</span><span class="sxs-lookup"><span data-stu-id="6fe4e-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6fe4e-171">Propriedades</span><span class="sxs-lookup"><span data-stu-id="6fe4e-171">Properties</span></span>

|<span data-ttu-id="6fe4e-172">Nome</span><span class="sxs-lookup"><span data-stu-id="6fe4e-172">Name</span></span>| <span data-ttu-id="6fe4e-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="6fe4e-173">Type</span></span>| <span data-ttu-id="6fe4e-174">Descrição</span><span class="sxs-lookup"><span data-stu-id="6fe4e-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="6fe4e-175">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6fe4e-175">String</span></span>|<span data-ttu-id="6fe4e-176">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="6fe4e-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="6fe4e-177">String</span><span class="sxs-lookup"><span data-stu-id="6fe4e-177">String</span></span>|<span data-ttu-id="6fe4e-178">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="6fe4e-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6fe4e-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6fe4e-179">Requirements</span></span>

|<span data-ttu-id="6fe4e-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="6fe4e-180">Requirement</span></span>| <span data-ttu-id="6fe4e-181">Valor</span><span class="sxs-lookup"><span data-stu-id="6fe4e-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fe4e-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6fe4e-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fe4e-183">1.1</span><span class="sxs-lookup"><span data-stu-id="6fe4e-183">1.1</span></span>|
|[<span data-ttu-id="6fe4e-184">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6fe4e-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fe4e-185">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6fe4e-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="6fe4e-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="6fe4e-186">SourceProperty: String</span></span>

<span data-ttu-id="6fe4e-187">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="6fe4e-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6fe4e-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="6fe4e-188">Type</span></span>

*   <span data-ttu-id="6fe4e-189">String</span><span class="sxs-lookup"><span data-stu-id="6fe4e-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6fe4e-190">Propriedades</span><span class="sxs-lookup"><span data-stu-id="6fe4e-190">Properties</span></span>

|<span data-ttu-id="6fe4e-191">Nome</span><span class="sxs-lookup"><span data-stu-id="6fe4e-191">Name</span></span>| <span data-ttu-id="6fe4e-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="6fe4e-192">Type</span></span>| <span data-ttu-id="6fe4e-193">Descrição</span><span class="sxs-lookup"><span data-stu-id="6fe4e-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="6fe4e-194">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6fe4e-194">String</span></span>|<span data-ttu-id="6fe4e-195">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="6fe4e-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="6fe4e-196">String</span><span class="sxs-lookup"><span data-stu-id="6fe4e-196">String</span></span>|<span data-ttu-id="6fe4e-197">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="6fe4e-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6fe4e-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6fe4e-198">Requirements</span></span>

|<span data-ttu-id="6fe4e-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="6fe4e-199">Requirement</span></span>| <span data-ttu-id="6fe4e-200">Valor</span><span class="sxs-lookup"><span data-stu-id="6fe4e-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fe4e-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6fe4e-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fe4e-202">1.1</span><span class="sxs-lookup"><span data-stu-id="6fe4e-202">1.1</span></span>|
|[<span data-ttu-id="6fe4e-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6fe4e-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fe4e-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6fe4e-204">Compose or Read</span></span>|
