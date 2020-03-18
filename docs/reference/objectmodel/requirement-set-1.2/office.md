---
title: Namespace do Office – conjunto de requisitos 1.2
description: O modelo de objeto para o namespace de nível superior da API de suplementos do Outlook (versão da API de caixa de correio 1,2).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 10445204d3007d816ebed74ede9eeab5d3dfd83c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720159"
---
# <a name="office"></a><span data-ttu-id="a2ddf-103">Office</span><span class="sxs-lookup"><span data-stu-id="a2ddf-103">Office</span></span>

<span data-ttu-id="a2ddf-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a2ddf-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a2ddf-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a2ddf-106">Requirements</span></span>

|<span data-ttu-id="a2ddf-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="a2ddf-107">Requirement</span></span>| <span data-ttu-id="a2ddf-108">Valor</span><span class="sxs-lookup"><span data-stu-id="a2ddf-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2ddf-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a2ddf-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a2ddf-110">1.1</span><span class="sxs-lookup"><span data-stu-id="a2ddf-110">1.1</span></span>|
|[<span data-ttu-id="a2ddf-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a2ddf-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a2ddf-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a2ddf-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="a2ddf-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="a2ddf-113">Properties</span></span>

| <span data-ttu-id="a2ddf-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="a2ddf-114">Property</span></span> | <span data-ttu-id="a2ddf-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="a2ddf-115">Modes</span></span> | <span data-ttu-id="a2ddf-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="a2ddf-116">Return type</span></span> | <span data-ttu-id="a2ddf-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="a2ddf-117">Minimum</span></span><br><span data-ttu-id="a2ddf-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="a2ddf-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a2ddf-119">context</span><span class="sxs-lookup"><span data-stu-id="a2ddf-119">context</span></span>](office.context.md) | <span data-ttu-id="a2ddf-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="a2ddf-120">Compose</span></span><br><span data-ttu-id="a2ddf-121">Ler</span><span class="sxs-lookup"><span data-stu-id="a2ddf-121">Read</span></span> | [<span data-ttu-id="a2ddf-122">Context</span><span class="sxs-lookup"><span data-stu-id="a2ddf-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2) | [<span data-ttu-id="a2ddf-123">1.1</span><span class="sxs-lookup"><span data-stu-id="a2ddf-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="a2ddf-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="a2ddf-124">Enumerations</span></span>

| <span data-ttu-id="a2ddf-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="a2ddf-125">Enumeration</span></span> | <span data-ttu-id="a2ddf-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="a2ddf-126">Modes</span></span> | <span data-ttu-id="a2ddf-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="a2ddf-127">Return type</span></span> | <span data-ttu-id="a2ddf-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="a2ddf-128">Minimum</span></span><br><span data-ttu-id="a2ddf-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="a2ddf-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a2ddf-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a2ddf-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a2ddf-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="a2ddf-131">Compose</span></span><br><span data-ttu-id="a2ddf-132">Ler</span><span class="sxs-lookup"><span data-stu-id="a2ddf-132">Read</span></span> | <span data-ttu-id="a2ddf-133">String</span><span class="sxs-lookup"><span data-stu-id="a2ddf-133">String</span></span> | [<span data-ttu-id="a2ddf-134">1.1</span><span class="sxs-lookup"><span data-stu-id="a2ddf-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a2ddf-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a2ddf-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a2ddf-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="a2ddf-136">Compose</span></span><br><span data-ttu-id="a2ddf-137">Ler</span><span class="sxs-lookup"><span data-stu-id="a2ddf-137">Read</span></span> | <span data-ttu-id="a2ddf-138">String</span><span class="sxs-lookup"><span data-stu-id="a2ddf-138">String</span></span> | [<span data-ttu-id="a2ddf-139">1.1</span><span class="sxs-lookup"><span data-stu-id="a2ddf-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a2ddf-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a2ddf-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a2ddf-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="a2ddf-141">Compose</span></span><br><span data-ttu-id="a2ddf-142">Ler</span><span class="sxs-lookup"><span data-stu-id="a2ddf-142">Read</span></span> | <span data-ttu-id="a2ddf-143">String</span><span class="sxs-lookup"><span data-stu-id="a2ddf-143">String</span></span> | [<span data-ttu-id="a2ddf-144">1.1</span><span class="sxs-lookup"><span data-stu-id="a2ddf-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="a2ddf-145">Namespaces</span><span class="sxs-lookup"><span data-stu-id="a2ddf-145">Namespaces</span></span>

<span data-ttu-id="a2ddf-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="a2ddf-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="a2ddf-147">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="a2ddf-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="a2ddf-148">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a2ddf-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="a2ddf-149">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="a2ddf-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a2ddf-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="a2ddf-150">Type</span></span>

*   <span data-ttu-id="a2ddf-151">String</span><span class="sxs-lookup"><span data-stu-id="a2ddf-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2ddf-152">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a2ddf-152">Properties:</span></span>

|<span data-ttu-id="a2ddf-153">Nome</span><span class="sxs-lookup"><span data-stu-id="a2ddf-153">Name</span></span>| <span data-ttu-id="a2ddf-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="a2ddf-154">Type</span></span>| <span data-ttu-id="a2ddf-155">Descrição</span><span class="sxs-lookup"><span data-stu-id="a2ddf-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a2ddf-156">String</span><span class="sxs-lookup"><span data-stu-id="a2ddf-156">String</span></span>|<span data-ttu-id="a2ddf-157">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="a2ddf-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a2ddf-158">String</span><span class="sxs-lookup"><span data-stu-id="a2ddf-158">String</span></span>|<span data-ttu-id="a2ddf-159">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="a2ddf-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a2ddf-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a2ddf-160">Requirements</span></span>

|<span data-ttu-id="a2ddf-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="a2ddf-161">Requirement</span></span>| <span data-ttu-id="a2ddf-162">Valor</span><span class="sxs-lookup"><span data-stu-id="a2ddf-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2ddf-163">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a2ddf-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a2ddf-164">1.1</span><span class="sxs-lookup"><span data-stu-id="a2ddf-164">1.1</span></span>|
|[<span data-ttu-id="a2ddf-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a2ddf-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a2ddf-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a2ddf-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="a2ddf-167">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a2ddf-167">CoercionType: String</span></span>

<span data-ttu-id="a2ddf-168">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="a2ddf-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a2ddf-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="a2ddf-169">Type</span></span>

*   <span data-ttu-id="a2ddf-170">String</span><span class="sxs-lookup"><span data-stu-id="a2ddf-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2ddf-171">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a2ddf-171">Properties:</span></span>

|<span data-ttu-id="a2ddf-172">Nome</span><span class="sxs-lookup"><span data-stu-id="a2ddf-172">Name</span></span>| <span data-ttu-id="a2ddf-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="a2ddf-173">Type</span></span>| <span data-ttu-id="a2ddf-174">Descrição</span><span class="sxs-lookup"><span data-stu-id="a2ddf-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a2ddf-175">String</span><span class="sxs-lookup"><span data-stu-id="a2ddf-175">String</span></span>|<span data-ttu-id="a2ddf-176">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="a2ddf-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a2ddf-177">String</span><span class="sxs-lookup"><span data-stu-id="a2ddf-177">String</span></span>|<span data-ttu-id="a2ddf-178">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="a2ddf-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a2ddf-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a2ddf-179">Requirements</span></span>

|<span data-ttu-id="a2ddf-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="a2ddf-180">Requirement</span></span>| <span data-ttu-id="a2ddf-181">Valor</span><span class="sxs-lookup"><span data-stu-id="a2ddf-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2ddf-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a2ddf-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a2ddf-183">1.1</span><span class="sxs-lookup"><span data-stu-id="a2ddf-183">1.1</span></span>|
|[<span data-ttu-id="a2ddf-184">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a2ddf-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a2ddf-185">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a2ddf-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="a2ddf-186">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a2ddf-186">SourceProperty: String</span></span>

<span data-ttu-id="a2ddf-187">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="a2ddf-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a2ddf-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="a2ddf-188">Type</span></span>

*   <span data-ttu-id="a2ddf-189">String</span><span class="sxs-lookup"><span data-stu-id="a2ddf-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2ddf-190">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a2ddf-190">Properties:</span></span>

|<span data-ttu-id="a2ddf-191">Nome</span><span class="sxs-lookup"><span data-stu-id="a2ddf-191">Name</span></span>| <span data-ttu-id="a2ddf-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="a2ddf-192">Type</span></span>| <span data-ttu-id="a2ddf-193">Descrição</span><span class="sxs-lookup"><span data-stu-id="a2ddf-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a2ddf-194">String</span><span class="sxs-lookup"><span data-stu-id="a2ddf-194">String</span></span>|<span data-ttu-id="a2ddf-195">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a2ddf-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a2ddf-196">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a2ddf-196">String</span></span>|<span data-ttu-id="a2ddf-197">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a2ddf-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a2ddf-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a2ddf-198">Requirements</span></span>

|<span data-ttu-id="a2ddf-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="a2ddf-199">Requirement</span></span>| <span data-ttu-id="a2ddf-200">Valor</span><span class="sxs-lookup"><span data-stu-id="a2ddf-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2ddf-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a2ddf-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a2ddf-202">1.1</span><span class="sxs-lookup"><span data-stu-id="a2ddf-202">1.1</span></span>|
|[<span data-ttu-id="a2ddf-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a2ddf-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a2ddf-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a2ddf-204">Compose or Read</span></span>|
