---
title: Namespace do Office – conjunto de requisitos 1,4
description: O modelo de objeto para o namespace de nível superior da API de suplementos do Outlook (versão da API de caixa de correio 1,4).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: e5a5c6de5bb87cb32968d9d9d80c621f0acc238d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720054"
---
# <a name="office"></a><span data-ttu-id="a3f38-103">Office</span><span class="sxs-lookup"><span data-stu-id="a3f38-103">Office</span></span>

<span data-ttu-id="a3f38-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a3f38-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3f38-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3f38-106">Requirements</span></span>

|<span data-ttu-id="a3f38-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="a3f38-107">Requirement</span></span>| <span data-ttu-id="a3f38-108">Valor</span><span class="sxs-lookup"><span data-stu-id="a3f38-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3f38-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a3f38-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a3f38-110">1.1</span><span class="sxs-lookup"><span data-stu-id="a3f38-110">1.1</span></span>|
|[<span data-ttu-id="a3f38-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a3f38-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a3f38-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a3f38-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="a3f38-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="a3f38-113">Properties</span></span>

| <span data-ttu-id="a3f38-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="a3f38-114">Property</span></span> | <span data-ttu-id="a3f38-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="a3f38-115">Modes</span></span> | <span data-ttu-id="a3f38-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="a3f38-116">Return type</span></span> | <span data-ttu-id="a3f38-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="a3f38-117">Minimum</span></span><br><span data-ttu-id="a3f38-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="a3f38-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a3f38-119">context</span><span class="sxs-lookup"><span data-stu-id="a3f38-119">context</span></span>](office.context.md) | <span data-ttu-id="a3f38-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="a3f38-120">Compose</span></span><br><span data-ttu-id="a3f38-121">Ler</span><span class="sxs-lookup"><span data-stu-id="a3f38-121">Read</span></span> | [<span data-ttu-id="a3f38-122">Context</span><span class="sxs-lookup"><span data-stu-id="a3f38-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4) | [<span data-ttu-id="a3f38-123">1.1</span><span class="sxs-lookup"><span data-stu-id="a3f38-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="a3f38-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="a3f38-124">Enumerations</span></span>

| <span data-ttu-id="a3f38-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="a3f38-125">Enumeration</span></span> | <span data-ttu-id="a3f38-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="a3f38-126">Modes</span></span> | <span data-ttu-id="a3f38-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="a3f38-127">Return type</span></span> | <span data-ttu-id="a3f38-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="a3f38-128">Minimum</span></span><br><span data-ttu-id="a3f38-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="a3f38-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a3f38-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a3f38-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a3f38-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="a3f38-131">Compose</span></span><br><span data-ttu-id="a3f38-132">Ler</span><span class="sxs-lookup"><span data-stu-id="a3f38-132">Read</span></span> | <span data-ttu-id="a3f38-133">String</span><span class="sxs-lookup"><span data-stu-id="a3f38-133">String</span></span> | [<span data-ttu-id="a3f38-134">1.1</span><span class="sxs-lookup"><span data-stu-id="a3f38-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a3f38-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a3f38-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a3f38-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="a3f38-136">Compose</span></span><br><span data-ttu-id="a3f38-137">Ler</span><span class="sxs-lookup"><span data-stu-id="a3f38-137">Read</span></span> | <span data-ttu-id="a3f38-138">String</span><span class="sxs-lookup"><span data-stu-id="a3f38-138">String</span></span> | [<span data-ttu-id="a3f38-139">1.1</span><span class="sxs-lookup"><span data-stu-id="a3f38-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a3f38-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a3f38-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a3f38-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="a3f38-141">Compose</span></span><br><span data-ttu-id="a3f38-142">Ler</span><span class="sxs-lookup"><span data-stu-id="a3f38-142">Read</span></span> | <span data-ttu-id="a3f38-143">String</span><span class="sxs-lookup"><span data-stu-id="a3f38-143">String</span></span> | [<span data-ttu-id="a3f38-144">1.1</span><span class="sxs-lookup"><span data-stu-id="a3f38-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="a3f38-145">Namespaces</span><span class="sxs-lookup"><span data-stu-id="a3f38-145">Namespaces</span></span>

<span data-ttu-id="a3f38-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="a3f38-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="a3f38-147">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="a3f38-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="a3f38-148">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a3f38-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="a3f38-149">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="a3f38-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a3f38-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3f38-150">Type</span></span>

*   <span data-ttu-id="a3f38-151">String</span><span class="sxs-lookup"><span data-stu-id="a3f38-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a3f38-152">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a3f38-152">Properties:</span></span>

|<span data-ttu-id="a3f38-153">Nome</span><span class="sxs-lookup"><span data-stu-id="a3f38-153">Name</span></span>| <span data-ttu-id="a3f38-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3f38-154">Type</span></span>| <span data-ttu-id="a3f38-155">Descrição</span><span class="sxs-lookup"><span data-stu-id="a3f38-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a3f38-156">String</span><span class="sxs-lookup"><span data-stu-id="a3f38-156">String</span></span>|<span data-ttu-id="a3f38-157">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="a3f38-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a3f38-158">String</span><span class="sxs-lookup"><span data-stu-id="a3f38-158">String</span></span>|<span data-ttu-id="a3f38-159">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="a3f38-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3f38-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3f38-160">Requirements</span></span>

|<span data-ttu-id="a3f38-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="a3f38-161">Requirement</span></span>| <span data-ttu-id="a3f38-162">Valor</span><span class="sxs-lookup"><span data-stu-id="a3f38-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3f38-163">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a3f38-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a3f38-164">1.1</span><span class="sxs-lookup"><span data-stu-id="a3f38-164">1.1</span></span>|
|[<span data-ttu-id="a3f38-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a3f38-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a3f38-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a3f38-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="a3f38-167">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a3f38-167">CoercionType: String</span></span>

<span data-ttu-id="a3f38-168">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="a3f38-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a3f38-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3f38-169">Type</span></span>

*   <span data-ttu-id="a3f38-170">String</span><span class="sxs-lookup"><span data-stu-id="a3f38-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a3f38-171">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a3f38-171">Properties:</span></span>

|<span data-ttu-id="a3f38-172">Nome</span><span class="sxs-lookup"><span data-stu-id="a3f38-172">Name</span></span>| <span data-ttu-id="a3f38-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3f38-173">Type</span></span>| <span data-ttu-id="a3f38-174">Descrição</span><span class="sxs-lookup"><span data-stu-id="a3f38-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a3f38-175">String</span><span class="sxs-lookup"><span data-stu-id="a3f38-175">String</span></span>|<span data-ttu-id="a3f38-176">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="a3f38-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a3f38-177">String</span><span class="sxs-lookup"><span data-stu-id="a3f38-177">String</span></span>|<span data-ttu-id="a3f38-178">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="a3f38-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3f38-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3f38-179">Requirements</span></span>

|<span data-ttu-id="a3f38-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="a3f38-180">Requirement</span></span>| <span data-ttu-id="a3f38-181">Valor</span><span class="sxs-lookup"><span data-stu-id="a3f38-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3f38-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a3f38-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a3f38-183">1.1</span><span class="sxs-lookup"><span data-stu-id="a3f38-183">1.1</span></span>|
|[<span data-ttu-id="a3f38-184">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a3f38-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a3f38-185">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a3f38-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="a3f38-186">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a3f38-186">SourceProperty: String</span></span>

<span data-ttu-id="a3f38-187">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="a3f38-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a3f38-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3f38-188">Type</span></span>

*   <span data-ttu-id="a3f38-189">String</span><span class="sxs-lookup"><span data-stu-id="a3f38-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a3f38-190">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a3f38-190">Properties:</span></span>

|<span data-ttu-id="a3f38-191">Nome</span><span class="sxs-lookup"><span data-stu-id="a3f38-191">Name</span></span>| <span data-ttu-id="a3f38-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="a3f38-192">Type</span></span>| <span data-ttu-id="a3f38-193">Descrição</span><span class="sxs-lookup"><span data-stu-id="a3f38-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a3f38-194">String</span><span class="sxs-lookup"><span data-stu-id="a3f38-194">String</span></span>|<span data-ttu-id="a3f38-195">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a3f38-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a3f38-196">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a3f38-196">String</span></span>|<span data-ttu-id="a3f38-197">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a3f38-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3f38-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a3f38-198">Requirements</span></span>

|<span data-ttu-id="a3f38-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="a3f38-199">Requirement</span></span>| <span data-ttu-id="a3f38-200">Valor</span><span class="sxs-lookup"><span data-stu-id="a3f38-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3f38-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a3f38-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a3f38-202">1.1</span><span class="sxs-lookup"><span data-stu-id="a3f38-202">1.1</span></span>|
|[<span data-ttu-id="a3f38-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a3f38-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a3f38-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a3f38-204">Compose or Read</span></span>|
