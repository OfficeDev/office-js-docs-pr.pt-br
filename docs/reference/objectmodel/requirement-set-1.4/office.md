---
title: Namespace do Office – conjunto de requisitos 1,4
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,4.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: a1a19fb3432b703facfd69ca45fd35660e5e535b
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609822"
---
# <a name="office-mailbox-requirement-set-14"></a><span data-ttu-id="d4b07-103">Office (conjunto de requisitos de caixa de correio 1,4)</span><span class="sxs-lookup"><span data-stu-id="d4b07-103">Office (Mailbox requirement set 1.4)</span></span>

<span data-ttu-id="d4b07-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d4b07-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4b07-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d4b07-106">Requirements</span></span>

|<span data-ttu-id="d4b07-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="d4b07-107">Requirement</span></span>| <span data-ttu-id="d4b07-108">Valor</span><span class="sxs-lookup"><span data-stu-id="d4b07-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4b07-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d4b07-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d4b07-110">1.1</span><span class="sxs-lookup"><span data-stu-id="d4b07-110">1.1</span></span>|
|[<span data-ttu-id="d4b07-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d4b07-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d4b07-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d4b07-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="d4b07-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="d4b07-113">Properties</span></span>

| <span data-ttu-id="d4b07-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="d4b07-114">Property</span></span> | <span data-ttu-id="d4b07-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="d4b07-115">Modes</span></span> | <span data-ttu-id="d4b07-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="d4b07-116">Return type</span></span> | <span data-ttu-id="d4b07-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="d4b07-117">Minimum</span></span><br><span data-ttu-id="d4b07-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="d4b07-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d4b07-119">context</span><span class="sxs-lookup"><span data-stu-id="d4b07-119">context</span></span>](office.context.md) | <span data-ttu-id="d4b07-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="d4b07-120">Compose</span></span><br><span data-ttu-id="d4b07-121">Read</span><span class="sxs-lookup"><span data-stu-id="d4b07-121">Read</span></span> | [<span data-ttu-id="d4b07-122">Context</span><span class="sxs-lookup"><span data-stu-id="d4b07-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4) | [<span data-ttu-id="d4b07-123">1.1</span><span class="sxs-lookup"><span data-stu-id="d4b07-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="d4b07-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="d4b07-124">Enumerations</span></span>

| <span data-ttu-id="d4b07-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="d4b07-125">Enumeration</span></span> | <span data-ttu-id="d4b07-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="d4b07-126">Modes</span></span> | <span data-ttu-id="d4b07-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="d4b07-127">Return type</span></span> | <span data-ttu-id="d4b07-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="d4b07-128">Minimum</span></span><br><span data-ttu-id="d4b07-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="d4b07-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d4b07-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d4b07-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d4b07-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="d4b07-131">Compose</span></span><br><span data-ttu-id="d4b07-132">Read</span><span class="sxs-lookup"><span data-stu-id="d4b07-132">Read</span></span> | <span data-ttu-id="d4b07-133">String</span><span class="sxs-lookup"><span data-stu-id="d4b07-133">String</span></span> | [<span data-ttu-id="d4b07-134">1.1</span><span class="sxs-lookup"><span data-stu-id="d4b07-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d4b07-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d4b07-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d4b07-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="d4b07-136">Compose</span></span><br><span data-ttu-id="d4b07-137">Read</span><span class="sxs-lookup"><span data-stu-id="d4b07-137">Read</span></span> | <span data-ttu-id="d4b07-138">String</span><span class="sxs-lookup"><span data-stu-id="d4b07-138">String</span></span> | [<span data-ttu-id="d4b07-139">1.1</span><span class="sxs-lookup"><span data-stu-id="d4b07-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d4b07-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d4b07-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d4b07-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="d4b07-141">Compose</span></span><br><span data-ttu-id="d4b07-142">Read</span><span class="sxs-lookup"><span data-stu-id="d4b07-142">Read</span></span> | <span data-ttu-id="d4b07-143">String</span><span class="sxs-lookup"><span data-stu-id="d4b07-143">String</span></span> | [<span data-ttu-id="d4b07-144">1.1</span><span class="sxs-lookup"><span data-stu-id="d4b07-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="d4b07-145">Namespaces</span><span class="sxs-lookup"><span data-stu-id="d4b07-145">Namespaces</span></span>

<span data-ttu-id="d4b07-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): inclui uma série de enumerações específicas do Outlook, por exemplo,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` e `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="d4b07-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="d4b07-147">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="d4b07-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="d4b07-148">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d4b07-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="d4b07-149">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="d4b07-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d4b07-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="d4b07-150">Type</span></span>

*   <span data-ttu-id="d4b07-151">String</span><span class="sxs-lookup"><span data-stu-id="d4b07-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d4b07-152">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="d4b07-152">Properties:</span></span>

|<span data-ttu-id="d4b07-153">Nome</span><span class="sxs-lookup"><span data-stu-id="d4b07-153">Name</span></span>| <span data-ttu-id="d4b07-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="d4b07-154">Type</span></span>| <span data-ttu-id="d4b07-155">Descrição</span><span class="sxs-lookup"><span data-stu-id="d4b07-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d4b07-156">String</span><span class="sxs-lookup"><span data-stu-id="d4b07-156">String</span></span>|<span data-ttu-id="d4b07-157">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="d4b07-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d4b07-158">String</span><span class="sxs-lookup"><span data-stu-id="d4b07-158">String</span></span>|<span data-ttu-id="d4b07-159">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="d4b07-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4b07-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d4b07-160">Requirements</span></span>

|<span data-ttu-id="d4b07-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="d4b07-161">Requirement</span></span>| <span data-ttu-id="d4b07-162">Valor</span><span class="sxs-lookup"><span data-stu-id="d4b07-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4b07-163">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d4b07-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d4b07-164">1.1</span><span class="sxs-lookup"><span data-stu-id="d4b07-164">1.1</span></span>|
|[<span data-ttu-id="d4b07-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d4b07-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d4b07-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d4b07-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="d4b07-167">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d4b07-167">CoercionType: String</span></span>

<span data-ttu-id="d4b07-168">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="d4b07-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d4b07-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="d4b07-169">Type</span></span>

*   <span data-ttu-id="d4b07-170">String</span><span class="sxs-lookup"><span data-stu-id="d4b07-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d4b07-171">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="d4b07-171">Properties:</span></span>

|<span data-ttu-id="d4b07-172">Nome</span><span class="sxs-lookup"><span data-stu-id="d4b07-172">Name</span></span>| <span data-ttu-id="d4b07-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="d4b07-173">Type</span></span>| <span data-ttu-id="d4b07-174">Descrição</span><span class="sxs-lookup"><span data-stu-id="d4b07-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d4b07-175">String</span><span class="sxs-lookup"><span data-stu-id="d4b07-175">String</span></span>|<span data-ttu-id="d4b07-176">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="d4b07-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d4b07-177">String</span><span class="sxs-lookup"><span data-stu-id="d4b07-177">String</span></span>|<span data-ttu-id="d4b07-178">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="d4b07-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4b07-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d4b07-179">Requirements</span></span>

|<span data-ttu-id="d4b07-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="d4b07-180">Requirement</span></span>| <span data-ttu-id="d4b07-181">Valor</span><span class="sxs-lookup"><span data-stu-id="d4b07-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4b07-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d4b07-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d4b07-183">1.1</span><span class="sxs-lookup"><span data-stu-id="d4b07-183">1.1</span></span>|
|[<span data-ttu-id="d4b07-184">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d4b07-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d4b07-185">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d4b07-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="d4b07-186">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d4b07-186">SourceProperty: String</span></span>

<span data-ttu-id="d4b07-187">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="d4b07-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d4b07-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="d4b07-188">Type</span></span>

*   <span data-ttu-id="d4b07-189">String</span><span class="sxs-lookup"><span data-stu-id="d4b07-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d4b07-190">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="d4b07-190">Properties:</span></span>

|<span data-ttu-id="d4b07-191">Nome</span><span class="sxs-lookup"><span data-stu-id="d4b07-191">Name</span></span>| <span data-ttu-id="d4b07-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="d4b07-192">Type</span></span>| <span data-ttu-id="d4b07-193">Descrição</span><span class="sxs-lookup"><span data-stu-id="d4b07-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d4b07-194">String</span><span class="sxs-lookup"><span data-stu-id="d4b07-194">String</span></span>|<span data-ttu-id="d4b07-195">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="d4b07-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d4b07-196">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d4b07-196">String</span></span>|<span data-ttu-id="d4b07-197">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="d4b07-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4b07-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d4b07-198">Requirements</span></span>

|<span data-ttu-id="d4b07-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="d4b07-199">Requirement</span></span>| <span data-ttu-id="d4b07-200">Valor</span><span class="sxs-lookup"><span data-stu-id="d4b07-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4b07-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d4b07-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d4b07-202">1.1</span><span class="sxs-lookup"><span data-stu-id="d4b07-202">1.1</span></span>|
|[<span data-ttu-id="d4b07-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d4b07-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d4b07-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d4b07-204">Compose or Read</span></span>|
