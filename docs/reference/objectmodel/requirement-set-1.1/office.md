---
title: Namespace do Office – conjunto de requisitos 1,1
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,1.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: b821e4ee3f0a2ea2240ab0ff77131e12b3106c57
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431378"
---
# <a name="office-mailbox-requirement-set-11"></a><span data-ttu-id="adc36-103">Office (conjunto de requisitos de caixa de correio 1,1)</span><span class="sxs-lookup"><span data-stu-id="adc36-103">Office (Mailbox requirement set 1.1)</span></span>

<span data-ttu-id="adc36-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="adc36-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="adc36-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="adc36-106">Requirements</span></span>

|<span data-ttu-id="adc36-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="adc36-107">Requirement</span></span>| <span data-ttu-id="adc36-108">Valor</span><span class="sxs-lookup"><span data-stu-id="adc36-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="adc36-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="adc36-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="adc36-110">1.1</span><span class="sxs-lookup"><span data-stu-id="adc36-110">1.1</span></span>|
|[<span data-ttu-id="adc36-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="adc36-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="adc36-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="adc36-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="adc36-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="adc36-113">Properties</span></span>

| <span data-ttu-id="adc36-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="adc36-114">Property</span></span> | <span data-ttu-id="adc36-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="adc36-115">Modes</span></span> | <span data-ttu-id="adc36-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="adc36-116">Return type</span></span> | <span data-ttu-id="adc36-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="adc36-117">Minimum</span></span><br><span data-ttu-id="adc36-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="adc36-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="adc36-119">context</span><span class="sxs-lookup"><span data-stu-id="adc36-119">context</span></span>](office.context.md) | <span data-ttu-id="adc36-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="adc36-120">Compose</span></span><br><span data-ttu-id="adc36-121">Ler</span><span class="sxs-lookup"><span data-stu-id="adc36-121">Read</span></span> | [<span data-ttu-id="adc36-122">Context</span><span class="sxs-lookup"><span data-stu-id="adc36-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="adc36-123">1.1</span><span class="sxs-lookup"><span data-stu-id="adc36-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="adc36-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="adc36-124">Enumerations</span></span>

| <span data-ttu-id="adc36-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="adc36-125">Enumeration</span></span> | <span data-ttu-id="adc36-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="adc36-126">Modes</span></span> | <span data-ttu-id="adc36-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="adc36-127">Return type</span></span> | <span data-ttu-id="adc36-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="adc36-128">Minimum</span></span><br><span data-ttu-id="adc36-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="adc36-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="adc36-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="adc36-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="adc36-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="adc36-131">Compose</span></span><br><span data-ttu-id="adc36-132">Ler</span><span class="sxs-lookup"><span data-stu-id="adc36-132">Read</span></span> | <span data-ttu-id="adc36-133">String</span><span class="sxs-lookup"><span data-stu-id="adc36-133">String</span></span> | [<span data-ttu-id="adc36-134">1.1</span><span class="sxs-lookup"><span data-stu-id="adc36-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="adc36-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="adc36-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="adc36-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="adc36-136">Compose</span></span><br><span data-ttu-id="adc36-137">Ler</span><span class="sxs-lookup"><span data-stu-id="adc36-137">Read</span></span> | <span data-ttu-id="adc36-138">String</span><span class="sxs-lookup"><span data-stu-id="adc36-138">String</span></span> | [<span data-ttu-id="adc36-139">1.1</span><span class="sxs-lookup"><span data-stu-id="adc36-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="adc36-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="adc36-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="adc36-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="adc36-141">Compose</span></span><br><span data-ttu-id="adc36-142">Ler</span><span class="sxs-lookup"><span data-stu-id="adc36-142">Read</span></span> | <span data-ttu-id="adc36-143">String</span><span class="sxs-lookup"><span data-stu-id="adc36-143">String</span></span> | [<span data-ttu-id="adc36-144">1.1</span><span class="sxs-lookup"><span data-stu-id="adc36-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="adc36-145">Namespaces</span><span class="sxs-lookup"><span data-stu-id="adc36-145">Namespaces</span></span>

<span data-ttu-id="adc36-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1&preserve-view=true): inclui uma série de enumerações específicas do Outlook, por exemplo,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` e `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="adc36-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="adc36-147">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="adc36-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="adc36-148">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="adc36-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="adc36-149">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="adc36-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="adc36-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="adc36-150">Type</span></span>

*   <span data-ttu-id="adc36-151">String</span><span class="sxs-lookup"><span data-stu-id="adc36-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="adc36-152">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="adc36-152">Properties:</span></span>

|<span data-ttu-id="adc36-153">Nome</span><span class="sxs-lookup"><span data-stu-id="adc36-153">Name</span></span>| <span data-ttu-id="adc36-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="adc36-154">Type</span></span>| <span data-ttu-id="adc36-155">Descrição</span><span class="sxs-lookup"><span data-stu-id="adc36-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="adc36-156">String</span><span class="sxs-lookup"><span data-stu-id="adc36-156">String</span></span>|<span data-ttu-id="adc36-157">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="adc36-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="adc36-158">String</span><span class="sxs-lookup"><span data-stu-id="adc36-158">String</span></span>|<span data-ttu-id="adc36-159">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="adc36-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="adc36-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="adc36-160">Requirements</span></span>

|<span data-ttu-id="adc36-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="adc36-161">Requirement</span></span>| <span data-ttu-id="adc36-162">Valor</span><span class="sxs-lookup"><span data-stu-id="adc36-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="adc36-163">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="adc36-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="adc36-164">1.1</span><span class="sxs-lookup"><span data-stu-id="adc36-164">1.1</span></span>|
|[<span data-ttu-id="adc36-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="adc36-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="adc36-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="adc36-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="adc36-167">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="adc36-167">CoercionType: String</span></span>

<span data-ttu-id="adc36-168">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="adc36-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="adc36-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="adc36-169">Type</span></span>

*   <span data-ttu-id="adc36-170">String</span><span class="sxs-lookup"><span data-stu-id="adc36-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="adc36-171">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="adc36-171">Properties:</span></span>

|<span data-ttu-id="adc36-172">Nome</span><span class="sxs-lookup"><span data-stu-id="adc36-172">Name</span></span>| <span data-ttu-id="adc36-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="adc36-173">Type</span></span>| <span data-ttu-id="adc36-174">Descrição</span><span class="sxs-lookup"><span data-stu-id="adc36-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="adc36-175">String</span><span class="sxs-lookup"><span data-stu-id="adc36-175">String</span></span>|<span data-ttu-id="adc36-176">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="adc36-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="adc36-177">String</span><span class="sxs-lookup"><span data-stu-id="adc36-177">String</span></span>|<span data-ttu-id="adc36-178">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="adc36-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="adc36-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="adc36-179">Requirements</span></span>

|<span data-ttu-id="adc36-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="adc36-180">Requirement</span></span>| <span data-ttu-id="adc36-181">Valor</span><span class="sxs-lookup"><span data-stu-id="adc36-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="adc36-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="adc36-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="adc36-183">1.1</span><span class="sxs-lookup"><span data-stu-id="adc36-183">1.1</span></span>|
|[<span data-ttu-id="adc36-184">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="adc36-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="adc36-185">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="adc36-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="adc36-186">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="adc36-186">SourceProperty: String</span></span>

<span data-ttu-id="adc36-187">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="adc36-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="adc36-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="adc36-188">Type</span></span>

*   <span data-ttu-id="adc36-189">String</span><span class="sxs-lookup"><span data-stu-id="adc36-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="adc36-190">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="adc36-190">Properties:</span></span>

|<span data-ttu-id="adc36-191">Nome</span><span class="sxs-lookup"><span data-stu-id="adc36-191">Name</span></span>| <span data-ttu-id="adc36-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="adc36-192">Type</span></span>| <span data-ttu-id="adc36-193">Descrição</span><span class="sxs-lookup"><span data-stu-id="adc36-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="adc36-194">String</span><span class="sxs-lookup"><span data-stu-id="adc36-194">String</span></span>|<span data-ttu-id="adc36-195">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="adc36-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="adc36-196">String</span><span class="sxs-lookup"><span data-stu-id="adc36-196">String</span></span>|<span data-ttu-id="adc36-197">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="adc36-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="adc36-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="adc36-198">Requirements</span></span>

|<span data-ttu-id="adc36-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="adc36-199">Requirement</span></span>| <span data-ttu-id="adc36-200">Valor</span><span class="sxs-lookup"><span data-stu-id="adc36-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="adc36-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="adc36-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="adc36-202">1.1</span><span class="sxs-lookup"><span data-stu-id="adc36-202">1.1</span></span>|
|[<span data-ttu-id="adc36-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="adc36-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="adc36-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="adc36-204">Compose or Read</span></span>|