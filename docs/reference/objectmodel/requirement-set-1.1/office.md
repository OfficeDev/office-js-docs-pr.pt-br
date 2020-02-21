---
title: Namespace do Office – conjunto de requisitos 1,1
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 68363f101b4c818853cc118e39d05784c56ef3ad
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165472"
---
# <a name="office"></a><span data-ttu-id="a1610-102">Office</span><span class="sxs-lookup"><span data-stu-id="a1610-102">Office</span></span>

<span data-ttu-id="a1610-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a1610-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1610-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a1610-105">Requirements</span></span>

|<span data-ttu-id="a1610-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="a1610-106">Requirement</span></span>| <span data-ttu-id="a1610-107">Valor</span><span class="sxs-lookup"><span data-stu-id="a1610-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1610-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a1610-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1610-109">1.1</span><span class="sxs-lookup"><span data-stu-id="a1610-109">1.1</span></span>|
|[<span data-ttu-id="a1610-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a1610-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1610-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a1610-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="a1610-112">Propriedades</span><span class="sxs-lookup"><span data-stu-id="a1610-112">Properties</span></span>

| <span data-ttu-id="a1610-113">Propriedade</span><span class="sxs-lookup"><span data-stu-id="a1610-113">Property</span></span> | <span data-ttu-id="a1610-114">Modelos</span><span class="sxs-lookup"><span data-stu-id="a1610-114">Modes</span></span> | <span data-ttu-id="a1610-115">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="a1610-115">Return type</span></span> | <span data-ttu-id="a1610-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="a1610-116">Minimum</span></span><br><span data-ttu-id="a1610-117">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="a1610-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a1610-118">context</span><span class="sxs-lookup"><span data-stu-id="a1610-118">context</span></span>](office.context.md) | <span data-ttu-id="a1610-119">Escrever</span><span class="sxs-lookup"><span data-stu-id="a1610-119">Compose</span></span><br><span data-ttu-id="a1610-120">Ler</span><span class="sxs-lookup"><span data-stu-id="a1610-120">Read</span></span> | [<span data-ttu-id="a1610-121">Context</span><span class="sxs-lookup"><span data-stu-id="a1610-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="a1610-122">1.1</span><span class="sxs-lookup"><span data-stu-id="a1610-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="a1610-123">Enumerações</span><span class="sxs-lookup"><span data-stu-id="a1610-123">Enumerations</span></span>

| <span data-ttu-id="a1610-124">Enumeração</span><span class="sxs-lookup"><span data-stu-id="a1610-124">Enumeration</span></span> | <span data-ttu-id="a1610-125">Modelos</span><span class="sxs-lookup"><span data-stu-id="a1610-125">Modes</span></span> | <span data-ttu-id="a1610-126">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="a1610-126">Return type</span></span> | <span data-ttu-id="a1610-127">Mínimo</span><span class="sxs-lookup"><span data-stu-id="a1610-127">Minimum</span></span><br><span data-ttu-id="a1610-128">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="a1610-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a1610-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a1610-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a1610-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="a1610-130">Compose</span></span><br><span data-ttu-id="a1610-131">Ler</span><span class="sxs-lookup"><span data-stu-id="a1610-131">Read</span></span> | <span data-ttu-id="a1610-132">String</span><span class="sxs-lookup"><span data-stu-id="a1610-132">String</span></span> | [<span data-ttu-id="a1610-133">1.1</span><span class="sxs-lookup"><span data-stu-id="a1610-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a1610-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a1610-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a1610-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="a1610-135">Compose</span></span><br><span data-ttu-id="a1610-136">Ler</span><span class="sxs-lookup"><span data-stu-id="a1610-136">Read</span></span> | <span data-ttu-id="a1610-137">String</span><span class="sxs-lookup"><span data-stu-id="a1610-137">String</span></span> | [<span data-ttu-id="a1610-138">1.1</span><span class="sxs-lookup"><span data-stu-id="a1610-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a1610-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a1610-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a1610-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="a1610-140">Compose</span></span><br><span data-ttu-id="a1610-141">Ler</span><span class="sxs-lookup"><span data-stu-id="a1610-141">Read</span></span> | <span data-ttu-id="a1610-142">String</span><span class="sxs-lookup"><span data-stu-id="a1610-142">String</span></span> | [<span data-ttu-id="a1610-143">1.1</span><span class="sxs-lookup"><span data-stu-id="a1610-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="a1610-144">Namespaces</span><span class="sxs-lookup"><span data-stu-id="a1610-144">Namespaces</span></span>

<span data-ttu-id="a1610-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="a1610-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="a1610-146">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="a1610-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="a1610-147">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1610-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="a1610-148">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="a1610-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a1610-149">Tipo</span><span class="sxs-lookup"><span data-stu-id="a1610-149">Type</span></span>

*   <span data-ttu-id="a1610-150">String</span><span class="sxs-lookup"><span data-stu-id="a1610-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a1610-151">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a1610-151">Properties:</span></span>

|<span data-ttu-id="a1610-152">Nome</span><span class="sxs-lookup"><span data-stu-id="a1610-152">Name</span></span>| <span data-ttu-id="a1610-153">Tipo</span><span class="sxs-lookup"><span data-stu-id="a1610-153">Type</span></span>| <span data-ttu-id="a1610-154">Descrição</span><span class="sxs-lookup"><span data-stu-id="a1610-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a1610-155">String</span><span class="sxs-lookup"><span data-stu-id="a1610-155">String</span></span>|<span data-ttu-id="a1610-156">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="a1610-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a1610-157">String</span><span class="sxs-lookup"><span data-stu-id="a1610-157">String</span></span>|<span data-ttu-id="a1610-158">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="a1610-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a1610-159">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a1610-159">Requirements</span></span>

|<span data-ttu-id="a1610-160">Requisito</span><span class="sxs-lookup"><span data-stu-id="a1610-160">Requirement</span></span>| <span data-ttu-id="a1610-161">Valor</span><span class="sxs-lookup"><span data-stu-id="a1610-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1610-162">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a1610-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1610-163">1.1</span><span class="sxs-lookup"><span data-stu-id="a1610-163">1.1</span></span>|
|[<span data-ttu-id="a1610-164">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a1610-164">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1610-165">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a1610-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="a1610-166">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1610-166">CoercionType: String</span></span>

<span data-ttu-id="a1610-167">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="a1610-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a1610-168">Tipo</span><span class="sxs-lookup"><span data-stu-id="a1610-168">Type</span></span>

*   <span data-ttu-id="a1610-169">String</span><span class="sxs-lookup"><span data-stu-id="a1610-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a1610-170">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a1610-170">Properties:</span></span>

|<span data-ttu-id="a1610-171">Nome</span><span class="sxs-lookup"><span data-stu-id="a1610-171">Name</span></span>| <span data-ttu-id="a1610-172">Tipo</span><span class="sxs-lookup"><span data-stu-id="a1610-172">Type</span></span>| <span data-ttu-id="a1610-173">Descrição</span><span class="sxs-lookup"><span data-stu-id="a1610-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a1610-174">String</span><span class="sxs-lookup"><span data-stu-id="a1610-174">String</span></span>|<span data-ttu-id="a1610-175">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="a1610-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a1610-176">String</span><span class="sxs-lookup"><span data-stu-id="a1610-176">String</span></span>|<span data-ttu-id="a1610-177">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="a1610-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a1610-178">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a1610-178">Requirements</span></span>

|<span data-ttu-id="a1610-179">Requisito</span><span class="sxs-lookup"><span data-stu-id="a1610-179">Requirement</span></span>| <span data-ttu-id="a1610-180">Valor</span><span class="sxs-lookup"><span data-stu-id="a1610-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1610-181">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a1610-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1610-182">1.1</span><span class="sxs-lookup"><span data-stu-id="a1610-182">1.1</span></span>|
|[<span data-ttu-id="a1610-183">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a1610-183">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1610-184">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a1610-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="a1610-185">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1610-185">SourceProperty: String</span></span>

<span data-ttu-id="a1610-186">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="a1610-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a1610-187">Tipo</span><span class="sxs-lookup"><span data-stu-id="a1610-187">Type</span></span>

*   <span data-ttu-id="a1610-188">String</span><span class="sxs-lookup"><span data-stu-id="a1610-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a1610-189">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a1610-189">Properties:</span></span>

|<span data-ttu-id="a1610-190">Nome</span><span class="sxs-lookup"><span data-stu-id="a1610-190">Name</span></span>| <span data-ttu-id="a1610-191">Tipo</span><span class="sxs-lookup"><span data-stu-id="a1610-191">Type</span></span>| <span data-ttu-id="a1610-192">Descrição</span><span class="sxs-lookup"><span data-stu-id="a1610-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a1610-193">String</span><span class="sxs-lookup"><span data-stu-id="a1610-193">String</span></span>|<span data-ttu-id="a1610-194">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a1610-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a1610-195">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a1610-195">String</span></span>|<span data-ttu-id="a1610-196">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a1610-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a1610-197">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a1610-197">Requirements</span></span>

|<span data-ttu-id="a1610-198">Requisito</span><span class="sxs-lookup"><span data-stu-id="a1610-198">Requirement</span></span>| <span data-ttu-id="a1610-199">Valor</span><span class="sxs-lookup"><span data-stu-id="a1610-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1610-200">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a1610-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1610-201">1.1</span><span class="sxs-lookup"><span data-stu-id="a1610-201">1.1</span></span>|
|[<span data-ttu-id="a1610-202">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a1610-202">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1610-203">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a1610-203">Compose or Read</span></span>|