---
title: Namespace do Office – conjunto de requisitos 1,3
description: O modelo de objeto para o namespace de nível superior da API de suplementos do Outlook (versão da API de caixa de correio 1,3).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 706f12f4425a883f0d18fcd6f9ee18972972d72b
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717772"
---
# <a name="office"></a><span data-ttu-id="28929-103">Office</span><span class="sxs-lookup"><span data-stu-id="28929-103">Office</span></span>

<span data-ttu-id="28929-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="28929-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="28929-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28929-106">Requirements</span></span>

|<span data-ttu-id="28929-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="28929-107">Requirement</span></span>| <span data-ttu-id="28929-108">Valor</span><span class="sxs-lookup"><span data-stu-id="28929-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="28929-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28929-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28929-110">1.1</span><span class="sxs-lookup"><span data-stu-id="28929-110">1.1</span></span>|
|[<span data-ttu-id="28929-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28929-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28929-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28929-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="28929-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="28929-113">Properties</span></span>

| <span data-ttu-id="28929-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="28929-114">Property</span></span> | <span data-ttu-id="28929-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="28929-115">Modes</span></span> | <span data-ttu-id="28929-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="28929-116">Return type</span></span> | <span data-ttu-id="28929-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="28929-117">Minimum</span></span><br><span data-ttu-id="28929-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="28929-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="28929-119">context</span><span class="sxs-lookup"><span data-stu-id="28929-119">context</span></span>](office.context.md) | <span data-ttu-id="28929-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="28929-120">Compose</span></span><br><span data-ttu-id="28929-121">Ler</span><span class="sxs-lookup"><span data-stu-id="28929-121">Read</span></span> | [<span data-ttu-id="28929-122">Context</span><span class="sxs-lookup"><span data-stu-id="28929-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="28929-123">1.1</span><span class="sxs-lookup"><span data-stu-id="28929-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="28929-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="28929-124">Enumerations</span></span>

| <span data-ttu-id="28929-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="28929-125">Enumeration</span></span> | <span data-ttu-id="28929-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="28929-126">Modes</span></span> | <span data-ttu-id="28929-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="28929-127">Return type</span></span> | <span data-ttu-id="28929-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="28929-128">Minimum</span></span><br><span data-ttu-id="28929-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="28929-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="28929-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="28929-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="28929-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="28929-131">Compose</span></span><br><span data-ttu-id="28929-132">Ler</span><span class="sxs-lookup"><span data-stu-id="28929-132">Read</span></span> | <span data-ttu-id="28929-133">String</span><span class="sxs-lookup"><span data-stu-id="28929-133">String</span></span> | [<span data-ttu-id="28929-134">1.1</span><span class="sxs-lookup"><span data-stu-id="28929-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="28929-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="28929-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="28929-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="28929-136">Compose</span></span><br><span data-ttu-id="28929-137">Ler</span><span class="sxs-lookup"><span data-stu-id="28929-137">Read</span></span> | <span data-ttu-id="28929-138">String</span><span class="sxs-lookup"><span data-stu-id="28929-138">String</span></span> | [<span data-ttu-id="28929-139">1.1</span><span class="sxs-lookup"><span data-stu-id="28929-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="28929-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="28929-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="28929-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="28929-141">Compose</span></span><br><span data-ttu-id="28929-142">Ler</span><span class="sxs-lookup"><span data-stu-id="28929-142">Read</span></span> | <span data-ttu-id="28929-143">String</span><span class="sxs-lookup"><span data-stu-id="28929-143">String</span></span> | [<span data-ttu-id="28929-144">1.1</span><span class="sxs-lookup"><span data-stu-id="28929-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="28929-145">Namespaces</span><span class="sxs-lookup"><span data-stu-id="28929-145">Namespaces</span></span>

<span data-ttu-id="28929-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="28929-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="28929-147">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="28929-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="28929-148">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28929-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="28929-149">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="28929-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="28929-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="28929-150">Type</span></span>

*   <span data-ttu-id="28929-151">String</span><span class="sxs-lookup"><span data-stu-id="28929-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="28929-152">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="28929-152">Properties:</span></span>

|<span data-ttu-id="28929-153">Nome</span><span class="sxs-lookup"><span data-stu-id="28929-153">Name</span></span>| <span data-ttu-id="28929-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="28929-154">Type</span></span>| <span data-ttu-id="28929-155">Descrição</span><span class="sxs-lookup"><span data-stu-id="28929-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="28929-156">String</span><span class="sxs-lookup"><span data-stu-id="28929-156">String</span></span>|<span data-ttu-id="28929-157">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="28929-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="28929-158">String</span><span class="sxs-lookup"><span data-stu-id="28929-158">String</span></span>|<span data-ttu-id="28929-159">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="28929-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28929-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28929-160">Requirements</span></span>

|<span data-ttu-id="28929-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="28929-161">Requirement</span></span>| <span data-ttu-id="28929-162">Valor</span><span class="sxs-lookup"><span data-stu-id="28929-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="28929-163">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28929-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28929-164">1.1</span><span class="sxs-lookup"><span data-stu-id="28929-164">1.1</span></span>|
|[<span data-ttu-id="28929-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28929-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28929-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28929-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="28929-167">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28929-167">CoercionType: String</span></span>

<span data-ttu-id="28929-168">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="28929-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="28929-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="28929-169">Type</span></span>

*   <span data-ttu-id="28929-170">String</span><span class="sxs-lookup"><span data-stu-id="28929-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="28929-171">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="28929-171">Properties:</span></span>

|<span data-ttu-id="28929-172">Nome</span><span class="sxs-lookup"><span data-stu-id="28929-172">Name</span></span>| <span data-ttu-id="28929-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="28929-173">Type</span></span>| <span data-ttu-id="28929-174">Descrição</span><span class="sxs-lookup"><span data-stu-id="28929-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="28929-175">String</span><span class="sxs-lookup"><span data-stu-id="28929-175">String</span></span>|<span data-ttu-id="28929-176">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="28929-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="28929-177">String</span><span class="sxs-lookup"><span data-stu-id="28929-177">String</span></span>|<span data-ttu-id="28929-178">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="28929-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28929-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28929-179">Requirements</span></span>

|<span data-ttu-id="28929-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="28929-180">Requirement</span></span>| <span data-ttu-id="28929-181">Valor</span><span class="sxs-lookup"><span data-stu-id="28929-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="28929-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28929-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28929-183">1.1</span><span class="sxs-lookup"><span data-stu-id="28929-183">1.1</span></span>|
|[<span data-ttu-id="28929-184">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28929-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28929-185">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28929-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="28929-186">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28929-186">SourceProperty: String</span></span>

<span data-ttu-id="28929-187">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="28929-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="28929-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="28929-188">Type</span></span>

*   <span data-ttu-id="28929-189">String</span><span class="sxs-lookup"><span data-stu-id="28929-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="28929-190">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="28929-190">Properties:</span></span>

|<span data-ttu-id="28929-191">Nome</span><span class="sxs-lookup"><span data-stu-id="28929-191">Name</span></span>| <span data-ttu-id="28929-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="28929-192">Type</span></span>| <span data-ttu-id="28929-193">Descrição</span><span class="sxs-lookup"><span data-stu-id="28929-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="28929-194">String</span><span class="sxs-lookup"><span data-stu-id="28929-194">String</span></span>|<span data-ttu-id="28929-195">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="28929-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="28929-196">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="28929-196">String</span></span>|<span data-ttu-id="28929-197">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="28929-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="28929-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="28929-198">Requirements</span></span>

|<span data-ttu-id="28929-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="28929-199">Requirement</span></span>| <span data-ttu-id="28929-200">Valor</span><span class="sxs-lookup"><span data-stu-id="28929-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="28929-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="28929-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28929-202">1.1</span><span class="sxs-lookup"><span data-stu-id="28929-202">1.1</span></span>|
|[<span data-ttu-id="28929-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="28929-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28929-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="28929-204">Compose or Read</span></span>|
