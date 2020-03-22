---
title: Namespace do Office – conjunto de requisitos 1,3
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,3.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: eb3ebba75df8345402ab0ce4ac2b5cc5f0354e6c
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890694"
---
# <a name="office-mailbox-requirement-set-13"></a><span data-ttu-id="603fa-103">Office (conjunto de requisitos de caixa de correio 1,3)</span><span class="sxs-lookup"><span data-stu-id="603fa-103">Office (Mailbox requirement set 1.3)</span></span>

<span data-ttu-id="603fa-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="603fa-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="603fa-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="603fa-106">Requirements</span></span>

|<span data-ttu-id="603fa-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="603fa-107">Requirement</span></span>| <span data-ttu-id="603fa-108">Valor</span><span class="sxs-lookup"><span data-stu-id="603fa-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="603fa-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="603fa-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="603fa-110">1.1</span><span class="sxs-lookup"><span data-stu-id="603fa-110">1.1</span></span>|
|[<span data-ttu-id="603fa-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="603fa-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="603fa-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="603fa-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="603fa-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="603fa-113">Properties</span></span>

| <span data-ttu-id="603fa-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="603fa-114">Property</span></span> | <span data-ttu-id="603fa-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="603fa-115">Modes</span></span> | <span data-ttu-id="603fa-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="603fa-116">Return type</span></span> | <span data-ttu-id="603fa-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="603fa-117">Minimum</span></span><br><span data-ttu-id="603fa-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="603fa-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="603fa-119">context</span><span class="sxs-lookup"><span data-stu-id="603fa-119">context</span></span>](office.context.md) | <span data-ttu-id="603fa-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="603fa-120">Compose</span></span><br><span data-ttu-id="603fa-121">Ler</span><span class="sxs-lookup"><span data-stu-id="603fa-121">Read</span></span> | [<span data-ttu-id="603fa-122">Context</span><span class="sxs-lookup"><span data-stu-id="603fa-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="603fa-123">1.1</span><span class="sxs-lookup"><span data-stu-id="603fa-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="603fa-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="603fa-124">Enumerations</span></span>

| <span data-ttu-id="603fa-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="603fa-125">Enumeration</span></span> | <span data-ttu-id="603fa-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="603fa-126">Modes</span></span> | <span data-ttu-id="603fa-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="603fa-127">Return type</span></span> | <span data-ttu-id="603fa-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="603fa-128">Minimum</span></span><br><span data-ttu-id="603fa-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="603fa-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="603fa-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="603fa-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="603fa-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="603fa-131">Compose</span></span><br><span data-ttu-id="603fa-132">Ler</span><span class="sxs-lookup"><span data-stu-id="603fa-132">Read</span></span> | <span data-ttu-id="603fa-133">String</span><span class="sxs-lookup"><span data-stu-id="603fa-133">String</span></span> | [<span data-ttu-id="603fa-134">1.1</span><span class="sxs-lookup"><span data-stu-id="603fa-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="603fa-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="603fa-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="603fa-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="603fa-136">Compose</span></span><br><span data-ttu-id="603fa-137">Ler</span><span class="sxs-lookup"><span data-stu-id="603fa-137">Read</span></span> | <span data-ttu-id="603fa-138">String</span><span class="sxs-lookup"><span data-stu-id="603fa-138">String</span></span> | [<span data-ttu-id="603fa-139">1.1</span><span class="sxs-lookup"><span data-stu-id="603fa-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="603fa-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="603fa-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="603fa-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="603fa-141">Compose</span></span><br><span data-ttu-id="603fa-142">Ler</span><span class="sxs-lookup"><span data-stu-id="603fa-142">Read</span></span> | <span data-ttu-id="603fa-143">String</span><span class="sxs-lookup"><span data-stu-id="603fa-143">String</span></span> | [<span data-ttu-id="603fa-144">1.1</span><span class="sxs-lookup"><span data-stu-id="603fa-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="603fa-145">Namespaces</span><span class="sxs-lookup"><span data-stu-id="603fa-145">Namespaces</span></span>

<span data-ttu-id="603fa-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="603fa-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="603fa-147">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="603fa-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="603fa-148">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="603fa-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="603fa-149">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="603fa-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="603fa-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="603fa-150">Type</span></span>

*   <span data-ttu-id="603fa-151">String</span><span class="sxs-lookup"><span data-stu-id="603fa-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="603fa-152">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="603fa-152">Properties:</span></span>

|<span data-ttu-id="603fa-153">Nome</span><span class="sxs-lookup"><span data-stu-id="603fa-153">Name</span></span>| <span data-ttu-id="603fa-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="603fa-154">Type</span></span>| <span data-ttu-id="603fa-155">Descrição</span><span class="sxs-lookup"><span data-stu-id="603fa-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="603fa-156">String</span><span class="sxs-lookup"><span data-stu-id="603fa-156">String</span></span>|<span data-ttu-id="603fa-157">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="603fa-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="603fa-158">String</span><span class="sxs-lookup"><span data-stu-id="603fa-158">String</span></span>|<span data-ttu-id="603fa-159">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="603fa-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="603fa-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="603fa-160">Requirements</span></span>

|<span data-ttu-id="603fa-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="603fa-161">Requirement</span></span>| <span data-ttu-id="603fa-162">Valor</span><span class="sxs-lookup"><span data-stu-id="603fa-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="603fa-163">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="603fa-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="603fa-164">1.1</span><span class="sxs-lookup"><span data-stu-id="603fa-164">1.1</span></span>|
|[<span data-ttu-id="603fa-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="603fa-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="603fa-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="603fa-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="603fa-167">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="603fa-167">CoercionType: String</span></span>

<span data-ttu-id="603fa-168">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="603fa-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="603fa-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="603fa-169">Type</span></span>

*   <span data-ttu-id="603fa-170">String</span><span class="sxs-lookup"><span data-stu-id="603fa-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="603fa-171">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="603fa-171">Properties:</span></span>

|<span data-ttu-id="603fa-172">Nome</span><span class="sxs-lookup"><span data-stu-id="603fa-172">Name</span></span>| <span data-ttu-id="603fa-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="603fa-173">Type</span></span>| <span data-ttu-id="603fa-174">Descrição</span><span class="sxs-lookup"><span data-stu-id="603fa-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="603fa-175">String</span><span class="sxs-lookup"><span data-stu-id="603fa-175">String</span></span>|<span data-ttu-id="603fa-176">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="603fa-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="603fa-177">String</span><span class="sxs-lookup"><span data-stu-id="603fa-177">String</span></span>|<span data-ttu-id="603fa-178">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="603fa-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="603fa-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="603fa-179">Requirements</span></span>

|<span data-ttu-id="603fa-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="603fa-180">Requirement</span></span>| <span data-ttu-id="603fa-181">Valor</span><span class="sxs-lookup"><span data-stu-id="603fa-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="603fa-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="603fa-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="603fa-183">1.1</span><span class="sxs-lookup"><span data-stu-id="603fa-183">1.1</span></span>|
|[<span data-ttu-id="603fa-184">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="603fa-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="603fa-185">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="603fa-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="603fa-186">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="603fa-186">SourceProperty: String</span></span>

<span data-ttu-id="603fa-187">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="603fa-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="603fa-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="603fa-188">Type</span></span>

*   <span data-ttu-id="603fa-189">String</span><span class="sxs-lookup"><span data-stu-id="603fa-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="603fa-190">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="603fa-190">Properties:</span></span>

|<span data-ttu-id="603fa-191">Nome</span><span class="sxs-lookup"><span data-stu-id="603fa-191">Name</span></span>| <span data-ttu-id="603fa-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="603fa-192">Type</span></span>| <span data-ttu-id="603fa-193">Descrição</span><span class="sxs-lookup"><span data-stu-id="603fa-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="603fa-194">String</span><span class="sxs-lookup"><span data-stu-id="603fa-194">String</span></span>|<span data-ttu-id="603fa-195">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="603fa-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="603fa-196">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="603fa-196">String</span></span>|<span data-ttu-id="603fa-197">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="603fa-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="603fa-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="603fa-198">Requirements</span></span>

|<span data-ttu-id="603fa-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="603fa-199">Requirement</span></span>| <span data-ttu-id="603fa-200">Valor</span><span class="sxs-lookup"><span data-stu-id="603fa-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="603fa-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="603fa-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="603fa-202">1.1</span><span class="sxs-lookup"><span data-stu-id="603fa-202">1.1</span></span>|
|[<span data-ttu-id="603fa-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="603fa-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="603fa-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="603fa-204">Compose or Read</span></span>|
