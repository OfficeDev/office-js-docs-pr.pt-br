---
title: Namespace do Office – conjunto de requisitos 1.2
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,2.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 0dfde315cb71642a995b4c07a1966d3dee3c0d50
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431301"
---
# <a name="office-mailbox-requirement-set-12"></a><span data-ttu-id="78539-103">Office (conjunto de requisitos de caixa de correio 1,2)</span><span class="sxs-lookup"><span data-stu-id="78539-103">Office (Mailbox requirement set 1.2)</span></span>

<span data-ttu-id="78539-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="78539-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="78539-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78539-106">Requirements</span></span>

|<span data-ttu-id="78539-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="78539-107">Requirement</span></span>| <span data-ttu-id="78539-108">Valor</span><span class="sxs-lookup"><span data-stu-id="78539-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="78539-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78539-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="78539-110">1.1</span><span class="sxs-lookup"><span data-stu-id="78539-110">1.1</span></span>|
|[<span data-ttu-id="78539-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78539-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="78539-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78539-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="78539-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="78539-113">Properties</span></span>

| <span data-ttu-id="78539-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="78539-114">Property</span></span> | <span data-ttu-id="78539-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="78539-115">Modes</span></span> | <span data-ttu-id="78539-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="78539-116">Return type</span></span> | <span data-ttu-id="78539-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="78539-117">Minimum</span></span><br><span data-ttu-id="78539-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="78539-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="78539-119">context</span><span class="sxs-lookup"><span data-stu-id="78539-119">context</span></span>](office.context.md) | <span data-ttu-id="78539-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="78539-120">Compose</span></span><br><span data-ttu-id="78539-121">Ler</span><span class="sxs-lookup"><span data-stu-id="78539-121">Read</span></span> | [<span data-ttu-id="78539-122">Context</span><span class="sxs-lookup"><span data-stu-id="78539-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="78539-123">1.1</span><span class="sxs-lookup"><span data-stu-id="78539-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="78539-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="78539-124">Enumerations</span></span>

| <span data-ttu-id="78539-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="78539-125">Enumeration</span></span> | <span data-ttu-id="78539-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="78539-126">Modes</span></span> | <span data-ttu-id="78539-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="78539-127">Return type</span></span> | <span data-ttu-id="78539-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="78539-128">Minimum</span></span><br><span data-ttu-id="78539-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="78539-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="78539-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="78539-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="78539-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="78539-131">Compose</span></span><br><span data-ttu-id="78539-132">Ler</span><span class="sxs-lookup"><span data-stu-id="78539-132">Read</span></span> | <span data-ttu-id="78539-133">String</span><span class="sxs-lookup"><span data-stu-id="78539-133">String</span></span> | [<span data-ttu-id="78539-134">1.1</span><span class="sxs-lookup"><span data-stu-id="78539-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="78539-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="78539-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="78539-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="78539-136">Compose</span></span><br><span data-ttu-id="78539-137">Ler</span><span class="sxs-lookup"><span data-stu-id="78539-137">Read</span></span> | <span data-ttu-id="78539-138">String</span><span class="sxs-lookup"><span data-stu-id="78539-138">String</span></span> | [<span data-ttu-id="78539-139">1.1</span><span class="sxs-lookup"><span data-stu-id="78539-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="78539-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="78539-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="78539-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="78539-141">Compose</span></span><br><span data-ttu-id="78539-142">Ler</span><span class="sxs-lookup"><span data-stu-id="78539-142">Read</span></span> | <span data-ttu-id="78539-143">String</span><span class="sxs-lookup"><span data-stu-id="78539-143">String</span></span> | [<span data-ttu-id="78539-144">1.1</span><span class="sxs-lookup"><span data-stu-id="78539-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="78539-145">Namespaces</span><span class="sxs-lookup"><span data-stu-id="78539-145">Namespaces</span></span>

<span data-ttu-id="78539-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): inclui uma série de enumerações específicas do Outlook, por exemplo,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` e `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="78539-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="78539-147">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="78539-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="78539-148">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78539-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="78539-149">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="78539-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="78539-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="78539-150">Type</span></span>

*   <span data-ttu-id="78539-151">String</span><span class="sxs-lookup"><span data-stu-id="78539-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="78539-152">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="78539-152">Properties:</span></span>

|<span data-ttu-id="78539-153">Nome</span><span class="sxs-lookup"><span data-stu-id="78539-153">Name</span></span>| <span data-ttu-id="78539-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="78539-154">Type</span></span>| <span data-ttu-id="78539-155">Descrição</span><span class="sxs-lookup"><span data-stu-id="78539-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="78539-156">String</span><span class="sxs-lookup"><span data-stu-id="78539-156">String</span></span>|<span data-ttu-id="78539-157">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="78539-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="78539-158">String</span><span class="sxs-lookup"><span data-stu-id="78539-158">String</span></span>|<span data-ttu-id="78539-159">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="78539-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78539-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78539-160">Requirements</span></span>

|<span data-ttu-id="78539-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="78539-161">Requirement</span></span>| <span data-ttu-id="78539-162">Valor</span><span class="sxs-lookup"><span data-stu-id="78539-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="78539-163">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78539-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="78539-164">1.1</span><span class="sxs-lookup"><span data-stu-id="78539-164">1.1</span></span>|
|[<span data-ttu-id="78539-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78539-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="78539-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78539-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="78539-167">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78539-167">CoercionType: String</span></span>

<span data-ttu-id="78539-168">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="78539-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="78539-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="78539-169">Type</span></span>

*   <span data-ttu-id="78539-170">String</span><span class="sxs-lookup"><span data-stu-id="78539-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="78539-171">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="78539-171">Properties:</span></span>

|<span data-ttu-id="78539-172">Nome</span><span class="sxs-lookup"><span data-stu-id="78539-172">Name</span></span>| <span data-ttu-id="78539-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="78539-173">Type</span></span>| <span data-ttu-id="78539-174">Descrição</span><span class="sxs-lookup"><span data-stu-id="78539-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="78539-175">String</span><span class="sxs-lookup"><span data-stu-id="78539-175">String</span></span>|<span data-ttu-id="78539-176">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="78539-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="78539-177">String</span><span class="sxs-lookup"><span data-stu-id="78539-177">String</span></span>|<span data-ttu-id="78539-178">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="78539-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78539-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78539-179">Requirements</span></span>

|<span data-ttu-id="78539-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="78539-180">Requirement</span></span>| <span data-ttu-id="78539-181">Valor</span><span class="sxs-lookup"><span data-stu-id="78539-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="78539-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78539-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="78539-183">1.1</span><span class="sxs-lookup"><span data-stu-id="78539-183">1.1</span></span>|
|[<span data-ttu-id="78539-184">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78539-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="78539-185">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78539-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="78539-186">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="78539-186">SourceProperty: String</span></span>

<span data-ttu-id="78539-187">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="78539-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="78539-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="78539-188">Type</span></span>

*   <span data-ttu-id="78539-189">String</span><span class="sxs-lookup"><span data-stu-id="78539-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="78539-190">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="78539-190">Properties:</span></span>

|<span data-ttu-id="78539-191">Nome</span><span class="sxs-lookup"><span data-stu-id="78539-191">Name</span></span>| <span data-ttu-id="78539-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="78539-192">Type</span></span>| <span data-ttu-id="78539-193">Descrição</span><span class="sxs-lookup"><span data-stu-id="78539-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="78539-194">String</span><span class="sxs-lookup"><span data-stu-id="78539-194">String</span></span>|<span data-ttu-id="78539-195">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="78539-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="78539-196">String</span><span class="sxs-lookup"><span data-stu-id="78539-196">String</span></span>|<span data-ttu-id="78539-197">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="78539-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="78539-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="78539-198">Requirements</span></span>

|<span data-ttu-id="78539-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="78539-199">Requirement</span></span>| <span data-ttu-id="78539-200">Valor</span><span class="sxs-lookup"><span data-stu-id="78539-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="78539-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="78539-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="78539-202">1.1</span><span class="sxs-lookup"><span data-stu-id="78539-202">1.1</span></span>|
|[<span data-ttu-id="78539-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="78539-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="78539-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="78539-204">Compose or Read</span></span>|
