---
title: Namespace do Office – conjunto de requisitos 1,6
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,6.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: b0d1643727055c6b7ddb4d03c0488b82b24f3fad
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611453"
---
# <a name="office-mailbox-requirement-set-16"></a><span data-ttu-id="1480c-103">Office (conjunto de requisitos de caixa de correio 1,6)</span><span class="sxs-lookup"><span data-stu-id="1480c-103">Office (Mailbox requirement set 1.6)</span></span>

<span data-ttu-id="1480c-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="1480c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1480c-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1480c-106">Requirements</span></span>

|<span data-ttu-id="1480c-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="1480c-107">Requirement</span></span>| <span data-ttu-id="1480c-108">Valor</span><span class="sxs-lookup"><span data-stu-id="1480c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1480c-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1480c-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1480c-110">1.1</span><span class="sxs-lookup"><span data-stu-id="1480c-110">1.1</span></span>|
|[<span data-ttu-id="1480c-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1480c-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1480c-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1480c-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="1480c-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="1480c-113">Properties</span></span>

| <span data-ttu-id="1480c-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="1480c-114">Property</span></span> | <span data-ttu-id="1480c-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="1480c-115">Modes</span></span> | <span data-ttu-id="1480c-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="1480c-116">Return type</span></span> | <span data-ttu-id="1480c-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="1480c-117">Minimum</span></span><br><span data-ttu-id="1480c-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="1480c-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1480c-119">context</span><span class="sxs-lookup"><span data-stu-id="1480c-119">context</span></span>](office.context.md) | <span data-ttu-id="1480c-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="1480c-120">Compose</span></span><br><span data-ttu-id="1480c-121">Read</span><span class="sxs-lookup"><span data-stu-id="1480c-121">Read</span></span> | [<span data-ttu-id="1480c-122">Context</span><span class="sxs-lookup"><span data-stu-id="1480c-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="1480c-123">1.1</span><span class="sxs-lookup"><span data-stu-id="1480c-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="1480c-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="1480c-124">Enumerations</span></span>

| <span data-ttu-id="1480c-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="1480c-125">Enumeration</span></span> | <span data-ttu-id="1480c-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="1480c-126">Modes</span></span> | <span data-ttu-id="1480c-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="1480c-127">Return type</span></span> | <span data-ttu-id="1480c-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="1480c-128">Minimum</span></span><br><span data-ttu-id="1480c-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="1480c-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1480c-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="1480c-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="1480c-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="1480c-131">Compose</span></span><br><span data-ttu-id="1480c-132">Read</span><span class="sxs-lookup"><span data-stu-id="1480c-132">Read</span></span> | <span data-ttu-id="1480c-133">String</span><span class="sxs-lookup"><span data-stu-id="1480c-133">String</span></span> | [<span data-ttu-id="1480c-134">1.1</span><span class="sxs-lookup"><span data-stu-id="1480c-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1480c-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="1480c-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="1480c-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="1480c-136">Compose</span></span><br><span data-ttu-id="1480c-137">Read</span><span class="sxs-lookup"><span data-stu-id="1480c-137">Read</span></span> | <span data-ttu-id="1480c-138">String</span><span class="sxs-lookup"><span data-stu-id="1480c-138">String</span></span> | [<span data-ttu-id="1480c-139">1.1</span><span class="sxs-lookup"><span data-stu-id="1480c-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1480c-140">EventType</span><span class="sxs-lookup"><span data-stu-id="1480c-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="1480c-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="1480c-141">Compose</span></span><br><span data-ttu-id="1480c-142">Read</span><span class="sxs-lookup"><span data-stu-id="1480c-142">Read</span></span> | <span data-ttu-id="1480c-143">String</span><span class="sxs-lookup"><span data-stu-id="1480c-143">String</span></span> | [<span data-ttu-id="1480c-144">1,5</span><span class="sxs-lookup"><span data-stu-id="1480c-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="1480c-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="1480c-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="1480c-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="1480c-146">Compose</span></span><br><span data-ttu-id="1480c-147">Read</span><span class="sxs-lookup"><span data-stu-id="1480c-147">Read</span></span> | <span data-ttu-id="1480c-148">String</span><span class="sxs-lookup"><span data-stu-id="1480c-148">String</span></span> | [<span data-ttu-id="1480c-149">1.1</span><span class="sxs-lookup"><span data-stu-id="1480c-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="1480c-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="1480c-150">Namespaces</span></span>

<span data-ttu-id="1480c-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): inclui uma série de enumerações específicas do Outlook, por exemplo,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` e `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="1480c-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="1480c-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="1480c-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="1480c-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1480c-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="1480c-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="1480c-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="1480c-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="1480c-155">Type</span></span>

*   <span data-ttu-id="1480c-156">String</span><span class="sxs-lookup"><span data-stu-id="1480c-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1480c-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1480c-157">Properties:</span></span>

|<span data-ttu-id="1480c-158">Nome</span><span class="sxs-lookup"><span data-stu-id="1480c-158">Name</span></span>| <span data-ttu-id="1480c-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="1480c-159">Type</span></span>| <span data-ttu-id="1480c-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="1480c-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="1480c-161">String</span><span class="sxs-lookup"><span data-stu-id="1480c-161">String</span></span>|<span data-ttu-id="1480c-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="1480c-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="1480c-163">String</span><span class="sxs-lookup"><span data-stu-id="1480c-163">String</span></span>|<span data-ttu-id="1480c-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="1480c-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1480c-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1480c-165">Requirements</span></span>

|<span data-ttu-id="1480c-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="1480c-166">Requirement</span></span>| <span data-ttu-id="1480c-167">Valor</span><span class="sxs-lookup"><span data-stu-id="1480c-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="1480c-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1480c-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1480c-169">1.1</span><span class="sxs-lookup"><span data-stu-id="1480c-169">1.1</span></span>|
|[<span data-ttu-id="1480c-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1480c-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1480c-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1480c-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="1480c-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1480c-172">CoercionType: String</span></span>

<span data-ttu-id="1480c-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="1480c-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1480c-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="1480c-174">Type</span></span>

*   <span data-ttu-id="1480c-175">String</span><span class="sxs-lookup"><span data-stu-id="1480c-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1480c-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1480c-176">Properties:</span></span>

|<span data-ttu-id="1480c-177">Nome</span><span class="sxs-lookup"><span data-stu-id="1480c-177">Name</span></span>| <span data-ttu-id="1480c-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="1480c-178">Type</span></span>| <span data-ttu-id="1480c-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="1480c-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="1480c-180">String</span><span class="sxs-lookup"><span data-stu-id="1480c-180">String</span></span>|<span data-ttu-id="1480c-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="1480c-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="1480c-182">String</span><span class="sxs-lookup"><span data-stu-id="1480c-182">String</span></span>|<span data-ttu-id="1480c-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="1480c-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1480c-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1480c-184">Requirements</span></span>

|<span data-ttu-id="1480c-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="1480c-185">Requirement</span></span>| <span data-ttu-id="1480c-186">Valor</span><span class="sxs-lookup"><span data-stu-id="1480c-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="1480c-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1480c-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1480c-188">1.1</span><span class="sxs-lookup"><span data-stu-id="1480c-188">1.1</span></span>|
|[<span data-ttu-id="1480c-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1480c-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1480c-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1480c-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="1480c-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1480c-191">EventType: String</span></span>

<span data-ttu-id="1480c-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="1480c-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="1480c-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="1480c-193">Type</span></span>

*   <span data-ttu-id="1480c-194">String</span><span class="sxs-lookup"><span data-stu-id="1480c-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1480c-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1480c-195">Properties:</span></span>

| <span data-ttu-id="1480c-196">Nome</span><span class="sxs-lookup"><span data-stu-id="1480c-196">Name</span></span> | <span data-ttu-id="1480c-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="1480c-197">Type</span></span> | <span data-ttu-id="1480c-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="1480c-198">Description</span></span> | <span data-ttu-id="1480c-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="1480c-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="1480c-200">String</span><span class="sxs-lookup"><span data-stu-id="1480c-200">String</span></span> | <span data-ttu-id="1480c-201">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="1480c-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="1480c-202">1,5</span><span class="sxs-lookup"><span data-stu-id="1480c-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1480c-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1480c-203">Requirements</span></span>

|<span data-ttu-id="1480c-204">Requisito</span><span class="sxs-lookup"><span data-stu-id="1480c-204">Requirement</span></span>| <span data-ttu-id="1480c-205">Valor</span><span class="sxs-lookup"><span data-stu-id="1480c-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="1480c-206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1480c-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1480c-207">1,5</span><span class="sxs-lookup"><span data-stu-id="1480c-207">1.5</span></span> |
|[<span data-ttu-id="1480c-208">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1480c-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1480c-209">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1480c-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="1480c-210">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1480c-210">SourceProperty: String</span></span>

<span data-ttu-id="1480c-211">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="1480c-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1480c-212">Tipo</span><span class="sxs-lookup"><span data-stu-id="1480c-212">Type</span></span>

*   <span data-ttu-id="1480c-213">String</span><span class="sxs-lookup"><span data-stu-id="1480c-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1480c-214">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1480c-214">Properties:</span></span>

|<span data-ttu-id="1480c-215">Nome</span><span class="sxs-lookup"><span data-stu-id="1480c-215">Name</span></span>| <span data-ttu-id="1480c-216">Tipo</span><span class="sxs-lookup"><span data-stu-id="1480c-216">Type</span></span>| <span data-ttu-id="1480c-217">Descrição</span><span class="sxs-lookup"><span data-stu-id="1480c-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="1480c-218">String</span><span class="sxs-lookup"><span data-stu-id="1480c-218">String</span></span>|<span data-ttu-id="1480c-219">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1480c-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="1480c-220">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1480c-220">String</span></span>|<span data-ttu-id="1480c-221">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1480c-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1480c-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1480c-222">Requirements</span></span>

|<span data-ttu-id="1480c-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="1480c-223">Requirement</span></span>| <span data-ttu-id="1480c-224">Valor</span><span class="sxs-lookup"><span data-stu-id="1480c-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="1480c-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1480c-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1480c-226">1.1</span><span class="sxs-lookup"><span data-stu-id="1480c-226">1.1</span></span>|
|[<span data-ttu-id="1480c-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1480c-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1480c-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1480c-228">Compose or Read</span></span>|
