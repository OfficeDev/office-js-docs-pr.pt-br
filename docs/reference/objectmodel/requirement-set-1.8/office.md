---
title: Namespace do Office – conjunto de requisitos 1,8
description: O namespace do Office fornece interfaces compartilhadas para suplementos do Outlook Office (conjunto de requisitos 1,8)
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0bbe212b0b8e5dc1348cb5cdc03509c44a716d1a
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717499"
---
# <a name="office"></a><span data-ttu-id="3161f-103">Office</span><span class="sxs-lookup"><span data-stu-id="3161f-103">Office</span></span>

<span data-ttu-id="3161f-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="3161f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3161f-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3161f-106">Requirements</span></span>

|<span data-ttu-id="3161f-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="3161f-107">Requirement</span></span>| <span data-ttu-id="3161f-108">Valor</span><span class="sxs-lookup"><span data-stu-id="3161f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="3161f-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3161f-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3161f-110">1.1</span><span class="sxs-lookup"><span data-stu-id="3161f-110">1.1</span></span>|
|[<span data-ttu-id="3161f-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3161f-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3161f-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3161f-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="3161f-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="3161f-113">Properties</span></span>

| <span data-ttu-id="3161f-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="3161f-114">Property</span></span> | <span data-ttu-id="3161f-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="3161f-115">Modes</span></span> | <span data-ttu-id="3161f-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="3161f-116">Return type</span></span> | <span data-ttu-id="3161f-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="3161f-117">Minimum</span></span><br><span data-ttu-id="3161f-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="3161f-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3161f-119">context</span><span class="sxs-lookup"><span data-stu-id="3161f-119">context</span></span>](office.context.md) | <span data-ttu-id="3161f-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="3161f-120">Compose</span></span><br><span data-ttu-id="3161f-121">Ler</span><span class="sxs-lookup"><span data-stu-id="3161f-121">Read</span></span> | [<span data-ttu-id="3161f-122">Context</span><span class="sxs-lookup"><span data-stu-id="3161f-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="3161f-123">1.1</span><span class="sxs-lookup"><span data-stu-id="3161f-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="3161f-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="3161f-124">Enumerations</span></span>

| <span data-ttu-id="3161f-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="3161f-125">Enumeration</span></span> | <span data-ttu-id="3161f-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="3161f-126">Modes</span></span> | <span data-ttu-id="3161f-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="3161f-127">Return type</span></span> | <span data-ttu-id="3161f-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="3161f-128">Minimum</span></span><br><span data-ttu-id="3161f-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="3161f-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3161f-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3161f-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3161f-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="3161f-131">Compose</span></span><br><span data-ttu-id="3161f-132">Ler</span><span class="sxs-lookup"><span data-stu-id="3161f-132">Read</span></span> | <span data-ttu-id="3161f-133">String</span><span class="sxs-lookup"><span data-stu-id="3161f-133">String</span></span> | [<span data-ttu-id="3161f-134">1.1</span><span class="sxs-lookup"><span data-stu-id="3161f-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3161f-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3161f-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3161f-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="3161f-136">Compose</span></span><br><span data-ttu-id="3161f-137">Ler</span><span class="sxs-lookup"><span data-stu-id="3161f-137">Read</span></span> | <span data-ttu-id="3161f-138">String</span><span class="sxs-lookup"><span data-stu-id="3161f-138">String</span></span> | [<span data-ttu-id="3161f-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3161f-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3161f-140">EventType</span><span class="sxs-lookup"><span data-stu-id="3161f-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="3161f-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="3161f-141">Compose</span></span><br><span data-ttu-id="3161f-142">Ler</span><span class="sxs-lookup"><span data-stu-id="3161f-142">Read</span></span> | <span data-ttu-id="3161f-143">String</span><span class="sxs-lookup"><span data-stu-id="3161f-143">String</span></span> | [<span data-ttu-id="3161f-144">1,5</span><span class="sxs-lookup"><span data-stu-id="3161f-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="3161f-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3161f-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3161f-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="3161f-146">Compose</span></span><br><span data-ttu-id="3161f-147">Ler</span><span class="sxs-lookup"><span data-stu-id="3161f-147">Read</span></span> | <span data-ttu-id="3161f-148">String</span><span class="sxs-lookup"><span data-stu-id="3161f-148">String</span></span> | [<span data-ttu-id="3161f-149">1.1</span><span class="sxs-lookup"><span data-stu-id="3161f-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="3161f-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="3161f-150">Namespaces</span></span>

<span data-ttu-id="3161f-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="3161f-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="3161f-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="3161f-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="3161f-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3161f-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="3161f-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="3161f-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3161f-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="3161f-155">Type</span></span>

*   <span data-ttu-id="3161f-156">String</span><span class="sxs-lookup"><span data-stu-id="3161f-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3161f-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3161f-157">Properties:</span></span>

|<span data-ttu-id="3161f-158">Nome</span><span class="sxs-lookup"><span data-stu-id="3161f-158">Name</span></span>| <span data-ttu-id="3161f-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="3161f-159">Type</span></span>| <span data-ttu-id="3161f-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="3161f-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3161f-161">String</span><span class="sxs-lookup"><span data-stu-id="3161f-161">String</span></span>|<span data-ttu-id="3161f-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="3161f-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3161f-163">String</span><span class="sxs-lookup"><span data-stu-id="3161f-163">String</span></span>|<span data-ttu-id="3161f-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="3161f-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3161f-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3161f-165">Requirements</span></span>

|<span data-ttu-id="3161f-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="3161f-166">Requirement</span></span>| <span data-ttu-id="3161f-167">Valor</span><span class="sxs-lookup"><span data-stu-id="3161f-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="3161f-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3161f-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3161f-169">1.1</span><span class="sxs-lookup"><span data-stu-id="3161f-169">1.1</span></span>|
|[<span data-ttu-id="3161f-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3161f-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3161f-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3161f-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="3161f-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3161f-172">CoercionType: String</span></span>

<span data-ttu-id="3161f-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="3161f-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3161f-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="3161f-174">Type</span></span>

*   <span data-ttu-id="3161f-175">String</span><span class="sxs-lookup"><span data-stu-id="3161f-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3161f-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3161f-176">Properties:</span></span>

|<span data-ttu-id="3161f-177">Nome</span><span class="sxs-lookup"><span data-stu-id="3161f-177">Name</span></span>| <span data-ttu-id="3161f-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="3161f-178">Type</span></span>| <span data-ttu-id="3161f-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="3161f-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3161f-180">String</span><span class="sxs-lookup"><span data-stu-id="3161f-180">String</span></span>|<span data-ttu-id="3161f-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="3161f-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3161f-182">String</span><span class="sxs-lookup"><span data-stu-id="3161f-182">String</span></span>|<span data-ttu-id="3161f-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="3161f-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3161f-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3161f-184">Requirements</span></span>

|<span data-ttu-id="3161f-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="3161f-185">Requirement</span></span>| <span data-ttu-id="3161f-186">Valor</span><span class="sxs-lookup"><span data-stu-id="3161f-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="3161f-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3161f-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3161f-188">1.1</span><span class="sxs-lookup"><span data-stu-id="3161f-188">1.1</span></span>|
|[<span data-ttu-id="3161f-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3161f-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3161f-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3161f-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="3161f-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3161f-191">EventType: String</span></span>

<span data-ttu-id="3161f-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="3161f-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="3161f-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="3161f-193">Type</span></span>

*   <span data-ttu-id="3161f-194">String</span><span class="sxs-lookup"><span data-stu-id="3161f-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3161f-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3161f-195">Properties:</span></span>

| <span data-ttu-id="3161f-196">Nome</span><span class="sxs-lookup"><span data-stu-id="3161f-196">Name</span></span> | <span data-ttu-id="3161f-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="3161f-197">Type</span></span> | <span data-ttu-id="3161f-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="3161f-198">Description</span></span> | <span data-ttu-id="3161f-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="3161f-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="3161f-200">String</span><span class="sxs-lookup"><span data-stu-id="3161f-200">String</span></span> | <span data-ttu-id="3161f-201">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="3161f-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="3161f-202">1.7</span><span class="sxs-lookup"><span data-stu-id="3161f-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="3161f-203">String</span><span class="sxs-lookup"><span data-stu-id="3161f-203">String</span></span> | <span data-ttu-id="3161f-204">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="3161f-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="3161f-205">1,8</span><span class="sxs-lookup"><span data-stu-id="3161f-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="3161f-206">String</span><span class="sxs-lookup"><span data-stu-id="3161f-206">String</span></span> | <span data-ttu-id="3161f-207">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="3161f-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="3161f-208">1,8</span><span class="sxs-lookup"><span data-stu-id="3161f-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="3161f-209">String</span><span class="sxs-lookup"><span data-stu-id="3161f-209">String</span></span> | <span data-ttu-id="3161f-210">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="3161f-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="3161f-211">1,5</span><span class="sxs-lookup"><span data-stu-id="3161f-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="3161f-212">String</span><span class="sxs-lookup"><span data-stu-id="3161f-212">String</span></span> | <span data-ttu-id="3161f-213">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="3161f-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="3161f-214">1.7</span><span class="sxs-lookup"><span data-stu-id="3161f-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="3161f-215">String</span><span class="sxs-lookup"><span data-stu-id="3161f-215">String</span></span> | <span data-ttu-id="3161f-216">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="3161f-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="3161f-217">1.7</span><span class="sxs-lookup"><span data-stu-id="3161f-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3161f-218">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3161f-218">Requirements</span></span>

|<span data-ttu-id="3161f-219">Requisito</span><span class="sxs-lookup"><span data-stu-id="3161f-219">Requirement</span></span>| <span data-ttu-id="3161f-220">Valor</span><span class="sxs-lookup"><span data-stu-id="3161f-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="3161f-221">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3161f-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3161f-222">1,5</span><span class="sxs-lookup"><span data-stu-id="3161f-222">1.5</span></span> |
|[<span data-ttu-id="3161f-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3161f-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3161f-224">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3161f-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="3161f-225">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3161f-225">SourceProperty: String</span></span>

<span data-ttu-id="3161f-226">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="3161f-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3161f-227">Tipo</span><span class="sxs-lookup"><span data-stu-id="3161f-227">Type</span></span>

*   <span data-ttu-id="3161f-228">String</span><span class="sxs-lookup"><span data-stu-id="3161f-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3161f-229">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3161f-229">Properties:</span></span>

|<span data-ttu-id="3161f-230">Nome</span><span class="sxs-lookup"><span data-stu-id="3161f-230">Name</span></span>| <span data-ttu-id="3161f-231">Tipo</span><span class="sxs-lookup"><span data-stu-id="3161f-231">Type</span></span>| <span data-ttu-id="3161f-232">Descrição</span><span class="sxs-lookup"><span data-stu-id="3161f-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3161f-233">String</span><span class="sxs-lookup"><span data-stu-id="3161f-233">String</span></span>|<span data-ttu-id="3161f-234">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="3161f-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3161f-235">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3161f-235">String</span></span>|<span data-ttu-id="3161f-236">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="3161f-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3161f-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3161f-237">Requirements</span></span>

|<span data-ttu-id="3161f-238">Requisito</span><span class="sxs-lookup"><span data-stu-id="3161f-238">Requirement</span></span>| <span data-ttu-id="3161f-239">Valor</span><span class="sxs-lookup"><span data-stu-id="3161f-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="3161f-240">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3161f-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3161f-241">1.1</span><span class="sxs-lookup"><span data-stu-id="3161f-241">1.1</span></span>|
|[<span data-ttu-id="3161f-242">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3161f-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3161f-243">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3161f-243">Compose or Read</span></span>|
