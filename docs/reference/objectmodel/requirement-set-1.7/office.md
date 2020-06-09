---
title: Namespace do Office – conjunto de requisitos 1,7
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,7.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 718de46689fc2fcb52ad455763581ecab06a4c39
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612196"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="5cc0c-103">Office (conjunto de requisitos de caixa de correio 1,7)</span><span class="sxs-lookup"><span data-stu-id="5cc0c-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="5cc0c-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="5cc0c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5cc0c-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5cc0c-106">Requirements</span></span>

|<span data-ttu-id="5cc0c-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="5cc0c-107">Requirement</span></span>| <span data-ttu-id="5cc0c-108">Valor</span><span class="sxs-lookup"><span data-stu-id="5cc0c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="5cc0c-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5cc0c-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5cc0c-110">1.1</span><span class="sxs-lookup"><span data-stu-id="5cc0c-110">1.1</span></span>|
|[<span data-ttu-id="5cc0c-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5cc0c-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5cc0c-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5cc0c-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="5cc0c-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="5cc0c-113">Properties</span></span>

| <span data-ttu-id="5cc0c-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="5cc0c-114">Property</span></span> | <span data-ttu-id="5cc0c-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="5cc0c-115">Modes</span></span> | <span data-ttu-id="5cc0c-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="5cc0c-116">Return type</span></span> | <span data-ttu-id="5cc0c-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="5cc0c-117">Minimum</span></span><br><span data-ttu-id="5cc0c-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="5cc0c-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="5cc0c-119">context</span><span class="sxs-lookup"><span data-stu-id="5cc0c-119">context</span></span>](office.context.md) | <span data-ttu-id="5cc0c-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="5cc0c-120">Compose</span></span><br><span data-ttu-id="5cc0c-121">Read</span><span class="sxs-lookup"><span data-stu-id="5cc0c-121">Read</span></span> | [<span data-ttu-id="5cc0c-122">Context</span><span class="sxs-lookup"><span data-stu-id="5cc0c-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="5cc0c-123">1.1</span><span class="sxs-lookup"><span data-stu-id="5cc0c-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="5cc0c-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="5cc0c-124">Enumerations</span></span>

| <span data-ttu-id="5cc0c-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="5cc0c-125">Enumeration</span></span> | <span data-ttu-id="5cc0c-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="5cc0c-126">Modes</span></span> | <span data-ttu-id="5cc0c-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="5cc0c-127">Return type</span></span> | <span data-ttu-id="5cc0c-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="5cc0c-128">Minimum</span></span><br><span data-ttu-id="5cc0c-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="5cc0c-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="5cc0c-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="5cc0c-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="5cc0c-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="5cc0c-131">Compose</span></span><br><span data-ttu-id="5cc0c-132">Read</span><span class="sxs-lookup"><span data-stu-id="5cc0c-132">Read</span></span> | <span data-ttu-id="5cc0c-133">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-133">String</span></span> | [<span data-ttu-id="5cc0c-134">1.1</span><span class="sxs-lookup"><span data-stu-id="5cc0c-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5cc0c-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="5cc0c-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="5cc0c-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="5cc0c-136">Compose</span></span><br><span data-ttu-id="5cc0c-137">Read</span><span class="sxs-lookup"><span data-stu-id="5cc0c-137">Read</span></span> | <span data-ttu-id="5cc0c-138">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-138">String</span></span> | [<span data-ttu-id="5cc0c-139">1.1</span><span class="sxs-lookup"><span data-stu-id="5cc0c-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5cc0c-140">EventType</span><span class="sxs-lookup"><span data-stu-id="5cc0c-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="5cc0c-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="5cc0c-141">Compose</span></span><br><span data-ttu-id="5cc0c-142">Read</span><span class="sxs-lookup"><span data-stu-id="5cc0c-142">Read</span></span> | <span data-ttu-id="5cc0c-143">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-143">String</span></span> | [<span data-ttu-id="5cc0c-144">1,5</span><span class="sxs-lookup"><span data-stu-id="5cc0c-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="5cc0c-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="5cc0c-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="5cc0c-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="5cc0c-146">Compose</span></span><br><span data-ttu-id="5cc0c-147">Read</span><span class="sxs-lookup"><span data-stu-id="5cc0c-147">Read</span></span> | <span data-ttu-id="5cc0c-148">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-148">String</span></span> | [<span data-ttu-id="5cc0c-149">1.1</span><span class="sxs-lookup"><span data-stu-id="5cc0c-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="5cc0c-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="5cc0c-150">Namespaces</span></span>

<span data-ttu-id="5cc0c-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): inclui uma série de enumerações específicas do Outlook, por exemplo,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` e `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="5cc0c-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="5cc0c-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="5cc0c-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="5cc0c-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5cc0c-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="5cc0c-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5cc0c-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="5cc0c-155">Type</span></span>

*   <span data-ttu-id="5cc0c-156">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5cc0c-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5cc0c-157">Properties:</span></span>

|<span data-ttu-id="5cc0c-158">Nome</span><span class="sxs-lookup"><span data-stu-id="5cc0c-158">Name</span></span>| <span data-ttu-id="5cc0c-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="5cc0c-159">Type</span></span>| <span data-ttu-id="5cc0c-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="5cc0c-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5cc0c-161">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-161">String</span></span>|<span data-ttu-id="5cc0c-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5cc0c-163">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-163">String</span></span>|<span data-ttu-id="5cc0c-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5cc0c-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5cc0c-165">Requirements</span></span>

|<span data-ttu-id="5cc0c-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="5cc0c-166">Requirement</span></span>| <span data-ttu-id="5cc0c-167">Valor</span><span class="sxs-lookup"><span data-stu-id="5cc0c-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="5cc0c-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5cc0c-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5cc0c-169">1.1</span><span class="sxs-lookup"><span data-stu-id="5cc0c-169">1.1</span></span>|
|[<span data-ttu-id="5cc0c-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5cc0c-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5cc0c-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5cc0c-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="5cc0c-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5cc0c-172">CoercionType: String</span></span>

<span data-ttu-id="5cc0c-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5cc0c-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="5cc0c-174">Type</span></span>

*   <span data-ttu-id="5cc0c-175">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5cc0c-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5cc0c-176">Properties:</span></span>

|<span data-ttu-id="5cc0c-177">Nome</span><span class="sxs-lookup"><span data-stu-id="5cc0c-177">Name</span></span>| <span data-ttu-id="5cc0c-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="5cc0c-178">Type</span></span>| <span data-ttu-id="5cc0c-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="5cc0c-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5cc0c-180">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-180">String</span></span>|<span data-ttu-id="5cc0c-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5cc0c-182">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-182">String</span></span>|<span data-ttu-id="5cc0c-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5cc0c-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5cc0c-184">Requirements</span></span>

|<span data-ttu-id="5cc0c-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="5cc0c-185">Requirement</span></span>| <span data-ttu-id="5cc0c-186">Valor</span><span class="sxs-lookup"><span data-stu-id="5cc0c-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="5cc0c-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5cc0c-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5cc0c-188">1.1</span><span class="sxs-lookup"><span data-stu-id="5cc0c-188">1.1</span></span>|
|[<span data-ttu-id="5cc0c-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5cc0c-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5cc0c-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5cc0c-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="5cc0c-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5cc0c-191">EventType: String</span></span>

<span data-ttu-id="5cc0c-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="5cc0c-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="5cc0c-193">Type</span></span>

*   <span data-ttu-id="5cc0c-194">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5cc0c-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5cc0c-195">Properties:</span></span>

| <span data-ttu-id="5cc0c-196">Nome</span><span class="sxs-lookup"><span data-stu-id="5cc0c-196">Name</span></span> | <span data-ttu-id="5cc0c-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="5cc0c-197">Type</span></span> | <span data-ttu-id="5cc0c-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="5cc0c-198">Description</span></span> | <span data-ttu-id="5cc0c-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="5cc0c-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="5cc0c-200">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-200">String</span></span> | <span data-ttu-id="5cc0c-201">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="5cc0c-202">1.7</span><span class="sxs-lookup"><span data-stu-id="5cc0c-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="5cc0c-203">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-203">String</span></span> | <span data-ttu-id="5cc0c-204">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="5cc0c-205">1,5</span><span class="sxs-lookup"><span data-stu-id="5cc0c-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="5cc0c-206">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-206">String</span></span> | <span data-ttu-id="5cc0c-207">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="5cc0c-208">1.7</span><span class="sxs-lookup"><span data-stu-id="5cc0c-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="5cc0c-209">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-209">String</span></span> | <span data-ttu-id="5cc0c-210">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="5cc0c-211">1.7</span><span class="sxs-lookup"><span data-stu-id="5cc0c-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5cc0c-212">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5cc0c-212">Requirements</span></span>

|<span data-ttu-id="5cc0c-213">Requisito</span><span class="sxs-lookup"><span data-stu-id="5cc0c-213">Requirement</span></span>| <span data-ttu-id="5cc0c-214">Valor</span><span class="sxs-lookup"><span data-stu-id="5cc0c-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="5cc0c-215">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5cc0c-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5cc0c-216">1,5</span><span class="sxs-lookup"><span data-stu-id="5cc0c-216">1.5</span></span> |
|[<span data-ttu-id="5cc0c-217">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5cc0c-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5cc0c-218">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5cc0c-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="5cc0c-219">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5cc0c-219">SourceProperty: String</span></span>

<span data-ttu-id="5cc0c-220">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5cc0c-221">Tipo</span><span class="sxs-lookup"><span data-stu-id="5cc0c-221">Type</span></span>

*   <span data-ttu-id="5cc0c-222">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5cc0c-223">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5cc0c-223">Properties:</span></span>

|<span data-ttu-id="5cc0c-224">Nome</span><span class="sxs-lookup"><span data-stu-id="5cc0c-224">Name</span></span>| <span data-ttu-id="5cc0c-225">Tipo</span><span class="sxs-lookup"><span data-stu-id="5cc0c-225">Type</span></span>| <span data-ttu-id="5cc0c-226">Descrição</span><span class="sxs-lookup"><span data-stu-id="5cc0c-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5cc0c-227">String</span><span class="sxs-lookup"><span data-stu-id="5cc0c-227">String</span></span>|<span data-ttu-id="5cc0c-228">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5cc0c-229">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5cc0c-229">String</span></span>|<span data-ttu-id="5cc0c-230">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5cc0c-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5cc0c-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5cc0c-231">Requirements</span></span>

|<span data-ttu-id="5cc0c-232">Requisito</span><span class="sxs-lookup"><span data-stu-id="5cc0c-232">Requirement</span></span>| <span data-ttu-id="5cc0c-233">Valor</span><span class="sxs-lookup"><span data-stu-id="5cc0c-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="5cc0c-234">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5cc0c-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5cc0c-235">1.1</span><span class="sxs-lookup"><span data-stu-id="5cc0c-235">1.1</span></span>|
|[<span data-ttu-id="5cc0c-236">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5cc0c-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5cc0c-237">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5cc0c-237">Compose or Read</span></span>|
