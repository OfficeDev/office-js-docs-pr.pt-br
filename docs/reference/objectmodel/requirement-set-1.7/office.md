---
title: Namespace do Office – conjunto de requisitos 1,7
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,7.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 7991fd56097bbdebbfd4d4494a900626a1d3e02b
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891247"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="df9ff-103">Office (conjunto de requisitos de caixa de correio 1,7)</span><span class="sxs-lookup"><span data-stu-id="df9ff-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="df9ff-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="df9ff-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="df9ff-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="df9ff-106">Requirements</span></span>

|<span data-ttu-id="df9ff-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="df9ff-107">Requirement</span></span>| <span data-ttu-id="df9ff-108">Valor</span><span class="sxs-lookup"><span data-stu-id="df9ff-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="df9ff-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="df9ff-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="df9ff-110">1.1</span><span class="sxs-lookup"><span data-stu-id="df9ff-110">1.1</span></span>|
|[<span data-ttu-id="df9ff-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="df9ff-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="df9ff-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="df9ff-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="df9ff-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="df9ff-113">Properties</span></span>

| <span data-ttu-id="df9ff-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="df9ff-114">Property</span></span> | <span data-ttu-id="df9ff-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="df9ff-115">Modes</span></span> | <span data-ttu-id="df9ff-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="df9ff-116">Return type</span></span> | <span data-ttu-id="df9ff-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="df9ff-117">Minimum</span></span><br><span data-ttu-id="df9ff-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="df9ff-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="df9ff-119">context</span><span class="sxs-lookup"><span data-stu-id="df9ff-119">context</span></span>](office.context.md) | <span data-ttu-id="df9ff-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="df9ff-120">Compose</span></span><br><span data-ttu-id="df9ff-121">Ler</span><span class="sxs-lookup"><span data-stu-id="df9ff-121">Read</span></span> | [<span data-ttu-id="df9ff-122">Context</span><span class="sxs-lookup"><span data-stu-id="df9ff-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="df9ff-123">1.1</span><span class="sxs-lookup"><span data-stu-id="df9ff-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="df9ff-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="df9ff-124">Enumerations</span></span>

| <span data-ttu-id="df9ff-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="df9ff-125">Enumeration</span></span> | <span data-ttu-id="df9ff-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="df9ff-126">Modes</span></span> | <span data-ttu-id="df9ff-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="df9ff-127">Return type</span></span> | <span data-ttu-id="df9ff-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="df9ff-128">Minimum</span></span><br><span data-ttu-id="df9ff-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="df9ff-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="df9ff-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="df9ff-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="df9ff-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="df9ff-131">Compose</span></span><br><span data-ttu-id="df9ff-132">Ler</span><span class="sxs-lookup"><span data-stu-id="df9ff-132">Read</span></span> | <span data-ttu-id="df9ff-133">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-133">String</span></span> | [<span data-ttu-id="df9ff-134">1.1</span><span class="sxs-lookup"><span data-stu-id="df9ff-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="df9ff-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="df9ff-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="df9ff-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="df9ff-136">Compose</span></span><br><span data-ttu-id="df9ff-137">Ler</span><span class="sxs-lookup"><span data-stu-id="df9ff-137">Read</span></span> | <span data-ttu-id="df9ff-138">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-138">String</span></span> | [<span data-ttu-id="df9ff-139">1.1</span><span class="sxs-lookup"><span data-stu-id="df9ff-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="df9ff-140">EventType</span><span class="sxs-lookup"><span data-stu-id="df9ff-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="df9ff-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="df9ff-141">Compose</span></span><br><span data-ttu-id="df9ff-142">Ler</span><span class="sxs-lookup"><span data-stu-id="df9ff-142">Read</span></span> | <span data-ttu-id="df9ff-143">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-143">String</span></span> | [<span data-ttu-id="df9ff-144">1,5</span><span class="sxs-lookup"><span data-stu-id="df9ff-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="df9ff-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="df9ff-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="df9ff-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="df9ff-146">Compose</span></span><br><span data-ttu-id="df9ff-147">Ler</span><span class="sxs-lookup"><span data-stu-id="df9ff-147">Read</span></span> | <span data-ttu-id="df9ff-148">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-148">String</span></span> | [<span data-ttu-id="df9ff-149">1.1</span><span class="sxs-lookup"><span data-stu-id="df9ff-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="df9ff-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="df9ff-150">Namespaces</span></span>

<span data-ttu-id="df9ff-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="df9ff-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="df9ff-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="df9ff-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="df9ff-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="df9ff-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="df9ff-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="df9ff-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="df9ff-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="df9ff-155">Type</span></span>

*   <span data-ttu-id="df9ff-156">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="df9ff-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="df9ff-157">Properties:</span></span>

|<span data-ttu-id="df9ff-158">Nome</span><span class="sxs-lookup"><span data-stu-id="df9ff-158">Name</span></span>| <span data-ttu-id="df9ff-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="df9ff-159">Type</span></span>| <span data-ttu-id="df9ff-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="df9ff-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="df9ff-161">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-161">String</span></span>|<span data-ttu-id="df9ff-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="df9ff-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="df9ff-163">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-163">String</span></span>|<span data-ttu-id="df9ff-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="df9ff-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df9ff-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="df9ff-165">Requirements</span></span>

|<span data-ttu-id="df9ff-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="df9ff-166">Requirement</span></span>| <span data-ttu-id="df9ff-167">Valor</span><span class="sxs-lookup"><span data-stu-id="df9ff-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="df9ff-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="df9ff-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="df9ff-169">1.1</span><span class="sxs-lookup"><span data-stu-id="df9ff-169">1.1</span></span>|
|[<span data-ttu-id="df9ff-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="df9ff-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="df9ff-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="df9ff-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="df9ff-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="df9ff-172">CoercionType: String</span></span>

<span data-ttu-id="df9ff-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="df9ff-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="df9ff-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="df9ff-174">Type</span></span>

*   <span data-ttu-id="df9ff-175">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="df9ff-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="df9ff-176">Properties:</span></span>

|<span data-ttu-id="df9ff-177">Nome</span><span class="sxs-lookup"><span data-stu-id="df9ff-177">Name</span></span>| <span data-ttu-id="df9ff-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="df9ff-178">Type</span></span>| <span data-ttu-id="df9ff-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="df9ff-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="df9ff-180">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-180">String</span></span>|<span data-ttu-id="df9ff-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="df9ff-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="df9ff-182">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-182">String</span></span>|<span data-ttu-id="df9ff-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="df9ff-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df9ff-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="df9ff-184">Requirements</span></span>

|<span data-ttu-id="df9ff-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="df9ff-185">Requirement</span></span>| <span data-ttu-id="df9ff-186">Valor</span><span class="sxs-lookup"><span data-stu-id="df9ff-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="df9ff-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="df9ff-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="df9ff-188">1.1</span><span class="sxs-lookup"><span data-stu-id="df9ff-188">1.1</span></span>|
|[<span data-ttu-id="df9ff-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="df9ff-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="df9ff-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="df9ff-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="df9ff-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="df9ff-191">EventType: String</span></span>

<span data-ttu-id="df9ff-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="df9ff-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="df9ff-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="df9ff-193">Type</span></span>

*   <span data-ttu-id="df9ff-194">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="df9ff-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="df9ff-195">Properties:</span></span>

| <span data-ttu-id="df9ff-196">Nome</span><span class="sxs-lookup"><span data-stu-id="df9ff-196">Name</span></span> | <span data-ttu-id="df9ff-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="df9ff-197">Type</span></span> | <span data-ttu-id="df9ff-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="df9ff-198">Description</span></span> | <span data-ttu-id="df9ff-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="df9ff-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="df9ff-200">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-200">String</span></span> | <span data-ttu-id="df9ff-201">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="df9ff-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="df9ff-202">1.7</span><span class="sxs-lookup"><span data-stu-id="df9ff-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="df9ff-203">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-203">String</span></span> | <span data-ttu-id="df9ff-204">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="df9ff-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="df9ff-205">1,5</span><span class="sxs-lookup"><span data-stu-id="df9ff-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="df9ff-206">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-206">String</span></span> | <span data-ttu-id="df9ff-207">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="df9ff-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="df9ff-208">1.7</span><span class="sxs-lookup"><span data-stu-id="df9ff-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="df9ff-209">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-209">String</span></span> | <span data-ttu-id="df9ff-210">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="df9ff-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="df9ff-211">1.7</span><span class="sxs-lookup"><span data-stu-id="df9ff-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="df9ff-212">Requisitos</span><span class="sxs-lookup"><span data-stu-id="df9ff-212">Requirements</span></span>

|<span data-ttu-id="df9ff-213">Requisito</span><span class="sxs-lookup"><span data-stu-id="df9ff-213">Requirement</span></span>| <span data-ttu-id="df9ff-214">Valor</span><span class="sxs-lookup"><span data-stu-id="df9ff-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="df9ff-215">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="df9ff-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="df9ff-216">1,5</span><span class="sxs-lookup"><span data-stu-id="df9ff-216">1.5</span></span> |
|[<span data-ttu-id="df9ff-217">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="df9ff-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="df9ff-218">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="df9ff-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="df9ff-219">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="df9ff-219">SourceProperty: String</span></span>

<span data-ttu-id="df9ff-220">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="df9ff-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="df9ff-221">Tipo</span><span class="sxs-lookup"><span data-stu-id="df9ff-221">Type</span></span>

*   <span data-ttu-id="df9ff-222">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="df9ff-223">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="df9ff-223">Properties:</span></span>

|<span data-ttu-id="df9ff-224">Nome</span><span class="sxs-lookup"><span data-stu-id="df9ff-224">Name</span></span>| <span data-ttu-id="df9ff-225">Tipo</span><span class="sxs-lookup"><span data-stu-id="df9ff-225">Type</span></span>| <span data-ttu-id="df9ff-226">Descrição</span><span class="sxs-lookup"><span data-stu-id="df9ff-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="df9ff-227">String</span><span class="sxs-lookup"><span data-stu-id="df9ff-227">String</span></span>|<span data-ttu-id="df9ff-228">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="df9ff-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="df9ff-229">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="df9ff-229">String</span></span>|<span data-ttu-id="df9ff-230">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="df9ff-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="df9ff-231">Requisitos</span><span class="sxs-lookup"><span data-stu-id="df9ff-231">Requirements</span></span>

|<span data-ttu-id="df9ff-232">Requisito</span><span class="sxs-lookup"><span data-stu-id="df9ff-232">Requirement</span></span>| <span data-ttu-id="df9ff-233">Valor</span><span class="sxs-lookup"><span data-stu-id="df9ff-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="df9ff-234">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="df9ff-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="df9ff-235">1.1</span><span class="sxs-lookup"><span data-stu-id="df9ff-235">1.1</span></span>|
|[<span data-ttu-id="df9ff-236">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="df9ff-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="df9ff-237">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="df9ff-237">Compose or Read</span></span>|
