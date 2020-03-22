---
title: Namespace do Office – conjunto de requisitos 1,8
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,8.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 773a12d2f2b6c2d164b94d0b6b6c2dd0def90a41
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891177"
---
# <a name="office-mailbox-requirement-set-18"></a><span data-ttu-id="95026-103">Office (conjunto de requisitos de caixa de correio 1,8)</span><span class="sxs-lookup"><span data-stu-id="95026-103">Office (Mailbox requirement set 1.8)</span></span>

<span data-ttu-id="95026-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="95026-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="95026-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="95026-106">Requirements</span></span>

|<span data-ttu-id="95026-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="95026-107">Requirement</span></span>| <span data-ttu-id="95026-108">Valor</span><span class="sxs-lookup"><span data-stu-id="95026-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="95026-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="95026-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95026-110">1.1</span><span class="sxs-lookup"><span data-stu-id="95026-110">1.1</span></span>|
|[<span data-ttu-id="95026-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="95026-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95026-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="95026-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="95026-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="95026-113">Properties</span></span>

| <span data-ttu-id="95026-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="95026-114">Property</span></span> | <span data-ttu-id="95026-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="95026-115">Modes</span></span> | <span data-ttu-id="95026-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="95026-116">Return type</span></span> | <span data-ttu-id="95026-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="95026-117">Minimum</span></span><br><span data-ttu-id="95026-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="95026-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="95026-119">context</span><span class="sxs-lookup"><span data-stu-id="95026-119">context</span></span>](office.context.md) | <span data-ttu-id="95026-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="95026-120">Compose</span></span><br><span data-ttu-id="95026-121">Ler</span><span class="sxs-lookup"><span data-stu-id="95026-121">Read</span></span> | [<span data-ttu-id="95026-122">Context</span><span class="sxs-lookup"><span data-stu-id="95026-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="95026-123">1.1</span><span class="sxs-lookup"><span data-stu-id="95026-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="95026-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="95026-124">Enumerations</span></span>

| <span data-ttu-id="95026-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="95026-125">Enumeration</span></span> | <span data-ttu-id="95026-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="95026-126">Modes</span></span> | <span data-ttu-id="95026-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="95026-127">Return type</span></span> | <span data-ttu-id="95026-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="95026-128">Minimum</span></span><br><span data-ttu-id="95026-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="95026-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="95026-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="95026-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="95026-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="95026-131">Compose</span></span><br><span data-ttu-id="95026-132">Ler</span><span class="sxs-lookup"><span data-stu-id="95026-132">Read</span></span> | <span data-ttu-id="95026-133">String</span><span class="sxs-lookup"><span data-stu-id="95026-133">String</span></span> | [<span data-ttu-id="95026-134">1.1</span><span class="sxs-lookup"><span data-stu-id="95026-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="95026-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="95026-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="95026-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="95026-136">Compose</span></span><br><span data-ttu-id="95026-137">Ler</span><span class="sxs-lookup"><span data-stu-id="95026-137">Read</span></span> | <span data-ttu-id="95026-138">String</span><span class="sxs-lookup"><span data-stu-id="95026-138">String</span></span> | [<span data-ttu-id="95026-139">1.1</span><span class="sxs-lookup"><span data-stu-id="95026-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="95026-140">EventType</span><span class="sxs-lookup"><span data-stu-id="95026-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="95026-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="95026-141">Compose</span></span><br><span data-ttu-id="95026-142">Ler</span><span class="sxs-lookup"><span data-stu-id="95026-142">Read</span></span> | <span data-ttu-id="95026-143">String</span><span class="sxs-lookup"><span data-stu-id="95026-143">String</span></span> | [<span data-ttu-id="95026-144">1,5</span><span class="sxs-lookup"><span data-stu-id="95026-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="95026-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="95026-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="95026-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="95026-146">Compose</span></span><br><span data-ttu-id="95026-147">Ler</span><span class="sxs-lookup"><span data-stu-id="95026-147">Read</span></span> | <span data-ttu-id="95026-148">String</span><span class="sxs-lookup"><span data-stu-id="95026-148">String</span></span> | [<span data-ttu-id="95026-149">1.1</span><span class="sxs-lookup"><span data-stu-id="95026-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="95026-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="95026-150">Namespaces</span></span>

<span data-ttu-id="95026-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="95026-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="95026-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="95026-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="95026-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="95026-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="95026-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="95026-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="95026-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="95026-155">Type</span></span>

*   <span data-ttu-id="95026-156">String</span><span class="sxs-lookup"><span data-stu-id="95026-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="95026-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="95026-157">Properties:</span></span>

|<span data-ttu-id="95026-158">Nome</span><span class="sxs-lookup"><span data-stu-id="95026-158">Name</span></span>| <span data-ttu-id="95026-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="95026-159">Type</span></span>| <span data-ttu-id="95026-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="95026-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="95026-161">String</span><span class="sxs-lookup"><span data-stu-id="95026-161">String</span></span>|<span data-ttu-id="95026-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="95026-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="95026-163">String</span><span class="sxs-lookup"><span data-stu-id="95026-163">String</span></span>|<span data-ttu-id="95026-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="95026-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95026-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="95026-165">Requirements</span></span>

|<span data-ttu-id="95026-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="95026-166">Requirement</span></span>| <span data-ttu-id="95026-167">Valor</span><span class="sxs-lookup"><span data-stu-id="95026-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="95026-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="95026-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95026-169">1.1</span><span class="sxs-lookup"><span data-stu-id="95026-169">1.1</span></span>|
|[<span data-ttu-id="95026-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="95026-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95026-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="95026-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="95026-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="95026-172">CoercionType: String</span></span>

<span data-ttu-id="95026-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="95026-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="95026-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="95026-174">Type</span></span>

*   <span data-ttu-id="95026-175">String</span><span class="sxs-lookup"><span data-stu-id="95026-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="95026-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="95026-176">Properties:</span></span>

|<span data-ttu-id="95026-177">Nome</span><span class="sxs-lookup"><span data-stu-id="95026-177">Name</span></span>| <span data-ttu-id="95026-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="95026-178">Type</span></span>| <span data-ttu-id="95026-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="95026-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="95026-180">String</span><span class="sxs-lookup"><span data-stu-id="95026-180">String</span></span>|<span data-ttu-id="95026-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="95026-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="95026-182">String</span><span class="sxs-lookup"><span data-stu-id="95026-182">String</span></span>|<span data-ttu-id="95026-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="95026-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95026-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="95026-184">Requirements</span></span>

|<span data-ttu-id="95026-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="95026-185">Requirement</span></span>| <span data-ttu-id="95026-186">Valor</span><span class="sxs-lookup"><span data-stu-id="95026-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="95026-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="95026-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95026-188">1.1</span><span class="sxs-lookup"><span data-stu-id="95026-188">1.1</span></span>|
|[<span data-ttu-id="95026-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="95026-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95026-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="95026-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="95026-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="95026-191">EventType: String</span></span>

<span data-ttu-id="95026-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="95026-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="95026-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="95026-193">Type</span></span>

*   <span data-ttu-id="95026-194">String</span><span class="sxs-lookup"><span data-stu-id="95026-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="95026-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="95026-195">Properties:</span></span>

| <span data-ttu-id="95026-196">Nome</span><span class="sxs-lookup"><span data-stu-id="95026-196">Name</span></span> | <span data-ttu-id="95026-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="95026-197">Type</span></span> | <span data-ttu-id="95026-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="95026-198">Description</span></span> | <span data-ttu-id="95026-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="95026-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="95026-200">String</span><span class="sxs-lookup"><span data-stu-id="95026-200">String</span></span> | <span data-ttu-id="95026-201">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="95026-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="95026-202">1.7</span><span class="sxs-lookup"><span data-stu-id="95026-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="95026-203">String</span><span class="sxs-lookup"><span data-stu-id="95026-203">String</span></span> | <span data-ttu-id="95026-204">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="95026-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="95026-205">1,8</span><span class="sxs-lookup"><span data-stu-id="95026-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="95026-206">String</span><span class="sxs-lookup"><span data-stu-id="95026-206">String</span></span> | <span data-ttu-id="95026-207">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="95026-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="95026-208">1,8</span><span class="sxs-lookup"><span data-stu-id="95026-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="95026-209">String</span><span class="sxs-lookup"><span data-stu-id="95026-209">String</span></span> | <span data-ttu-id="95026-210">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="95026-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="95026-211">1,5</span><span class="sxs-lookup"><span data-stu-id="95026-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="95026-212">String</span><span class="sxs-lookup"><span data-stu-id="95026-212">String</span></span> | <span data-ttu-id="95026-213">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="95026-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="95026-214">1.7</span><span class="sxs-lookup"><span data-stu-id="95026-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="95026-215">String</span><span class="sxs-lookup"><span data-stu-id="95026-215">String</span></span> | <span data-ttu-id="95026-216">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="95026-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="95026-217">1.7</span><span class="sxs-lookup"><span data-stu-id="95026-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="95026-218">Requisitos</span><span class="sxs-lookup"><span data-stu-id="95026-218">Requirements</span></span>

|<span data-ttu-id="95026-219">Requisito</span><span class="sxs-lookup"><span data-stu-id="95026-219">Requirement</span></span>| <span data-ttu-id="95026-220">Valor</span><span class="sxs-lookup"><span data-stu-id="95026-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="95026-221">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="95026-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95026-222">1,5</span><span class="sxs-lookup"><span data-stu-id="95026-222">1.5</span></span> |
|[<span data-ttu-id="95026-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="95026-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95026-224">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="95026-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="95026-225">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="95026-225">SourceProperty: String</span></span>

<span data-ttu-id="95026-226">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="95026-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="95026-227">Tipo</span><span class="sxs-lookup"><span data-stu-id="95026-227">Type</span></span>

*   <span data-ttu-id="95026-228">String</span><span class="sxs-lookup"><span data-stu-id="95026-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="95026-229">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="95026-229">Properties:</span></span>

|<span data-ttu-id="95026-230">Nome</span><span class="sxs-lookup"><span data-stu-id="95026-230">Name</span></span>| <span data-ttu-id="95026-231">Tipo</span><span class="sxs-lookup"><span data-stu-id="95026-231">Type</span></span>| <span data-ttu-id="95026-232">Descrição</span><span class="sxs-lookup"><span data-stu-id="95026-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="95026-233">String</span><span class="sxs-lookup"><span data-stu-id="95026-233">String</span></span>|<span data-ttu-id="95026-234">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="95026-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="95026-235">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="95026-235">String</span></span>|<span data-ttu-id="95026-236">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="95026-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="95026-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="95026-237">Requirements</span></span>

|<span data-ttu-id="95026-238">Requisito</span><span class="sxs-lookup"><span data-stu-id="95026-238">Requirement</span></span>| <span data-ttu-id="95026-239">Valor</span><span class="sxs-lookup"><span data-stu-id="95026-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="95026-240">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="95026-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95026-241">1.1</span><span class="sxs-lookup"><span data-stu-id="95026-241">1.1</span></span>|
|[<span data-ttu-id="95026-242">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="95026-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95026-243">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="95026-243">Compose or Read</span></span>|
