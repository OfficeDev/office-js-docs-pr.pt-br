---
title: Namespace do Office – conjunto de requisitos 1,9
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,9.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: e6a932c528dea692ff5fd7ea8d3e1454bb9a7e03
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628034"
---
# <a name="office-mailbox-requirement-set-19"></a><span data-ttu-id="7652b-103">Office (conjunto de requisitos de caixa de correio 1,9)</span><span class="sxs-lookup"><span data-stu-id="7652b-103">Office (Mailbox requirement set 1.9)</span></span>

<span data-ttu-id="7652b-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="7652b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7652b-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7652b-106">Requirements</span></span>

|<span data-ttu-id="7652b-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="7652b-107">Requirement</span></span>| <span data-ttu-id="7652b-108">Valor</span><span class="sxs-lookup"><span data-stu-id="7652b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="7652b-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7652b-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7652b-110">1.1</span><span class="sxs-lookup"><span data-stu-id="7652b-110">1.1</span></span>|
|[<span data-ttu-id="7652b-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7652b-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7652b-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7652b-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="7652b-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="7652b-113">Properties</span></span>

| <span data-ttu-id="7652b-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="7652b-114">Property</span></span> | <span data-ttu-id="7652b-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="7652b-115">Modes</span></span> | <span data-ttu-id="7652b-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="7652b-116">Return type</span></span> | <span data-ttu-id="7652b-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="7652b-117">Minimum</span></span><br><span data-ttu-id="7652b-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="7652b-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="7652b-119">context</span><span class="sxs-lookup"><span data-stu-id="7652b-119">context</span></span>](office.context.md) | <span data-ttu-id="7652b-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="7652b-120">Compose</span></span><br><span data-ttu-id="7652b-121">Leitura</span><span class="sxs-lookup"><span data-stu-id="7652b-121">Read</span></span> | [<span data-ttu-id="7652b-122">Context</span><span class="sxs-lookup"><span data-stu-id="7652b-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7652b-123">1.1</span><span class="sxs-lookup"><span data-stu-id="7652b-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="7652b-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="7652b-124">Enumerations</span></span>

| <span data-ttu-id="7652b-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="7652b-125">Enumeration</span></span> | <span data-ttu-id="7652b-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="7652b-126">Modes</span></span> | <span data-ttu-id="7652b-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="7652b-127">Return type</span></span> | <span data-ttu-id="7652b-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="7652b-128">Minimum</span></span><br><span data-ttu-id="7652b-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="7652b-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="7652b-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="7652b-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="7652b-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="7652b-131">Compose</span></span><br><span data-ttu-id="7652b-132">Leitura</span><span class="sxs-lookup"><span data-stu-id="7652b-132">Read</span></span> | <span data-ttu-id="7652b-133">String</span><span class="sxs-lookup"><span data-stu-id="7652b-133">String</span></span> | [<span data-ttu-id="7652b-134">1.1</span><span class="sxs-lookup"><span data-stu-id="7652b-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7652b-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="7652b-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="7652b-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="7652b-136">Compose</span></span><br><span data-ttu-id="7652b-137">Leitura</span><span class="sxs-lookup"><span data-stu-id="7652b-137">Read</span></span> | <span data-ttu-id="7652b-138">String</span><span class="sxs-lookup"><span data-stu-id="7652b-138">String</span></span> | [<span data-ttu-id="7652b-139">1.1</span><span class="sxs-lookup"><span data-stu-id="7652b-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7652b-140">EventType</span><span class="sxs-lookup"><span data-stu-id="7652b-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="7652b-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="7652b-141">Compose</span></span><br><span data-ttu-id="7652b-142">Leitura</span><span class="sxs-lookup"><span data-stu-id="7652b-142">Read</span></span> | <span data-ttu-id="7652b-143">String</span><span class="sxs-lookup"><span data-stu-id="7652b-143">String</span></span> | [<span data-ttu-id="7652b-144">1,5</span><span class="sxs-lookup"><span data-stu-id="7652b-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="7652b-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="7652b-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="7652b-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="7652b-146">Compose</span></span><br><span data-ttu-id="7652b-147">Leitura</span><span class="sxs-lookup"><span data-stu-id="7652b-147">Read</span></span> | <span data-ttu-id="7652b-148">String</span><span class="sxs-lookup"><span data-stu-id="7652b-148">String</span></span> | [<span data-ttu-id="7652b-149">1.1</span><span class="sxs-lookup"><span data-stu-id="7652b-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="7652b-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="7652b-150">Namespaces</span></span>

<span data-ttu-id="7652b-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true): inclui uma série de enumerações específicas do Outlook, por exemplo,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` e `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="7652b-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="7652b-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="7652b-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="7652b-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7652b-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="7652b-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="7652b-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="7652b-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="7652b-155">Type</span></span>

*   <span data-ttu-id="7652b-156">String</span><span class="sxs-lookup"><span data-stu-id="7652b-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7652b-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7652b-157">Properties:</span></span>

|<span data-ttu-id="7652b-158">Nome</span><span class="sxs-lookup"><span data-stu-id="7652b-158">Name</span></span>| <span data-ttu-id="7652b-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="7652b-159">Type</span></span>| <span data-ttu-id="7652b-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="7652b-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="7652b-161">String</span><span class="sxs-lookup"><span data-stu-id="7652b-161">String</span></span>|<span data-ttu-id="7652b-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="7652b-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="7652b-163">String</span><span class="sxs-lookup"><span data-stu-id="7652b-163">String</span></span>|<span data-ttu-id="7652b-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="7652b-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7652b-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7652b-165">Requirements</span></span>

|<span data-ttu-id="7652b-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="7652b-166">Requirement</span></span>| <span data-ttu-id="7652b-167">Valor</span><span class="sxs-lookup"><span data-stu-id="7652b-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="7652b-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7652b-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7652b-169">1.1</span><span class="sxs-lookup"><span data-stu-id="7652b-169">1.1</span></span>|
|[<span data-ttu-id="7652b-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7652b-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7652b-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7652b-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="7652b-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7652b-172">CoercionType: String</span></span>

<span data-ttu-id="7652b-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="7652b-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7652b-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="7652b-174">Type</span></span>

*   <span data-ttu-id="7652b-175">String</span><span class="sxs-lookup"><span data-stu-id="7652b-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7652b-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7652b-176">Properties:</span></span>

|<span data-ttu-id="7652b-177">Nome</span><span class="sxs-lookup"><span data-stu-id="7652b-177">Name</span></span>| <span data-ttu-id="7652b-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="7652b-178">Type</span></span>| <span data-ttu-id="7652b-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="7652b-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="7652b-180">String</span><span class="sxs-lookup"><span data-stu-id="7652b-180">String</span></span>|<span data-ttu-id="7652b-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="7652b-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="7652b-182">String</span><span class="sxs-lookup"><span data-stu-id="7652b-182">String</span></span>|<span data-ttu-id="7652b-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="7652b-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7652b-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7652b-184">Requirements</span></span>

|<span data-ttu-id="7652b-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="7652b-185">Requirement</span></span>| <span data-ttu-id="7652b-186">Valor</span><span class="sxs-lookup"><span data-stu-id="7652b-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="7652b-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7652b-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7652b-188">1.1</span><span class="sxs-lookup"><span data-stu-id="7652b-188">1.1</span></span>|
|[<span data-ttu-id="7652b-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7652b-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7652b-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7652b-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="7652b-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7652b-191">EventType: String</span></span>

<span data-ttu-id="7652b-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="7652b-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="7652b-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="7652b-193">Type</span></span>

*   <span data-ttu-id="7652b-194">String</span><span class="sxs-lookup"><span data-stu-id="7652b-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7652b-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7652b-195">Properties:</span></span>

| <span data-ttu-id="7652b-196">Nome</span><span class="sxs-lookup"><span data-stu-id="7652b-196">Name</span></span> | <span data-ttu-id="7652b-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="7652b-197">Type</span></span> | <span data-ttu-id="7652b-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="7652b-198">Description</span></span> | <span data-ttu-id="7652b-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="7652b-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="7652b-200">String</span><span class="sxs-lookup"><span data-stu-id="7652b-200">String</span></span> | <span data-ttu-id="7652b-201">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="7652b-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="7652b-202">1.7</span><span class="sxs-lookup"><span data-stu-id="7652b-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="7652b-203">String</span><span class="sxs-lookup"><span data-stu-id="7652b-203">String</span></span> | <span data-ttu-id="7652b-204">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="7652b-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="7652b-205">1,8</span><span class="sxs-lookup"><span data-stu-id="7652b-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="7652b-206">String</span><span class="sxs-lookup"><span data-stu-id="7652b-206">String</span></span> | <span data-ttu-id="7652b-207">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="7652b-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="7652b-208">1,8</span><span class="sxs-lookup"><span data-stu-id="7652b-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="7652b-209">String</span><span class="sxs-lookup"><span data-stu-id="7652b-209">String</span></span> | <span data-ttu-id="7652b-210">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="7652b-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="7652b-211">1,5</span><span class="sxs-lookup"><span data-stu-id="7652b-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="7652b-212">String</span><span class="sxs-lookup"><span data-stu-id="7652b-212">String</span></span> | <span data-ttu-id="7652b-213">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="7652b-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="7652b-214">1.7</span><span class="sxs-lookup"><span data-stu-id="7652b-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="7652b-215">String</span><span class="sxs-lookup"><span data-stu-id="7652b-215">String</span></span> | <span data-ttu-id="7652b-216">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="7652b-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="7652b-217">1.7</span><span class="sxs-lookup"><span data-stu-id="7652b-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7652b-218">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7652b-218">Requirements</span></span>

|<span data-ttu-id="7652b-219">Requisito</span><span class="sxs-lookup"><span data-stu-id="7652b-219">Requirement</span></span>| <span data-ttu-id="7652b-220">Valor</span><span class="sxs-lookup"><span data-stu-id="7652b-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="7652b-221">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7652b-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7652b-222">1,5</span><span class="sxs-lookup"><span data-stu-id="7652b-222">1.5</span></span> |
|[<span data-ttu-id="7652b-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7652b-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7652b-224">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7652b-224">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="7652b-225">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7652b-225">SourceProperty: String</span></span>

<span data-ttu-id="7652b-226">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="7652b-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7652b-227">Tipo</span><span class="sxs-lookup"><span data-stu-id="7652b-227">Type</span></span>

*   <span data-ttu-id="7652b-228">String</span><span class="sxs-lookup"><span data-stu-id="7652b-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7652b-229">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7652b-229">Properties:</span></span>

|<span data-ttu-id="7652b-230">Nome</span><span class="sxs-lookup"><span data-stu-id="7652b-230">Name</span></span>| <span data-ttu-id="7652b-231">Tipo</span><span class="sxs-lookup"><span data-stu-id="7652b-231">Type</span></span>| <span data-ttu-id="7652b-232">Descrição</span><span class="sxs-lookup"><span data-stu-id="7652b-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="7652b-233">String</span><span class="sxs-lookup"><span data-stu-id="7652b-233">String</span></span>|<span data-ttu-id="7652b-234">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7652b-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="7652b-235">String</span><span class="sxs-lookup"><span data-stu-id="7652b-235">String</span></span>|<span data-ttu-id="7652b-236">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7652b-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7652b-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7652b-237">Requirements</span></span>

|<span data-ttu-id="7652b-238">Requisito</span><span class="sxs-lookup"><span data-stu-id="7652b-238">Requirement</span></span>| <span data-ttu-id="7652b-239">Valor</span><span class="sxs-lookup"><span data-stu-id="7652b-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="7652b-240">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7652b-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7652b-241">1.1</span><span class="sxs-lookup"><span data-stu-id="7652b-241">1.1</span></span>|
|[<span data-ttu-id="7652b-242">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7652b-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7652b-243">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="7652b-243">Compose or Read</span></span>|
