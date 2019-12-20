---
title: Namespace do Office – conjunto de requisitos 1,8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: b23afd7b84dcd18e120f6aea4bd4fb0952791f1c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814163"
---
# <a name="office"></a><span data-ttu-id="f779b-102">Office</span><span class="sxs-lookup"><span data-stu-id="f779b-102">Office</span></span>

<span data-ttu-id="f779b-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f779b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f779b-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f779b-105">Requirements</span></span>

|<span data-ttu-id="f779b-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="f779b-106">Requirement</span></span>| <span data-ttu-id="f779b-107">Valor</span><span class="sxs-lookup"><span data-stu-id="f779b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f779b-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f779b-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f779b-109">1.1</span><span class="sxs-lookup"><span data-stu-id="f779b-109">1.1</span></span>|
|[<span data-ttu-id="f779b-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f779b-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f779b-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f779b-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f779b-112">Propriedades</span><span class="sxs-lookup"><span data-stu-id="f779b-112">Properties</span></span>

| <span data-ttu-id="f779b-113">Propriedade</span><span class="sxs-lookup"><span data-stu-id="f779b-113">Property</span></span> | <span data-ttu-id="f779b-114">Modelos</span><span class="sxs-lookup"><span data-stu-id="f779b-114">Modes</span></span> | <span data-ttu-id="f779b-115">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="f779b-115">Return type</span></span> | <span data-ttu-id="f779b-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="f779b-116">Minimum</span></span><br><span data-ttu-id="f779b-117">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="f779b-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f779b-118">context</span><span class="sxs-lookup"><span data-stu-id="f779b-118">context</span></span>](office.context.md) | <span data-ttu-id="f779b-119">Escrever</span><span class="sxs-lookup"><span data-stu-id="f779b-119">Compose</span></span><br><span data-ttu-id="f779b-120">Leitura</span><span class="sxs-lookup"><span data-stu-id="f779b-120">Read</span></span> | [<span data-ttu-id="f779b-121">Context</span><span class="sxs-lookup"><span data-stu-id="f779b-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="f779b-122">1.1</span><span class="sxs-lookup"><span data-stu-id="f779b-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="f779b-123">Enumerações</span><span class="sxs-lookup"><span data-stu-id="f779b-123">Enumerations</span></span>

| <span data-ttu-id="f779b-124">Enumeração</span><span class="sxs-lookup"><span data-stu-id="f779b-124">Enumeration</span></span> | <span data-ttu-id="f779b-125">Modelos</span><span class="sxs-lookup"><span data-stu-id="f779b-125">Modes</span></span> | <span data-ttu-id="f779b-126">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="f779b-126">Return type</span></span> | <span data-ttu-id="f779b-127">Mínimo</span><span class="sxs-lookup"><span data-stu-id="f779b-127">Minimum</span></span><br><span data-ttu-id="f779b-128">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="f779b-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f779b-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f779b-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f779b-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="f779b-130">Compose</span></span><br><span data-ttu-id="f779b-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="f779b-131">Read</span></span> | <span data-ttu-id="f779b-132">String</span><span class="sxs-lookup"><span data-stu-id="f779b-132">String</span></span> | [<span data-ttu-id="f779b-133">1.1</span><span class="sxs-lookup"><span data-stu-id="f779b-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f779b-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f779b-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f779b-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="f779b-135">Compose</span></span><br><span data-ttu-id="f779b-136">Leitura</span><span class="sxs-lookup"><span data-stu-id="f779b-136">Read</span></span> | <span data-ttu-id="f779b-137">String</span><span class="sxs-lookup"><span data-stu-id="f779b-137">String</span></span> | [<span data-ttu-id="f779b-138">1.1</span><span class="sxs-lookup"><span data-stu-id="f779b-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f779b-139">EventType</span><span class="sxs-lookup"><span data-stu-id="f779b-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f779b-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="f779b-140">Compose</span></span><br><span data-ttu-id="f779b-141">Leitura</span><span class="sxs-lookup"><span data-stu-id="f779b-141">Read</span></span> | <span data-ttu-id="f779b-142">String</span><span class="sxs-lookup"><span data-stu-id="f779b-142">String</span></span> | [<span data-ttu-id="f779b-143">1,5</span><span class="sxs-lookup"><span data-stu-id="f779b-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="f779b-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f779b-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f779b-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="f779b-145">Compose</span></span><br><span data-ttu-id="f779b-146">Leitura</span><span class="sxs-lookup"><span data-stu-id="f779b-146">Read</span></span> | <span data-ttu-id="f779b-147">String</span><span class="sxs-lookup"><span data-stu-id="f779b-147">String</span></span> | [<span data-ttu-id="f779b-148">1.1</span><span class="sxs-lookup"><span data-stu-id="f779b-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="f779b-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="f779b-149">Namespaces</span></span>

<span data-ttu-id="f779b-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="f779b-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="f779b-151">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="f779b-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f779b-152">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f779b-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="f779b-153">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="f779b-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f779b-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="f779b-154">Type</span></span>

*   <span data-ttu-id="f779b-155">String</span><span class="sxs-lookup"><span data-stu-id="f779b-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f779b-156">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f779b-156">Properties:</span></span>

|<span data-ttu-id="f779b-157">Nome</span><span class="sxs-lookup"><span data-stu-id="f779b-157">Name</span></span>| <span data-ttu-id="f779b-158">Tipo</span><span class="sxs-lookup"><span data-stu-id="f779b-158">Type</span></span>| <span data-ttu-id="f779b-159">Descrição</span><span class="sxs-lookup"><span data-stu-id="f779b-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f779b-160">String</span><span class="sxs-lookup"><span data-stu-id="f779b-160">String</span></span>|<span data-ttu-id="f779b-161">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="f779b-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f779b-162">String</span><span class="sxs-lookup"><span data-stu-id="f779b-162">String</span></span>|<span data-ttu-id="f779b-163">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="f779b-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f779b-164">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f779b-164">Requirements</span></span>

|<span data-ttu-id="f779b-165">Requisito</span><span class="sxs-lookup"><span data-stu-id="f779b-165">Requirement</span></span>| <span data-ttu-id="f779b-166">Valor</span><span class="sxs-lookup"><span data-stu-id="f779b-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="f779b-167">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f779b-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f779b-168">1.1</span><span class="sxs-lookup"><span data-stu-id="f779b-168">1.1</span></span>|
|[<span data-ttu-id="f779b-169">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f779b-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f779b-170">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f779b-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f779b-171">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f779b-171">CoercionType: String</span></span>

<span data-ttu-id="f779b-172">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="f779b-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f779b-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="f779b-173">Type</span></span>

*   <span data-ttu-id="f779b-174">String</span><span class="sxs-lookup"><span data-stu-id="f779b-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f779b-175">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f779b-175">Properties:</span></span>

|<span data-ttu-id="f779b-176">Nome</span><span class="sxs-lookup"><span data-stu-id="f779b-176">Name</span></span>| <span data-ttu-id="f779b-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="f779b-177">Type</span></span>| <span data-ttu-id="f779b-178">Descrição</span><span class="sxs-lookup"><span data-stu-id="f779b-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f779b-179">String</span><span class="sxs-lookup"><span data-stu-id="f779b-179">String</span></span>|<span data-ttu-id="f779b-180">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="f779b-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f779b-181">String</span><span class="sxs-lookup"><span data-stu-id="f779b-181">String</span></span>|<span data-ttu-id="f779b-182">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="f779b-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f779b-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f779b-183">Requirements</span></span>

|<span data-ttu-id="f779b-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="f779b-184">Requirement</span></span>| <span data-ttu-id="f779b-185">Valor</span><span class="sxs-lookup"><span data-stu-id="f779b-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="f779b-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f779b-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f779b-187">1.1</span><span class="sxs-lookup"><span data-stu-id="f779b-187">1.1</span></span>|
|[<span data-ttu-id="f779b-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f779b-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f779b-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f779b-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="f779b-190">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f779b-190">EventType: String</span></span>

<span data-ttu-id="f779b-191">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="f779b-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f779b-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="f779b-192">Type</span></span>

*   <span data-ttu-id="f779b-193">String</span><span class="sxs-lookup"><span data-stu-id="f779b-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f779b-194">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f779b-194">Properties:</span></span>

| <span data-ttu-id="f779b-195">Nome</span><span class="sxs-lookup"><span data-stu-id="f779b-195">Name</span></span> | <span data-ttu-id="f779b-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="f779b-196">Type</span></span> | <span data-ttu-id="f779b-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="f779b-197">Description</span></span> | <span data-ttu-id="f779b-198">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="f779b-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="f779b-199">String</span><span class="sxs-lookup"><span data-stu-id="f779b-199">String</span></span> | <span data-ttu-id="f779b-200">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="f779b-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="f779b-201">1.7</span><span class="sxs-lookup"><span data-stu-id="f779b-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="f779b-202">String</span><span class="sxs-lookup"><span data-stu-id="f779b-202">String</span></span> | <span data-ttu-id="f779b-203">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="f779b-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="f779b-204">1,8</span><span class="sxs-lookup"><span data-stu-id="f779b-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="f779b-205">String</span><span class="sxs-lookup"><span data-stu-id="f779b-205">String</span></span> | <span data-ttu-id="f779b-206">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="f779b-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="f779b-207">1,8</span><span class="sxs-lookup"><span data-stu-id="f779b-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="f779b-208">String</span><span class="sxs-lookup"><span data-stu-id="f779b-208">String</span></span> | <span data-ttu-id="f779b-209">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="f779b-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="f779b-210">1,5</span><span class="sxs-lookup"><span data-stu-id="f779b-210">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="f779b-211">String</span><span class="sxs-lookup"><span data-stu-id="f779b-211">String</span></span> | <span data-ttu-id="f779b-212">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="f779b-212">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="f779b-213">1.7</span><span class="sxs-lookup"><span data-stu-id="f779b-213">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="f779b-214">String</span><span class="sxs-lookup"><span data-stu-id="f779b-214">String</span></span> | <span data-ttu-id="f779b-215">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="f779b-215">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="f779b-216">1.7</span><span class="sxs-lookup"><span data-stu-id="f779b-216">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f779b-217">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f779b-217">Requirements</span></span>

|<span data-ttu-id="f779b-218">Requisito</span><span class="sxs-lookup"><span data-stu-id="f779b-218">Requirement</span></span>| <span data-ttu-id="f779b-219">Valor</span><span class="sxs-lookup"><span data-stu-id="f779b-219">Value</span></span>|
|---|---|
|[<span data-ttu-id="f779b-220">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f779b-220">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f779b-221">1,5</span><span class="sxs-lookup"><span data-stu-id="f779b-221">1.5</span></span> |
|[<span data-ttu-id="f779b-222">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f779b-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f779b-223">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f779b-223">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f779b-224">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f779b-224">SourceProperty: String</span></span>

<span data-ttu-id="f779b-225">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="f779b-225">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f779b-226">Tipo</span><span class="sxs-lookup"><span data-stu-id="f779b-226">Type</span></span>

*   <span data-ttu-id="f779b-227">String</span><span class="sxs-lookup"><span data-stu-id="f779b-227">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f779b-228">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f779b-228">Properties:</span></span>

|<span data-ttu-id="f779b-229">Nome</span><span class="sxs-lookup"><span data-stu-id="f779b-229">Name</span></span>| <span data-ttu-id="f779b-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="f779b-230">Type</span></span>| <span data-ttu-id="f779b-231">Descrição</span><span class="sxs-lookup"><span data-stu-id="f779b-231">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f779b-232">String</span><span class="sxs-lookup"><span data-stu-id="f779b-232">String</span></span>|<span data-ttu-id="f779b-233">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f779b-233">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f779b-234">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f779b-234">String</span></span>|<span data-ttu-id="f779b-235">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f779b-235">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f779b-236">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f779b-236">Requirements</span></span>

|<span data-ttu-id="f779b-237">Requisito</span><span class="sxs-lookup"><span data-stu-id="f779b-237">Requirement</span></span>| <span data-ttu-id="f779b-238">Valor</span><span class="sxs-lookup"><span data-stu-id="f779b-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="f779b-239">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f779b-239">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f779b-240">1.1</span><span class="sxs-lookup"><span data-stu-id="f779b-240">1.1</span></span>|
|[<span data-ttu-id="f779b-241">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f779b-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f779b-242">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f779b-242">Compose or Read</span></span>|
