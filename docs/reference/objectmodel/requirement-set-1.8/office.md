---
title: Namespace do Office – conjunto de requisitos 1,8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: c5c431f7a958f1c2a956f36e90ad0f3a205c6669
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163623"
---
# <a name="office"></a><span data-ttu-id="cd6ba-102">Office</span><span class="sxs-lookup"><span data-stu-id="cd6ba-102">Office</span></span>

<span data-ttu-id="cd6ba-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="cd6ba-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="cd6ba-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cd6ba-105">Requirements</span></span>

|<span data-ttu-id="cd6ba-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="cd6ba-106">Requirement</span></span>| <span data-ttu-id="cd6ba-107">Valor</span><span class="sxs-lookup"><span data-stu-id="cd6ba-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="cd6ba-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cd6ba-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cd6ba-109">1.1</span><span class="sxs-lookup"><span data-stu-id="cd6ba-109">1.1</span></span>|
|[<span data-ttu-id="cd6ba-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cd6ba-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cd6ba-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cd6ba-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="cd6ba-112">Propriedades</span><span class="sxs-lookup"><span data-stu-id="cd6ba-112">Properties</span></span>

| <span data-ttu-id="cd6ba-113">Propriedade</span><span class="sxs-lookup"><span data-stu-id="cd6ba-113">Property</span></span> | <span data-ttu-id="cd6ba-114">Modelos</span><span class="sxs-lookup"><span data-stu-id="cd6ba-114">Modes</span></span> | <span data-ttu-id="cd6ba-115">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="cd6ba-115">Return type</span></span> | <span data-ttu-id="cd6ba-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="cd6ba-116">Minimum</span></span><br><span data-ttu-id="cd6ba-117">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="cd6ba-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cd6ba-118">context</span><span class="sxs-lookup"><span data-stu-id="cd6ba-118">context</span></span>](office.context.md) | <span data-ttu-id="cd6ba-119">Escrever</span><span class="sxs-lookup"><span data-stu-id="cd6ba-119">Compose</span></span><br><span data-ttu-id="cd6ba-120">Ler</span><span class="sxs-lookup"><span data-stu-id="cd6ba-120">Read</span></span> | [<span data-ttu-id="cd6ba-121">Context</span><span class="sxs-lookup"><span data-stu-id="cd6ba-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="cd6ba-122">1.1</span><span class="sxs-lookup"><span data-stu-id="cd6ba-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="cd6ba-123">Enumerações</span><span class="sxs-lookup"><span data-stu-id="cd6ba-123">Enumerations</span></span>

| <span data-ttu-id="cd6ba-124">Enumeração</span><span class="sxs-lookup"><span data-stu-id="cd6ba-124">Enumeration</span></span> | <span data-ttu-id="cd6ba-125">Modelos</span><span class="sxs-lookup"><span data-stu-id="cd6ba-125">Modes</span></span> | <span data-ttu-id="cd6ba-126">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="cd6ba-126">Return type</span></span> | <span data-ttu-id="cd6ba-127">Mínimo</span><span class="sxs-lookup"><span data-stu-id="cd6ba-127">Minimum</span></span><br><span data-ttu-id="cd6ba-128">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="cd6ba-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cd6ba-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="cd6ba-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="cd6ba-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="cd6ba-130">Compose</span></span><br><span data-ttu-id="cd6ba-131">Ler</span><span class="sxs-lookup"><span data-stu-id="cd6ba-131">Read</span></span> | <span data-ttu-id="cd6ba-132">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-132">String</span></span> | [<span data-ttu-id="cd6ba-133">1.1</span><span class="sxs-lookup"><span data-stu-id="cd6ba-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cd6ba-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="cd6ba-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="cd6ba-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="cd6ba-135">Compose</span></span><br><span data-ttu-id="cd6ba-136">Ler</span><span class="sxs-lookup"><span data-stu-id="cd6ba-136">Read</span></span> | <span data-ttu-id="cd6ba-137">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-137">String</span></span> | [<span data-ttu-id="cd6ba-138">1.1</span><span class="sxs-lookup"><span data-stu-id="cd6ba-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cd6ba-139">EventType</span><span class="sxs-lookup"><span data-stu-id="cd6ba-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="cd6ba-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="cd6ba-140">Compose</span></span><br><span data-ttu-id="cd6ba-141">Ler</span><span class="sxs-lookup"><span data-stu-id="cd6ba-141">Read</span></span> | <span data-ttu-id="cd6ba-142">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-142">String</span></span> | [<span data-ttu-id="cd6ba-143">1,5</span><span class="sxs-lookup"><span data-stu-id="cd6ba-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="cd6ba-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="cd6ba-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="cd6ba-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="cd6ba-145">Compose</span></span><br><span data-ttu-id="cd6ba-146">Ler</span><span class="sxs-lookup"><span data-stu-id="cd6ba-146">Read</span></span> | <span data-ttu-id="cd6ba-147">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-147">String</span></span> | [<span data-ttu-id="cd6ba-148">1.1</span><span class="sxs-lookup"><span data-stu-id="cd6ba-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="cd6ba-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="cd6ba-149">Namespaces</span></span>

<span data-ttu-id="cd6ba-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="cd6ba-151">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="cd6ba-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="cd6ba-152">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="cd6ba-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="cd6ba-153">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="cd6ba-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="cd6ba-154">Type</span></span>

*   <span data-ttu-id="cd6ba-155">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cd6ba-156">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="cd6ba-156">Properties:</span></span>

|<span data-ttu-id="cd6ba-157">Nome</span><span class="sxs-lookup"><span data-stu-id="cd6ba-157">Name</span></span>| <span data-ttu-id="cd6ba-158">Tipo</span><span class="sxs-lookup"><span data-stu-id="cd6ba-158">Type</span></span>| <span data-ttu-id="cd6ba-159">Descrição</span><span class="sxs-lookup"><span data-stu-id="cd6ba-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="cd6ba-160">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-160">String</span></span>|<span data-ttu-id="cd6ba-161">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="cd6ba-162">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-162">String</span></span>|<span data-ttu-id="cd6ba-163">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cd6ba-164">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cd6ba-164">Requirements</span></span>

|<span data-ttu-id="cd6ba-165">Requisito</span><span class="sxs-lookup"><span data-stu-id="cd6ba-165">Requirement</span></span>| <span data-ttu-id="cd6ba-166">Valor</span><span class="sxs-lookup"><span data-stu-id="cd6ba-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="cd6ba-167">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cd6ba-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cd6ba-168">1.1</span><span class="sxs-lookup"><span data-stu-id="cd6ba-168">1.1</span></span>|
|[<span data-ttu-id="cd6ba-169">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cd6ba-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cd6ba-170">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cd6ba-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="cd6ba-171">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="cd6ba-171">CoercionType: String</span></span>

<span data-ttu-id="cd6ba-172">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cd6ba-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="cd6ba-173">Type</span></span>

*   <span data-ttu-id="cd6ba-174">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cd6ba-175">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="cd6ba-175">Properties:</span></span>

|<span data-ttu-id="cd6ba-176">Nome</span><span class="sxs-lookup"><span data-stu-id="cd6ba-176">Name</span></span>| <span data-ttu-id="cd6ba-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="cd6ba-177">Type</span></span>| <span data-ttu-id="cd6ba-178">Descrição</span><span class="sxs-lookup"><span data-stu-id="cd6ba-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="cd6ba-179">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-179">String</span></span>|<span data-ttu-id="cd6ba-180">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="cd6ba-181">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-181">String</span></span>|<span data-ttu-id="cd6ba-182">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cd6ba-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cd6ba-183">Requirements</span></span>

|<span data-ttu-id="cd6ba-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="cd6ba-184">Requirement</span></span>| <span data-ttu-id="cd6ba-185">Valor</span><span class="sxs-lookup"><span data-stu-id="cd6ba-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="cd6ba-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cd6ba-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cd6ba-187">1.1</span><span class="sxs-lookup"><span data-stu-id="cd6ba-187">1.1</span></span>|
|[<span data-ttu-id="cd6ba-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cd6ba-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cd6ba-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cd6ba-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="cd6ba-190">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="cd6ba-190">EventType: String</span></span>

<span data-ttu-id="cd6ba-191">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="cd6ba-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="cd6ba-192">Type</span></span>

*   <span data-ttu-id="cd6ba-193">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cd6ba-194">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="cd6ba-194">Properties:</span></span>

| <span data-ttu-id="cd6ba-195">Nome</span><span class="sxs-lookup"><span data-stu-id="cd6ba-195">Name</span></span> | <span data-ttu-id="cd6ba-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="cd6ba-196">Type</span></span> | <span data-ttu-id="cd6ba-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="cd6ba-197">Description</span></span> | <span data-ttu-id="cd6ba-198">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="cd6ba-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="cd6ba-199">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-199">String</span></span> | <span data-ttu-id="cd6ba-200">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="cd6ba-201">1.7</span><span class="sxs-lookup"><span data-stu-id="cd6ba-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="cd6ba-202">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-202">String</span></span> | <span data-ttu-id="cd6ba-203">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="cd6ba-204">1,8</span><span class="sxs-lookup"><span data-stu-id="cd6ba-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="cd6ba-205">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-205">String</span></span> | <span data-ttu-id="cd6ba-206">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="cd6ba-207">1,8</span><span class="sxs-lookup"><span data-stu-id="cd6ba-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="cd6ba-208">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-208">String</span></span> | <span data-ttu-id="cd6ba-209">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="cd6ba-210">1,5</span><span class="sxs-lookup"><span data-stu-id="cd6ba-210">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="cd6ba-211">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-211">String</span></span> | <span data-ttu-id="cd6ba-212">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-212">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="cd6ba-213">1.7</span><span class="sxs-lookup"><span data-stu-id="cd6ba-213">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="cd6ba-214">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-214">String</span></span> | <span data-ttu-id="cd6ba-215">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-215">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="cd6ba-216">1.7</span><span class="sxs-lookup"><span data-stu-id="cd6ba-216">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cd6ba-217">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cd6ba-217">Requirements</span></span>

|<span data-ttu-id="cd6ba-218">Requisito</span><span class="sxs-lookup"><span data-stu-id="cd6ba-218">Requirement</span></span>| <span data-ttu-id="cd6ba-219">Valor</span><span class="sxs-lookup"><span data-stu-id="cd6ba-219">Value</span></span>|
|---|---|
|[<span data-ttu-id="cd6ba-220">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cd6ba-220">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cd6ba-221">1,5</span><span class="sxs-lookup"><span data-stu-id="cd6ba-221">1.5</span></span> |
|[<span data-ttu-id="cd6ba-222">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cd6ba-222">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cd6ba-223">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cd6ba-223">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="cd6ba-224">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="cd6ba-224">SourceProperty: String</span></span>

<span data-ttu-id="cd6ba-225">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-225">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cd6ba-226">Tipo</span><span class="sxs-lookup"><span data-stu-id="cd6ba-226">Type</span></span>

*   <span data-ttu-id="cd6ba-227">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-227">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cd6ba-228">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="cd6ba-228">Properties:</span></span>

|<span data-ttu-id="cd6ba-229">Nome</span><span class="sxs-lookup"><span data-stu-id="cd6ba-229">Name</span></span>| <span data-ttu-id="cd6ba-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="cd6ba-230">Type</span></span>| <span data-ttu-id="cd6ba-231">Descrição</span><span class="sxs-lookup"><span data-stu-id="cd6ba-231">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="cd6ba-232">String</span><span class="sxs-lookup"><span data-stu-id="cd6ba-232">String</span></span>|<span data-ttu-id="cd6ba-233">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-233">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="cd6ba-234">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="cd6ba-234">String</span></span>|<span data-ttu-id="cd6ba-235">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="cd6ba-235">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cd6ba-236">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cd6ba-236">Requirements</span></span>

|<span data-ttu-id="cd6ba-237">Requisito</span><span class="sxs-lookup"><span data-stu-id="cd6ba-237">Requirement</span></span>| <span data-ttu-id="cd6ba-238">Valor</span><span class="sxs-lookup"><span data-stu-id="cd6ba-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="cd6ba-239">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="cd6ba-239">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cd6ba-240">1.1</span><span class="sxs-lookup"><span data-stu-id="cd6ba-240">1.1</span></span>|
|[<span data-ttu-id="cd6ba-241">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="cd6ba-241">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cd6ba-242">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="cd6ba-242">Compose or Read</span></span>|
