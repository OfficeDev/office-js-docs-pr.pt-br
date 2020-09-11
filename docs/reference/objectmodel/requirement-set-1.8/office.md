---
title: Namespace do Office – conjunto de requisitos 1,8
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,8.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: e0580cd1bb327c8673c46d3d0292aec9f2f1c971
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431518"
---
# <a name="office-mailbox-requirement-set-18"></a><span data-ttu-id="eacae-103">Office (conjunto de requisitos de caixa de correio 1,8)</span><span class="sxs-lookup"><span data-stu-id="eacae-103">Office (Mailbox requirement set 1.8)</span></span>

<span data-ttu-id="eacae-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="eacae-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="eacae-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="eacae-106">Requirements</span></span>

|<span data-ttu-id="eacae-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="eacae-107">Requirement</span></span>| <span data-ttu-id="eacae-108">Valor</span><span class="sxs-lookup"><span data-stu-id="eacae-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="eacae-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="eacae-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eacae-110">1.1</span><span class="sxs-lookup"><span data-stu-id="eacae-110">1.1</span></span>|
|[<span data-ttu-id="eacae-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="eacae-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eacae-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="eacae-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="eacae-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="eacae-113">Properties</span></span>

| <span data-ttu-id="eacae-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="eacae-114">Property</span></span> | <span data-ttu-id="eacae-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="eacae-115">Modes</span></span> | <span data-ttu-id="eacae-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="eacae-116">Return type</span></span> | <span data-ttu-id="eacae-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="eacae-117">Minimum</span></span><br><span data-ttu-id="eacae-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="eacae-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="eacae-119">context</span><span class="sxs-lookup"><span data-stu-id="eacae-119">context</span></span>](office.context.md) | <span data-ttu-id="eacae-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="eacae-120">Compose</span></span><br><span data-ttu-id="eacae-121">Ler</span><span class="sxs-lookup"><span data-stu-id="eacae-121">Read</span></span> | [<span data-ttu-id="eacae-122">Context</span><span class="sxs-lookup"><span data-stu-id="eacae-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="eacae-123">1.1</span><span class="sxs-lookup"><span data-stu-id="eacae-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="eacae-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="eacae-124">Enumerations</span></span>

| <span data-ttu-id="eacae-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="eacae-125">Enumeration</span></span> | <span data-ttu-id="eacae-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="eacae-126">Modes</span></span> | <span data-ttu-id="eacae-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="eacae-127">Return type</span></span> | <span data-ttu-id="eacae-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="eacae-128">Minimum</span></span><br><span data-ttu-id="eacae-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="eacae-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="eacae-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="eacae-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="eacae-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="eacae-131">Compose</span></span><br><span data-ttu-id="eacae-132">Ler</span><span class="sxs-lookup"><span data-stu-id="eacae-132">Read</span></span> | <span data-ttu-id="eacae-133">String</span><span class="sxs-lookup"><span data-stu-id="eacae-133">String</span></span> | [<span data-ttu-id="eacae-134">1.1</span><span class="sxs-lookup"><span data-stu-id="eacae-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eacae-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="eacae-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="eacae-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="eacae-136">Compose</span></span><br><span data-ttu-id="eacae-137">Ler</span><span class="sxs-lookup"><span data-stu-id="eacae-137">Read</span></span> | <span data-ttu-id="eacae-138">String</span><span class="sxs-lookup"><span data-stu-id="eacae-138">String</span></span> | [<span data-ttu-id="eacae-139">1.1</span><span class="sxs-lookup"><span data-stu-id="eacae-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eacae-140">EventType</span><span class="sxs-lookup"><span data-stu-id="eacae-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="eacae-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="eacae-141">Compose</span></span><br><span data-ttu-id="eacae-142">Ler</span><span class="sxs-lookup"><span data-stu-id="eacae-142">Read</span></span> | <span data-ttu-id="eacae-143">String</span><span class="sxs-lookup"><span data-stu-id="eacae-143">String</span></span> | [<span data-ttu-id="eacae-144">1,5</span><span class="sxs-lookup"><span data-stu-id="eacae-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="eacae-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="eacae-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="eacae-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="eacae-146">Compose</span></span><br><span data-ttu-id="eacae-147">Ler</span><span class="sxs-lookup"><span data-stu-id="eacae-147">Read</span></span> | <span data-ttu-id="eacae-148">String</span><span class="sxs-lookup"><span data-stu-id="eacae-148">String</span></span> | [<span data-ttu-id="eacae-149">1.1</span><span class="sxs-lookup"><span data-stu-id="eacae-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="eacae-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="eacae-150">Namespaces</span></span>

<span data-ttu-id="eacae-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true): inclui uma série de enumerações específicas do Outlook, por exemplo,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` e `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="eacae-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="eacae-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="eacae-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="eacae-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="eacae-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="eacae-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="eacae-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="eacae-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="eacae-155">Type</span></span>

*   <span data-ttu-id="eacae-156">String</span><span class="sxs-lookup"><span data-stu-id="eacae-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eacae-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="eacae-157">Properties:</span></span>

|<span data-ttu-id="eacae-158">Nome</span><span class="sxs-lookup"><span data-stu-id="eacae-158">Name</span></span>| <span data-ttu-id="eacae-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="eacae-159">Type</span></span>| <span data-ttu-id="eacae-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="eacae-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="eacae-161">String</span><span class="sxs-lookup"><span data-stu-id="eacae-161">String</span></span>|<span data-ttu-id="eacae-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="eacae-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="eacae-163">String</span><span class="sxs-lookup"><span data-stu-id="eacae-163">String</span></span>|<span data-ttu-id="eacae-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="eacae-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eacae-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="eacae-165">Requirements</span></span>

|<span data-ttu-id="eacae-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="eacae-166">Requirement</span></span>| <span data-ttu-id="eacae-167">Valor</span><span class="sxs-lookup"><span data-stu-id="eacae-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="eacae-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="eacae-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eacae-169">1.1</span><span class="sxs-lookup"><span data-stu-id="eacae-169">1.1</span></span>|
|[<span data-ttu-id="eacae-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="eacae-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eacae-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="eacae-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="eacae-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="eacae-172">CoercionType: String</span></span>

<span data-ttu-id="eacae-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="eacae-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="eacae-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="eacae-174">Type</span></span>

*   <span data-ttu-id="eacae-175">String</span><span class="sxs-lookup"><span data-stu-id="eacae-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eacae-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="eacae-176">Properties:</span></span>

|<span data-ttu-id="eacae-177">Nome</span><span class="sxs-lookup"><span data-stu-id="eacae-177">Name</span></span>| <span data-ttu-id="eacae-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="eacae-178">Type</span></span>| <span data-ttu-id="eacae-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="eacae-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="eacae-180">String</span><span class="sxs-lookup"><span data-stu-id="eacae-180">String</span></span>|<span data-ttu-id="eacae-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="eacae-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="eacae-182">String</span><span class="sxs-lookup"><span data-stu-id="eacae-182">String</span></span>|<span data-ttu-id="eacae-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="eacae-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eacae-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="eacae-184">Requirements</span></span>

|<span data-ttu-id="eacae-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="eacae-185">Requirement</span></span>| <span data-ttu-id="eacae-186">Valor</span><span class="sxs-lookup"><span data-stu-id="eacae-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="eacae-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="eacae-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eacae-188">1.1</span><span class="sxs-lookup"><span data-stu-id="eacae-188">1.1</span></span>|
|[<span data-ttu-id="eacae-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="eacae-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eacae-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="eacae-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="eacae-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="eacae-191">EventType: String</span></span>

<span data-ttu-id="eacae-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="eacae-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="eacae-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="eacae-193">Type</span></span>

*   <span data-ttu-id="eacae-194">String</span><span class="sxs-lookup"><span data-stu-id="eacae-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eacae-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="eacae-195">Properties:</span></span>

| <span data-ttu-id="eacae-196">Nome</span><span class="sxs-lookup"><span data-stu-id="eacae-196">Name</span></span> | <span data-ttu-id="eacae-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="eacae-197">Type</span></span> | <span data-ttu-id="eacae-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="eacae-198">Description</span></span> | <span data-ttu-id="eacae-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="eacae-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="eacae-200">String</span><span class="sxs-lookup"><span data-stu-id="eacae-200">String</span></span> | <span data-ttu-id="eacae-201">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="eacae-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="eacae-202">1.7</span><span class="sxs-lookup"><span data-stu-id="eacae-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="eacae-203">String</span><span class="sxs-lookup"><span data-stu-id="eacae-203">String</span></span> | <span data-ttu-id="eacae-204">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="eacae-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="eacae-205">1,8</span><span class="sxs-lookup"><span data-stu-id="eacae-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="eacae-206">String</span><span class="sxs-lookup"><span data-stu-id="eacae-206">String</span></span> | <span data-ttu-id="eacae-207">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="eacae-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="eacae-208">1,8</span><span class="sxs-lookup"><span data-stu-id="eacae-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="eacae-209">String</span><span class="sxs-lookup"><span data-stu-id="eacae-209">String</span></span> | <span data-ttu-id="eacae-210">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="eacae-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="eacae-211">1,5</span><span class="sxs-lookup"><span data-stu-id="eacae-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="eacae-212">String</span><span class="sxs-lookup"><span data-stu-id="eacae-212">String</span></span> | <span data-ttu-id="eacae-213">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="eacae-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="eacae-214">1.7</span><span class="sxs-lookup"><span data-stu-id="eacae-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="eacae-215">String</span><span class="sxs-lookup"><span data-stu-id="eacae-215">String</span></span> | <span data-ttu-id="eacae-216">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="eacae-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="eacae-217">1.7</span><span class="sxs-lookup"><span data-stu-id="eacae-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="eacae-218">Requisitos</span><span class="sxs-lookup"><span data-stu-id="eacae-218">Requirements</span></span>

|<span data-ttu-id="eacae-219">Requisito</span><span class="sxs-lookup"><span data-stu-id="eacae-219">Requirement</span></span>| <span data-ttu-id="eacae-220">Valor</span><span class="sxs-lookup"><span data-stu-id="eacae-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="eacae-221">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="eacae-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eacae-222">1,5</span><span class="sxs-lookup"><span data-stu-id="eacae-222">1.5</span></span> |
|[<span data-ttu-id="eacae-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="eacae-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eacae-224">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="eacae-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="eacae-225">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="eacae-225">SourceProperty: String</span></span>

<span data-ttu-id="eacae-226">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="eacae-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="eacae-227">Tipo</span><span class="sxs-lookup"><span data-stu-id="eacae-227">Type</span></span>

*   <span data-ttu-id="eacae-228">String</span><span class="sxs-lookup"><span data-stu-id="eacae-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eacae-229">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="eacae-229">Properties:</span></span>

|<span data-ttu-id="eacae-230">Nome</span><span class="sxs-lookup"><span data-stu-id="eacae-230">Name</span></span>| <span data-ttu-id="eacae-231">Tipo</span><span class="sxs-lookup"><span data-stu-id="eacae-231">Type</span></span>| <span data-ttu-id="eacae-232">Descrição</span><span class="sxs-lookup"><span data-stu-id="eacae-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="eacae-233">String</span><span class="sxs-lookup"><span data-stu-id="eacae-233">String</span></span>|<span data-ttu-id="eacae-234">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="eacae-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="eacae-235">String</span><span class="sxs-lookup"><span data-stu-id="eacae-235">String</span></span>|<span data-ttu-id="eacae-236">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="eacae-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eacae-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="eacae-237">Requirements</span></span>

|<span data-ttu-id="eacae-238">Requisito</span><span class="sxs-lookup"><span data-stu-id="eacae-238">Requirement</span></span>| <span data-ttu-id="eacae-239">Valor</span><span class="sxs-lookup"><span data-stu-id="eacae-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="eacae-240">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="eacae-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eacae-241">1.1</span><span class="sxs-lookup"><span data-stu-id="eacae-241">1.1</span></span>|
|[<span data-ttu-id="eacae-242">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="eacae-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eacae-243">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="eacae-243">Compose or Read</span></span>|
