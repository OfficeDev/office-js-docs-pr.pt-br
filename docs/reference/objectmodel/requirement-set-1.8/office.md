---
title: Office namespace - conjunto de requisitos 1.8
description: Office namespace disponíveis para os Outlook que usam o conjunto de requisitos da API de Caixa de Correio 1.8.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 00e236bed7e00159be8c94f727ca64ccaecd07b0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590523"
---
# <a name="office-mailbox-requirement-set-18"></a><span data-ttu-id="11dd9-103">Office (conjunto de requisitos de caixa de correio 1.8)</span><span class="sxs-lookup"><span data-stu-id="11dd9-103">Office (Mailbox requirement set 1.8)</span></span>

<span data-ttu-id="11dd9-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="11dd9-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="11dd9-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="11dd9-106">Requirements</span></span>

|<span data-ttu-id="11dd9-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="11dd9-107">Requirement</span></span>| <span data-ttu-id="11dd9-108">Valor</span><span class="sxs-lookup"><span data-stu-id="11dd9-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="11dd9-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="11dd9-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="11dd9-110">1.1</span><span class="sxs-lookup"><span data-stu-id="11dd9-110">1.1</span></span>|
|[<span data-ttu-id="11dd9-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="11dd9-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="11dd9-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="11dd9-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="11dd9-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="11dd9-113">Properties</span></span>

| <span data-ttu-id="11dd9-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="11dd9-114">Property</span></span> | <span data-ttu-id="11dd9-115">Modos</span><span class="sxs-lookup"><span data-stu-id="11dd9-115">Modes</span></span> | <span data-ttu-id="11dd9-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="11dd9-116">Return type</span></span> | <span data-ttu-id="11dd9-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="11dd9-117">Minimum</span></span><br><span data-ttu-id="11dd9-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="11dd9-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="11dd9-119">context</span><span class="sxs-lookup"><span data-stu-id="11dd9-119">context</span></span>](office.context.md) | <span data-ttu-id="11dd9-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="11dd9-120">Compose</span></span><br><span data-ttu-id="11dd9-121">Ler</span><span class="sxs-lookup"><span data-stu-id="11dd9-121">Read</span></span> | [<span data-ttu-id="11dd9-122">Context</span><span class="sxs-lookup"><span data-stu-id="11dd9-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="11dd9-123">1.1</span><span class="sxs-lookup"><span data-stu-id="11dd9-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="11dd9-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="11dd9-124">Enumerations</span></span>

| <span data-ttu-id="11dd9-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="11dd9-125">Enumeration</span></span> | <span data-ttu-id="11dd9-126">Modos</span><span class="sxs-lookup"><span data-stu-id="11dd9-126">Modes</span></span> | <span data-ttu-id="11dd9-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="11dd9-127">Return type</span></span> | <span data-ttu-id="11dd9-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="11dd9-128">Minimum</span></span><br><span data-ttu-id="11dd9-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="11dd9-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="11dd9-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="11dd9-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="11dd9-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="11dd9-131">Compose</span></span><br><span data-ttu-id="11dd9-132">Ler</span><span class="sxs-lookup"><span data-stu-id="11dd9-132">Read</span></span> | <span data-ttu-id="11dd9-133">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-133">String</span></span> | [<span data-ttu-id="11dd9-134">1.1</span><span class="sxs-lookup"><span data-stu-id="11dd9-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="11dd9-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="11dd9-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="11dd9-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="11dd9-136">Compose</span></span><br><span data-ttu-id="11dd9-137">Ler</span><span class="sxs-lookup"><span data-stu-id="11dd9-137">Read</span></span> | <span data-ttu-id="11dd9-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-138">String</span></span> | [<span data-ttu-id="11dd9-139">1.1</span><span class="sxs-lookup"><span data-stu-id="11dd9-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="11dd9-140">EventType</span><span class="sxs-lookup"><span data-stu-id="11dd9-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="11dd9-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="11dd9-141">Compose</span></span><br><span data-ttu-id="11dd9-142">Ler</span><span class="sxs-lookup"><span data-stu-id="11dd9-142">Read</span></span> | <span data-ttu-id="11dd9-143">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-143">String</span></span> | [<span data-ttu-id="11dd9-144">1.5</span><span class="sxs-lookup"><span data-stu-id="11dd9-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="11dd9-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="11dd9-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="11dd9-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="11dd9-146">Compose</span></span><br><span data-ttu-id="11dd9-147">Ler</span><span class="sxs-lookup"><span data-stu-id="11dd9-147">Read</span></span> | <span data-ttu-id="11dd9-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-148">String</span></span> | [<span data-ttu-id="11dd9-149">1.1</span><span class="sxs-lookup"><span data-stu-id="11dd9-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="11dd9-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="11dd9-150">Namespaces</span></span>

<span data-ttu-id="11dd9-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true): inclui várias enumerações específicas Outlook, por exemplo, `ItemType` , , , , , e `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="11dd9-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="11dd9-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="11dd9-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="11dd9-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="11dd9-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="11dd9-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="11dd9-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="11dd9-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="11dd9-155">Type</span></span>

*   <span data-ttu-id="11dd9-156">String</span><span class="sxs-lookup"><span data-stu-id="11dd9-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="11dd9-157">Propriedades</span><span class="sxs-lookup"><span data-stu-id="11dd9-157">Properties</span></span>

|<span data-ttu-id="11dd9-158">Nome</span><span class="sxs-lookup"><span data-stu-id="11dd9-158">Name</span></span>| <span data-ttu-id="11dd9-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="11dd9-159">Type</span></span>| <span data-ttu-id="11dd9-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="11dd9-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="11dd9-161">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-161">String</span></span>|<span data-ttu-id="11dd9-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="11dd9-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="11dd9-163">String</span><span class="sxs-lookup"><span data-stu-id="11dd9-163">String</span></span>|<span data-ttu-id="11dd9-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="11dd9-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="11dd9-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="11dd9-165">Requirements</span></span>

|<span data-ttu-id="11dd9-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="11dd9-166">Requirement</span></span>| <span data-ttu-id="11dd9-167">Valor</span><span class="sxs-lookup"><span data-stu-id="11dd9-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="11dd9-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="11dd9-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="11dd9-169">1.1</span><span class="sxs-lookup"><span data-stu-id="11dd9-169">1.1</span></span>|
|[<span data-ttu-id="11dd9-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="11dd9-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="11dd9-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="11dd9-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="11dd9-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="11dd9-172">CoercionType: String</span></span>

<span data-ttu-id="11dd9-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="11dd9-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="11dd9-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="11dd9-174">Type</span></span>

*   <span data-ttu-id="11dd9-175">String</span><span class="sxs-lookup"><span data-stu-id="11dd9-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="11dd9-176">Propriedades</span><span class="sxs-lookup"><span data-stu-id="11dd9-176">Properties</span></span>

|<span data-ttu-id="11dd9-177">Nome</span><span class="sxs-lookup"><span data-stu-id="11dd9-177">Name</span></span>| <span data-ttu-id="11dd9-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="11dd9-178">Type</span></span>| <span data-ttu-id="11dd9-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="11dd9-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="11dd9-180">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-180">String</span></span>|<span data-ttu-id="11dd9-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="11dd9-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="11dd9-182">String</span><span class="sxs-lookup"><span data-stu-id="11dd9-182">String</span></span>|<span data-ttu-id="11dd9-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="11dd9-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="11dd9-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="11dd9-184">Requirements</span></span>

|<span data-ttu-id="11dd9-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="11dd9-185">Requirement</span></span>| <span data-ttu-id="11dd9-186">Valor</span><span class="sxs-lookup"><span data-stu-id="11dd9-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="11dd9-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="11dd9-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="11dd9-188">1.1</span><span class="sxs-lookup"><span data-stu-id="11dd9-188">1.1</span></span>|
|[<span data-ttu-id="11dd9-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="11dd9-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="11dd9-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="11dd9-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="11dd9-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="11dd9-191">EventType: String</span></span>

<span data-ttu-id="11dd9-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="11dd9-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="11dd9-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="11dd9-193">Type</span></span>

*   <span data-ttu-id="11dd9-194">String</span><span class="sxs-lookup"><span data-stu-id="11dd9-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="11dd9-195">Propriedades</span><span class="sxs-lookup"><span data-stu-id="11dd9-195">Properties</span></span>

| <span data-ttu-id="11dd9-196">Nome</span><span class="sxs-lookup"><span data-stu-id="11dd9-196">Name</span></span> | <span data-ttu-id="11dd9-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="11dd9-197">Type</span></span> | <span data-ttu-id="11dd9-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="11dd9-198">Description</span></span> | <span data-ttu-id="11dd9-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="11dd9-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="11dd9-200">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-200">String</span></span> | <span data-ttu-id="11dd9-201">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="11dd9-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="11dd9-202">1.7</span><span class="sxs-lookup"><span data-stu-id="11dd9-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="11dd9-203">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-203">String</span></span> | <span data-ttu-id="11dd9-204">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="11dd9-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="11dd9-205">1,8</span><span class="sxs-lookup"><span data-stu-id="11dd9-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="11dd9-206">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-206">String</span></span> | <span data-ttu-id="11dd9-207">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="11dd9-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="11dd9-208">1,8</span><span class="sxs-lookup"><span data-stu-id="11dd9-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="11dd9-209">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-209">String</span></span> | <span data-ttu-id="11dd9-210">Um item Outlook diferente é selecionado para exibição enquanto o painel de tarefas é fixado.</span><span class="sxs-lookup"><span data-stu-id="11dd9-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="11dd9-211">1,5</span><span class="sxs-lookup"><span data-stu-id="11dd9-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="11dd9-212">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-212">String</span></span> | <span data-ttu-id="11dd9-213">A lista de destinatários do item ou local do compromisso selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="11dd9-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="11dd9-214">1.7</span><span class="sxs-lookup"><span data-stu-id="11dd9-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="11dd9-215">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-215">String</span></span> | <span data-ttu-id="11dd9-216">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="11dd9-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="11dd9-217">1.7</span><span class="sxs-lookup"><span data-stu-id="11dd9-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="11dd9-218">Requisitos</span><span class="sxs-lookup"><span data-stu-id="11dd9-218">Requirements</span></span>

|<span data-ttu-id="11dd9-219">Requisito</span><span class="sxs-lookup"><span data-stu-id="11dd9-219">Requirement</span></span>| <span data-ttu-id="11dd9-220">Valor</span><span class="sxs-lookup"><span data-stu-id="11dd9-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="11dd9-221">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="11dd9-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="11dd9-222">1,5</span><span class="sxs-lookup"><span data-stu-id="11dd9-222">1.5</span></span> |
|[<span data-ttu-id="11dd9-223">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="11dd9-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="11dd9-224">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="11dd9-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="11dd9-225">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="11dd9-225">SourceProperty: String</span></span>

<span data-ttu-id="11dd9-226">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="11dd9-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="11dd9-227">Tipo</span><span class="sxs-lookup"><span data-stu-id="11dd9-227">Type</span></span>

*   <span data-ttu-id="11dd9-228">String</span><span class="sxs-lookup"><span data-stu-id="11dd9-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="11dd9-229">Propriedades</span><span class="sxs-lookup"><span data-stu-id="11dd9-229">Properties</span></span>

|<span data-ttu-id="11dd9-230">Nome</span><span class="sxs-lookup"><span data-stu-id="11dd9-230">Name</span></span>| <span data-ttu-id="11dd9-231">Tipo</span><span class="sxs-lookup"><span data-stu-id="11dd9-231">Type</span></span>| <span data-ttu-id="11dd9-232">Descrição</span><span class="sxs-lookup"><span data-stu-id="11dd9-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="11dd9-233">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="11dd9-233">String</span></span>|<span data-ttu-id="11dd9-234">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="11dd9-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="11dd9-235">String</span><span class="sxs-lookup"><span data-stu-id="11dd9-235">String</span></span>|<span data-ttu-id="11dd9-236">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="11dd9-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="11dd9-237">Requisitos</span><span class="sxs-lookup"><span data-stu-id="11dd9-237">Requirements</span></span>

|<span data-ttu-id="11dd9-238">Requisito</span><span class="sxs-lookup"><span data-stu-id="11dd9-238">Requirement</span></span>| <span data-ttu-id="11dd9-239">Valor</span><span class="sxs-lookup"><span data-stu-id="11dd9-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="11dd9-240">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="11dd9-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="11dd9-241">1.1</span><span class="sxs-lookup"><span data-stu-id="11dd9-241">1.1</span></span>|
|[<span data-ttu-id="11dd9-242">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="11dd9-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="11dd9-243">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="11dd9-243">Compose or Read</span></span>|
