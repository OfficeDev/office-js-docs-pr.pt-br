---
title: Namespace do Office – conjunto de requisitos de visualização
description: Membros do namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de visualização da API da caixa de correio.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: d72e5c78a7fd8d3c00b8f84e7d9b05ee6defc0c5
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890855"
---
# <a name="office-mailbox-preview-requirement-set"></a><span data-ttu-id="edeae-103">Office (conjunto de requisitos de visualização da caixa de correio)</span><span class="sxs-lookup"><span data-stu-id="edeae-103">Office (Mailbox preview requirement set)</span></span>

<span data-ttu-id="edeae-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="edeae-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="edeae-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="edeae-106">Requirements</span></span>

|<span data-ttu-id="edeae-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="edeae-107">Requirement</span></span>| <span data-ttu-id="edeae-108">Valor</span><span class="sxs-lookup"><span data-stu-id="edeae-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="edeae-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="edeae-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="edeae-110">1.1</span><span class="sxs-lookup"><span data-stu-id="edeae-110">1.1</span></span>|
|[<span data-ttu-id="edeae-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="edeae-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="edeae-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="edeae-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="edeae-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="edeae-113">Properties</span></span>

| <span data-ttu-id="edeae-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="edeae-114">Property</span></span> | <span data-ttu-id="edeae-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="edeae-115">Modes</span></span> | <span data-ttu-id="edeae-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="edeae-116">Return type</span></span> | <span data-ttu-id="edeae-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="edeae-117">Minimum</span></span><br><span data-ttu-id="edeae-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="edeae-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="edeae-119">context</span><span class="sxs-lookup"><span data-stu-id="edeae-119">context</span></span>](office.context.md) | <span data-ttu-id="edeae-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="edeae-120">Compose</span></span><br><span data-ttu-id="edeae-121">Ler</span><span class="sxs-lookup"><span data-stu-id="edeae-121">Read</span></span> | [<span data-ttu-id="edeae-122">Context</span><span class="sxs-lookup"><span data-stu-id="edeae-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="edeae-123">1.1</span><span class="sxs-lookup"><span data-stu-id="edeae-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="edeae-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="edeae-124">Enumerations</span></span>

| <span data-ttu-id="edeae-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="edeae-125">Enumeration</span></span> | <span data-ttu-id="edeae-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="edeae-126">Modes</span></span> | <span data-ttu-id="edeae-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="edeae-127">Return type</span></span> | <span data-ttu-id="edeae-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="edeae-128">Minimum</span></span><br><span data-ttu-id="edeae-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="edeae-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="edeae-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="edeae-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="edeae-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="edeae-131">Compose</span></span><br><span data-ttu-id="edeae-132">Ler</span><span class="sxs-lookup"><span data-stu-id="edeae-132">Read</span></span> | <span data-ttu-id="edeae-133">String</span><span class="sxs-lookup"><span data-stu-id="edeae-133">String</span></span> | [<span data-ttu-id="edeae-134">1.1</span><span class="sxs-lookup"><span data-stu-id="edeae-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="edeae-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="edeae-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="edeae-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="edeae-136">Compose</span></span><br><span data-ttu-id="edeae-137">Ler</span><span class="sxs-lookup"><span data-stu-id="edeae-137">Read</span></span> | <span data-ttu-id="edeae-138">String</span><span class="sxs-lookup"><span data-stu-id="edeae-138">String</span></span> | [<span data-ttu-id="edeae-139">1.1</span><span class="sxs-lookup"><span data-stu-id="edeae-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="edeae-140">EventType</span><span class="sxs-lookup"><span data-stu-id="edeae-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="edeae-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="edeae-141">Compose</span></span><br><span data-ttu-id="edeae-142">Ler</span><span class="sxs-lookup"><span data-stu-id="edeae-142">Read</span></span> | <span data-ttu-id="edeae-143">String</span><span class="sxs-lookup"><span data-stu-id="edeae-143">String</span></span> | [<span data-ttu-id="edeae-144">1,5</span><span class="sxs-lookup"><span data-stu-id="edeae-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="edeae-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="edeae-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="edeae-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="edeae-146">Compose</span></span><br><span data-ttu-id="edeae-147">Ler</span><span class="sxs-lookup"><span data-stu-id="edeae-147">Read</span></span> | <span data-ttu-id="edeae-148">String</span><span class="sxs-lookup"><span data-stu-id="edeae-148">String</span></span> | [<span data-ttu-id="edeae-149">1.1</span><span class="sxs-lookup"><span data-stu-id="edeae-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="edeae-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="edeae-150">Namespaces</span></span>

<span data-ttu-id="edeae-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="edeae-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="edeae-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="edeae-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="edeae-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="edeae-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="edeae-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="edeae-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="edeae-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="edeae-155">Type</span></span>

*   <span data-ttu-id="edeae-156">String</span><span class="sxs-lookup"><span data-stu-id="edeae-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="edeae-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="edeae-157">Properties:</span></span>

|<span data-ttu-id="edeae-158">Nome</span><span class="sxs-lookup"><span data-stu-id="edeae-158">Name</span></span>| <span data-ttu-id="edeae-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="edeae-159">Type</span></span>| <span data-ttu-id="edeae-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="edeae-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="edeae-161">String</span><span class="sxs-lookup"><span data-stu-id="edeae-161">String</span></span>|<span data-ttu-id="edeae-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="edeae-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="edeae-163">String</span><span class="sxs-lookup"><span data-stu-id="edeae-163">String</span></span>|<span data-ttu-id="edeae-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="edeae-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="edeae-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="edeae-165">Requirements</span></span>

|<span data-ttu-id="edeae-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="edeae-166">Requirement</span></span>| <span data-ttu-id="edeae-167">Valor</span><span class="sxs-lookup"><span data-stu-id="edeae-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="edeae-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="edeae-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="edeae-169">1.1</span><span class="sxs-lookup"><span data-stu-id="edeae-169">1.1</span></span>|
|[<span data-ttu-id="edeae-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="edeae-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="edeae-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="edeae-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="edeae-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="edeae-172">CoercionType: String</span></span>

<span data-ttu-id="edeae-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="edeae-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="edeae-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="edeae-174">Type</span></span>

*   <span data-ttu-id="edeae-175">String</span><span class="sxs-lookup"><span data-stu-id="edeae-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="edeae-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="edeae-176">Properties:</span></span>

|<span data-ttu-id="edeae-177">Nome</span><span class="sxs-lookup"><span data-stu-id="edeae-177">Name</span></span>| <span data-ttu-id="edeae-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="edeae-178">Type</span></span>| <span data-ttu-id="edeae-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="edeae-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="edeae-180">String</span><span class="sxs-lookup"><span data-stu-id="edeae-180">String</span></span>|<span data-ttu-id="edeae-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="edeae-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="edeae-182">String</span><span class="sxs-lookup"><span data-stu-id="edeae-182">String</span></span>|<span data-ttu-id="edeae-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="edeae-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="edeae-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="edeae-184">Requirements</span></span>

|<span data-ttu-id="edeae-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="edeae-185">Requirement</span></span>| <span data-ttu-id="edeae-186">Valor</span><span class="sxs-lookup"><span data-stu-id="edeae-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="edeae-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="edeae-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="edeae-188">1.1</span><span class="sxs-lookup"><span data-stu-id="edeae-188">1.1</span></span>|
|[<span data-ttu-id="edeae-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="edeae-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="edeae-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="edeae-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="edeae-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="edeae-191">EventType: String</span></span>

<span data-ttu-id="edeae-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="edeae-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="edeae-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="edeae-193">Type</span></span>

*   <span data-ttu-id="edeae-194">String</span><span class="sxs-lookup"><span data-stu-id="edeae-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="edeae-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="edeae-195">Properties:</span></span>

| <span data-ttu-id="edeae-196">Nome</span><span class="sxs-lookup"><span data-stu-id="edeae-196">Name</span></span> | <span data-ttu-id="edeae-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="edeae-197">Type</span></span> | <span data-ttu-id="edeae-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="edeae-198">Description</span></span> | <span data-ttu-id="edeae-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="edeae-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="edeae-200">String</span><span class="sxs-lookup"><span data-stu-id="edeae-200">String</span></span> | <span data-ttu-id="edeae-201">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="edeae-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="edeae-202">1.7</span><span class="sxs-lookup"><span data-stu-id="edeae-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="edeae-203">String</span><span class="sxs-lookup"><span data-stu-id="edeae-203">String</span></span> | <span data-ttu-id="edeae-204">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="edeae-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="edeae-205">1,8</span><span class="sxs-lookup"><span data-stu-id="edeae-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="edeae-206">String</span><span class="sxs-lookup"><span data-stu-id="edeae-206">String</span></span> | <span data-ttu-id="edeae-207">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="edeae-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="edeae-208">1,8</span><span class="sxs-lookup"><span data-stu-id="edeae-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="edeae-209">String</span><span class="sxs-lookup"><span data-stu-id="edeae-209">String</span></span> | <span data-ttu-id="edeae-210">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="edeae-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="edeae-211">1,5</span><span class="sxs-lookup"><span data-stu-id="edeae-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="edeae-212">String</span><span class="sxs-lookup"><span data-stu-id="edeae-212">String</span></span> | <span data-ttu-id="edeae-213">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="edeae-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="edeae-214">Visualização</span><span class="sxs-lookup"><span data-stu-id="edeae-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="edeae-215">String</span><span class="sxs-lookup"><span data-stu-id="edeae-215">String</span></span> | <span data-ttu-id="edeae-216">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="edeae-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="edeae-217">1.7</span><span class="sxs-lookup"><span data-stu-id="edeae-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="edeae-218">String</span><span class="sxs-lookup"><span data-stu-id="edeae-218">String</span></span> | <span data-ttu-id="edeae-219">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="edeae-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="edeae-220">1.7</span><span class="sxs-lookup"><span data-stu-id="edeae-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="edeae-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="edeae-221">Requirements</span></span>

|<span data-ttu-id="edeae-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="edeae-222">Requirement</span></span>| <span data-ttu-id="edeae-223">Valor</span><span class="sxs-lookup"><span data-stu-id="edeae-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="edeae-224">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="edeae-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="edeae-225">1,5</span><span class="sxs-lookup"><span data-stu-id="edeae-225">1.5</span></span> |
|[<span data-ttu-id="edeae-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="edeae-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="edeae-227">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="edeae-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="edeae-228">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="edeae-228">SourceProperty: String</span></span>

<span data-ttu-id="edeae-229">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="edeae-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="edeae-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="edeae-230">Type</span></span>

*   <span data-ttu-id="edeae-231">String</span><span class="sxs-lookup"><span data-stu-id="edeae-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="edeae-232">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="edeae-232">Properties:</span></span>

|<span data-ttu-id="edeae-233">Nome</span><span class="sxs-lookup"><span data-stu-id="edeae-233">Name</span></span>| <span data-ttu-id="edeae-234">Tipo</span><span class="sxs-lookup"><span data-stu-id="edeae-234">Type</span></span>| <span data-ttu-id="edeae-235">Descrição</span><span class="sxs-lookup"><span data-stu-id="edeae-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="edeae-236">String</span><span class="sxs-lookup"><span data-stu-id="edeae-236">String</span></span>|<span data-ttu-id="edeae-237">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="edeae-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="edeae-238">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="edeae-238">String</span></span>|<span data-ttu-id="edeae-239">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="edeae-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="edeae-240">Requisitos</span><span class="sxs-lookup"><span data-stu-id="edeae-240">Requirements</span></span>

|<span data-ttu-id="edeae-241">Requisito</span><span class="sxs-lookup"><span data-stu-id="edeae-241">Requirement</span></span>| <span data-ttu-id="edeae-242">Valor</span><span class="sxs-lookup"><span data-stu-id="edeae-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="edeae-243">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="edeae-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="edeae-244">1.1</span><span class="sxs-lookup"><span data-stu-id="edeae-244">1.1</span></span>|
|[<span data-ttu-id="edeae-245">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="edeae-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="edeae-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="edeae-246">Compose or Read</span></span>|
