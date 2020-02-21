---
title: Namespace do Office – conjunto de requisitos de visualização
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 2cd04cc6d333439a679803e39357e4d19c550f95
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165507"
---
# <a name="office"></a><span data-ttu-id="370b6-102">Office</span><span class="sxs-lookup"><span data-stu-id="370b6-102">Office</span></span>

<span data-ttu-id="370b6-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="370b6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="370b6-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="370b6-105">Requirements</span></span>

|<span data-ttu-id="370b6-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="370b6-106">Requirement</span></span>| <span data-ttu-id="370b6-107">Valor</span><span class="sxs-lookup"><span data-stu-id="370b6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="370b6-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="370b6-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="370b6-109">1.1</span><span class="sxs-lookup"><span data-stu-id="370b6-109">1.1</span></span>|
|[<span data-ttu-id="370b6-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="370b6-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="370b6-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="370b6-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="370b6-112">Propriedades</span><span class="sxs-lookup"><span data-stu-id="370b6-112">Properties</span></span>

| <span data-ttu-id="370b6-113">Propriedade</span><span class="sxs-lookup"><span data-stu-id="370b6-113">Property</span></span> | <span data-ttu-id="370b6-114">Modelos</span><span class="sxs-lookup"><span data-stu-id="370b6-114">Modes</span></span> | <span data-ttu-id="370b6-115">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="370b6-115">Return type</span></span> | <span data-ttu-id="370b6-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="370b6-116">Minimum</span></span><br><span data-ttu-id="370b6-117">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="370b6-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="370b6-118">context</span><span class="sxs-lookup"><span data-stu-id="370b6-118">context</span></span>](office.context.md) | <span data-ttu-id="370b6-119">Escrever</span><span class="sxs-lookup"><span data-stu-id="370b6-119">Compose</span></span><br><span data-ttu-id="370b6-120">Ler</span><span class="sxs-lookup"><span data-stu-id="370b6-120">Read</span></span> | [<span data-ttu-id="370b6-121">Context</span><span class="sxs-lookup"><span data-stu-id="370b6-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="370b6-122">1.1</span><span class="sxs-lookup"><span data-stu-id="370b6-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="370b6-123">Enumerações</span><span class="sxs-lookup"><span data-stu-id="370b6-123">Enumerations</span></span>

| <span data-ttu-id="370b6-124">Enumeração</span><span class="sxs-lookup"><span data-stu-id="370b6-124">Enumeration</span></span> | <span data-ttu-id="370b6-125">Modelos</span><span class="sxs-lookup"><span data-stu-id="370b6-125">Modes</span></span> | <span data-ttu-id="370b6-126">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="370b6-126">Return type</span></span> | <span data-ttu-id="370b6-127">Mínimo</span><span class="sxs-lookup"><span data-stu-id="370b6-127">Minimum</span></span><br><span data-ttu-id="370b6-128">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="370b6-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="370b6-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="370b6-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="370b6-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="370b6-130">Compose</span></span><br><span data-ttu-id="370b6-131">Ler</span><span class="sxs-lookup"><span data-stu-id="370b6-131">Read</span></span> | <span data-ttu-id="370b6-132">String</span><span class="sxs-lookup"><span data-stu-id="370b6-132">String</span></span> | [<span data-ttu-id="370b6-133">1.1</span><span class="sxs-lookup"><span data-stu-id="370b6-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="370b6-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="370b6-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="370b6-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="370b6-135">Compose</span></span><br><span data-ttu-id="370b6-136">Ler</span><span class="sxs-lookup"><span data-stu-id="370b6-136">Read</span></span> | <span data-ttu-id="370b6-137">String</span><span class="sxs-lookup"><span data-stu-id="370b6-137">String</span></span> | [<span data-ttu-id="370b6-138">1.1</span><span class="sxs-lookup"><span data-stu-id="370b6-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="370b6-139">EventType</span><span class="sxs-lookup"><span data-stu-id="370b6-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="370b6-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="370b6-140">Compose</span></span><br><span data-ttu-id="370b6-141">Ler</span><span class="sxs-lookup"><span data-stu-id="370b6-141">Read</span></span> | <span data-ttu-id="370b6-142">String</span><span class="sxs-lookup"><span data-stu-id="370b6-142">String</span></span> | [<span data-ttu-id="370b6-143">1,5</span><span class="sxs-lookup"><span data-stu-id="370b6-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="370b6-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="370b6-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="370b6-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="370b6-145">Compose</span></span><br><span data-ttu-id="370b6-146">Ler</span><span class="sxs-lookup"><span data-stu-id="370b6-146">Read</span></span> | <span data-ttu-id="370b6-147">String</span><span class="sxs-lookup"><span data-stu-id="370b6-147">String</span></span> | [<span data-ttu-id="370b6-148">1.1</span><span class="sxs-lookup"><span data-stu-id="370b6-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="370b6-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="370b6-149">Namespaces</span></span>

<span data-ttu-id="370b6-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="370b6-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="370b6-151">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="370b6-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="370b6-152">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="370b6-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="370b6-153">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="370b6-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="370b6-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="370b6-154">Type</span></span>

*   <span data-ttu-id="370b6-155">String</span><span class="sxs-lookup"><span data-stu-id="370b6-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="370b6-156">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="370b6-156">Properties:</span></span>

|<span data-ttu-id="370b6-157">Nome</span><span class="sxs-lookup"><span data-stu-id="370b6-157">Name</span></span>| <span data-ttu-id="370b6-158">Tipo</span><span class="sxs-lookup"><span data-stu-id="370b6-158">Type</span></span>| <span data-ttu-id="370b6-159">Descrição</span><span class="sxs-lookup"><span data-stu-id="370b6-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="370b6-160">String</span><span class="sxs-lookup"><span data-stu-id="370b6-160">String</span></span>|<span data-ttu-id="370b6-161">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="370b6-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="370b6-162">String</span><span class="sxs-lookup"><span data-stu-id="370b6-162">String</span></span>|<span data-ttu-id="370b6-163">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="370b6-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="370b6-164">Requisitos</span><span class="sxs-lookup"><span data-stu-id="370b6-164">Requirements</span></span>

|<span data-ttu-id="370b6-165">Requisito</span><span class="sxs-lookup"><span data-stu-id="370b6-165">Requirement</span></span>| <span data-ttu-id="370b6-166">Valor</span><span class="sxs-lookup"><span data-stu-id="370b6-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="370b6-167">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="370b6-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="370b6-168">1.1</span><span class="sxs-lookup"><span data-stu-id="370b6-168">1.1</span></span>|
|[<span data-ttu-id="370b6-169">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="370b6-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="370b6-170">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="370b6-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="370b6-171">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="370b6-171">CoercionType: String</span></span>

<span data-ttu-id="370b6-172">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="370b6-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="370b6-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="370b6-173">Type</span></span>

*   <span data-ttu-id="370b6-174">String</span><span class="sxs-lookup"><span data-stu-id="370b6-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="370b6-175">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="370b6-175">Properties:</span></span>

|<span data-ttu-id="370b6-176">Nome</span><span class="sxs-lookup"><span data-stu-id="370b6-176">Name</span></span>| <span data-ttu-id="370b6-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="370b6-177">Type</span></span>| <span data-ttu-id="370b6-178">Descrição</span><span class="sxs-lookup"><span data-stu-id="370b6-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="370b6-179">String</span><span class="sxs-lookup"><span data-stu-id="370b6-179">String</span></span>|<span data-ttu-id="370b6-180">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="370b6-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="370b6-181">String</span><span class="sxs-lookup"><span data-stu-id="370b6-181">String</span></span>|<span data-ttu-id="370b6-182">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="370b6-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="370b6-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="370b6-183">Requirements</span></span>

|<span data-ttu-id="370b6-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="370b6-184">Requirement</span></span>| <span data-ttu-id="370b6-185">Valor</span><span class="sxs-lookup"><span data-stu-id="370b6-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="370b6-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="370b6-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="370b6-187">1.1</span><span class="sxs-lookup"><span data-stu-id="370b6-187">1.1</span></span>|
|[<span data-ttu-id="370b6-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="370b6-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="370b6-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="370b6-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="370b6-190">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="370b6-190">EventType: String</span></span>

<span data-ttu-id="370b6-191">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="370b6-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="370b6-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="370b6-192">Type</span></span>

*   <span data-ttu-id="370b6-193">String</span><span class="sxs-lookup"><span data-stu-id="370b6-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="370b6-194">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="370b6-194">Properties:</span></span>

| <span data-ttu-id="370b6-195">Nome</span><span class="sxs-lookup"><span data-stu-id="370b6-195">Name</span></span> | <span data-ttu-id="370b6-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="370b6-196">Type</span></span> | <span data-ttu-id="370b6-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="370b6-197">Description</span></span> | <span data-ttu-id="370b6-198">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="370b6-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="370b6-199">String</span><span class="sxs-lookup"><span data-stu-id="370b6-199">String</span></span> | <span data-ttu-id="370b6-200">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="370b6-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="370b6-201">1.7</span><span class="sxs-lookup"><span data-stu-id="370b6-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="370b6-202">String</span><span class="sxs-lookup"><span data-stu-id="370b6-202">String</span></span> | <span data-ttu-id="370b6-203">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="370b6-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="370b6-204">1,8</span><span class="sxs-lookup"><span data-stu-id="370b6-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="370b6-205">String</span><span class="sxs-lookup"><span data-stu-id="370b6-205">String</span></span> | <span data-ttu-id="370b6-206">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="370b6-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="370b6-207">1,8</span><span class="sxs-lookup"><span data-stu-id="370b6-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="370b6-208">String</span><span class="sxs-lookup"><span data-stu-id="370b6-208">String</span></span> | <span data-ttu-id="370b6-209">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="370b6-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="370b6-210">1,5</span><span class="sxs-lookup"><span data-stu-id="370b6-210">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="370b6-211">String</span><span class="sxs-lookup"><span data-stu-id="370b6-211">String</span></span> | <span data-ttu-id="370b6-212">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="370b6-212">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="370b6-213">Visualização</span><span class="sxs-lookup"><span data-stu-id="370b6-213">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="370b6-214">String</span><span class="sxs-lookup"><span data-stu-id="370b6-214">String</span></span> | <span data-ttu-id="370b6-215">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="370b6-215">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="370b6-216">1.7</span><span class="sxs-lookup"><span data-stu-id="370b6-216">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="370b6-217">String</span><span class="sxs-lookup"><span data-stu-id="370b6-217">String</span></span> | <span data-ttu-id="370b6-218">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="370b6-218">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="370b6-219">1.7</span><span class="sxs-lookup"><span data-stu-id="370b6-219">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="370b6-220">Requisitos</span><span class="sxs-lookup"><span data-stu-id="370b6-220">Requirements</span></span>

|<span data-ttu-id="370b6-221">Requisito</span><span class="sxs-lookup"><span data-stu-id="370b6-221">Requirement</span></span>| <span data-ttu-id="370b6-222">Valor</span><span class="sxs-lookup"><span data-stu-id="370b6-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="370b6-223">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="370b6-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="370b6-224">1,5</span><span class="sxs-lookup"><span data-stu-id="370b6-224">1.5</span></span> |
|[<span data-ttu-id="370b6-225">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="370b6-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="370b6-226">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="370b6-226">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="370b6-227">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="370b6-227">SourceProperty: String</span></span>

<span data-ttu-id="370b6-228">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="370b6-228">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="370b6-229">Tipo</span><span class="sxs-lookup"><span data-stu-id="370b6-229">Type</span></span>

*   <span data-ttu-id="370b6-230">String</span><span class="sxs-lookup"><span data-stu-id="370b6-230">String</span></span>

##### <a name="properties"></a><span data-ttu-id="370b6-231">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="370b6-231">Properties:</span></span>

|<span data-ttu-id="370b6-232">Nome</span><span class="sxs-lookup"><span data-stu-id="370b6-232">Name</span></span>| <span data-ttu-id="370b6-233">Tipo</span><span class="sxs-lookup"><span data-stu-id="370b6-233">Type</span></span>| <span data-ttu-id="370b6-234">Descrição</span><span class="sxs-lookup"><span data-stu-id="370b6-234">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="370b6-235">String</span><span class="sxs-lookup"><span data-stu-id="370b6-235">String</span></span>|<span data-ttu-id="370b6-236">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="370b6-236">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="370b6-237">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="370b6-237">String</span></span>|<span data-ttu-id="370b6-238">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="370b6-238">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="370b6-239">Requisitos</span><span class="sxs-lookup"><span data-stu-id="370b6-239">Requirements</span></span>

|<span data-ttu-id="370b6-240">Requisito</span><span class="sxs-lookup"><span data-stu-id="370b6-240">Requirement</span></span>| <span data-ttu-id="370b6-241">Valor</span><span class="sxs-lookup"><span data-stu-id="370b6-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="370b6-242">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="370b6-242">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="370b6-243">1.1</span><span class="sxs-lookup"><span data-stu-id="370b6-243">1.1</span></span>|
|[<span data-ttu-id="370b6-244">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="370b6-244">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="370b6-245">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="370b6-245">Compose or Read</span></span>|
