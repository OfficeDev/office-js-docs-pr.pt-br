---
title: Namespace do Office – conjunto de requisitos de visualização
description: O modelo de objeto para o namespace de nível superior da API de suplementos do Outlook (versão prévia da API da caixa de correio).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 40623c02fae820926d9162903320f30e5a424544
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720271"
---
# <a name="office"></a><span data-ttu-id="4954a-103">Office</span><span class="sxs-lookup"><span data-stu-id="4954a-103">Office</span></span>

<span data-ttu-id="4954a-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="4954a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4954a-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4954a-106">Requirements</span></span>

|<span data-ttu-id="4954a-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="4954a-107">Requirement</span></span>| <span data-ttu-id="4954a-108">Valor</span><span class="sxs-lookup"><span data-stu-id="4954a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4954a-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4954a-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4954a-110">1.1</span><span class="sxs-lookup"><span data-stu-id="4954a-110">1.1</span></span>|
|[<span data-ttu-id="4954a-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4954a-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4954a-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4954a-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="4954a-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="4954a-113">Properties</span></span>

| <span data-ttu-id="4954a-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="4954a-114">Property</span></span> | <span data-ttu-id="4954a-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="4954a-115">Modes</span></span> | <span data-ttu-id="4954a-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="4954a-116">Return type</span></span> | <span data-ttu-id="4954a-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="4954a-117">Minimum</span></span><br><span data-ttu-id="4954a-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="4954a-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4954a-119">context</span><span class="sxs-lookup"><span data-stu-id="4954a-119">context</span></span>](office.context.md) | <span data-ttu-id="4954a-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="4954a-120">Compose</span></span><br><span data-ttu-id="4954a-121">Ler</span><span class="sxs-lookup"><span data-stu-id="4954a-121">Read</span></span> | [<span data-ttu-id="4954a-122">Context</span><span class="sxs-lookup"><span data-stu-id="4954a-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="4954a-123">1.1</span><span class="sxs-lookup"><span data-stu-id="4954a-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="4954a-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="4954a-124">Enumerations</span></span>

| <span data-ttu-id="4954a-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="4954a-125">Enumeration</span></span> | <span data-ttu-id="4954a-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="4954a-126">Modes</span></span> | <span data-ttu-id="4954a-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="4954a-127">Return type</span></span> | <span data-ttu-id="4954a-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="4954a-128">Minimum</span></span><br><span data-ttu-id="4954a-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="4954a-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4954a-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4954a-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4954a-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="4954a-131">Compose</span></span><br><span data-ttu-id="4954a-132">Ler</span><span class="sxs-lookup"><span data-stu-id="4954a-132">Read</span></span> | <span data-ttu-id="4954a-133">String</span><span class="sxs-lookup"><span data-stu-id="4954a-133">String</span></span> | [<span data-ttu-id="4954a-134">1.1</span><span class="sxs-lookup"><span data-stu-id="4954a-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4954a-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4954a-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4954a-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="4954a-136">Compose</span></span><br><span data-ttu-id="4954a-137">Ler</span><span class="sxs-lookup"><span data-stu-id="4954a-137">Read</span></span> | <span data-ttu-id="4954a-138">String</span><span class="sxs-lookup"><span data-stu-id="4954a-138">String</span></span> | [<span data-ttu-id="4954a-139">1.1</span><span class="sxs-lookup"><span data-stu-id="4954a-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4954a-140">EventType</span><span class="sxs-lookup"><span data-stu-id="4954a-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4954a-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="4954a-141">Compose</span></span><br><span data-ttu-id="4954a-142">Ler</span><span class="sxs-lookup"><span data-stu-id="4954a-142">Read</span></span> | <span data-ttu-id="4954a-143">String</span><span class="sxs-lookup"><span data-stu-id="4954a-143">String</span></span> | [<span data-ttu-id="4954a-144">1,5</span><span class="sxs-lookup"><span data-stu-id="4954a-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="4954a-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4954a-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4954a-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="4954a-146">Compose</span></span><br><span data-ttu-id="4954a-147">Ler</span><span class="sxs-lookup"><span data-stu-id="4954a-147">Read</span></span> | <span data-ttu-id="4954a-148">String</span><span class="sxs-lookup"><span data-stu-id="4954a-148">String</span></span> | [<span data-ttu-id="4954a-149">1.1</span><span class="sxs-lookup"><span data-stu-id="4954a-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="4954a-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="4954a-150">Namespaces</span></span>

<span data-ttu-id="4954a-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="4954a-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="4954a-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="4954a-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="4954a-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4954a-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="4954a-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="4954a-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4954a-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="4954a-155">Type</span></span>

*   <span data-ttu-id="4954a-156">String</span><span class="sxs-lookup"><span data-stu-id="4954a-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4954a-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="4954a-157">Properties:</span></span>

|<span data-ttu-id="4954a-158">Nome</span><span class="sxs-lookup"><span data-stu-id="4954a-158">Name</span></span>| <span data-ttu-id="4954a-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="4954a-159">Type</span></span>| <span data-ttu-id="4954a-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="4954a-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4954a-161">String</span><span class="sxs-lookup"><span data-stu-id="4954a-161">String</span></span>|<span data-ttu-id="4954a-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="4954a-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4954a-163">String</span><span class="sxs-lookup"><span data-stu-id="4954a-163">String</span></span>|<span data-ttu-id="4954a-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="4954a-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4954a-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4954a-165">Requirements</span></span>

|<span data-ttu-id="4954a-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="4954a-166">Requirement</span></span>| <span data-ttu-id="4954a-167">Valor</span><span class="sxs-lookup"><span data-stu-id="4954a-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="4954a-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4954a-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4954a-169">1.1</span><span class="sxs-lookup"><span data-stu-id="4954a-169">1.1</span></span>|
|[<span data-ttu-id="4954a-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4954a-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4954a-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4954a-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="4954a-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4954a-172">CoercionType: String</span></span>

<span data-ttu-id="4954a-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="4954a-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4954a-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="4954a-174">Type</span></span>

*   <span data-ttu-id="4954a-175">String</span><span class="sxs-lookup"><span data-stu-id="4954a-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4954a-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="4954a-176">Properties:</span></span>

|<span data-ttu-id="4954a-177">Nome</span><span class="sxs-lookup"><span data-stu-id="4954a-177">Name</span></span>| <span data-ttu-id="4954a-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="4954a-178">Type</span></span>| <span data-ttu-id="4954a-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="4954a-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4954a-180">String</span><span class="sxs-lookup"><span data-stu-id="4954a-180">String</span></span>|<span data-ttu-id="4954a-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="4954a-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4954a-182">String</span><span class="sxs-lookup"><span data-stu-id="4954a-182">String</span></span>|<span data-ttu-id="4954a-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="4954a-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4954a-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4954a-184">Requirements</span></span>

|<span data-ttu-id="4954a-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="4954a-185">Requirement</span></span>| <span data-ttu-id="4954a-186">Valor</span><span class="sxs-lookup"><span data-stu-id="4954a-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="4954a-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4954a-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4954a-188">1.1</span><span class="sxs-lookup"><span data-stu-id="4954a-188">1.1</span></span>|
|[<span data-ttu-id="4954a-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4954a-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4954a-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4954a-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="4954a-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4954a-191">EventType: String</span></span>

<span data-ttu-id="4954a-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="4954a-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4954a-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="4954a-193">Type</span></span>

*   <span data-ttu-id="4954a-194">String</span><span class="sxs-lookup"><span data-stu-id="4954a-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4954a-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="4954a-195">Properties:</span></span>

| <span data-ttu-id="4954a-196">Nome</span><span class="sxs-lookup"><span data-stu-id="4954a-196">Name</span></span> | <span data-ttu-id="4954a-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="4954a-197">Type</span></span> | <span data-ttu-id="4954a-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="4954a-198">Description</span></span> | <span data-ttu-id="4954a-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="4954a-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="4954a-200">String</span><span class="sxs-lookup"><span data-stu-id="4954a-200">String</span></span> | <span data-ttu-id="4954a-201">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="4954a-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="4954a-202">1.7</span><span class="sxs-lookup"><span data-stu-id="4954a-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="4954a-203">String</span><span class="sxs-lookup"><span data-stu-id="4954a-203">String</span></span> | <span data-ttu-id="4954a-204">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="4954a-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="4954a-205">1,8</span><span class="sxs-lookup"><span data-stu-id="4954a-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="4954a-206">String</span><span class="sxs-lookup"><span data-stu-id="4954a-206">String</span></span> | <span data-ttu-id="4954a-207">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="4954a-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="4954a-208">1,8</span><span class="sxs-lookup"><span data-stu-id="4954a-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="4954a-209">String</span><span class="sxs-lookup"><span data-stu-id="4954a-209">String</span></span> | <span data-ttu-id="4954a-210">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="4954a-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="4954a-211">1,5</span><span class="sxs-lookup"><span data-stu-id="4954a-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="4954a-212">String</span><span class="sxs-lookup"><span data-stu-id="4954a-212">String</span></span> | <span data-ttu-id="4954a-213">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="4954a-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="4954a-214">Visualização</span><span class="sxs-lookup"><span data-stu-id="4954a-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="4954a-215">String</span><span class="sxs-lookup"><span data-stu-id="4954a-215">String</span></span> | <span data-ttu-id="4954a-216">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="4954a-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="4954a-217">1.7</span><span class="sxs-lookup"><span data-stu-id="4954a-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="4954a-218">String</span><span class="sxs-lookup"><span data-stu-id="4954a-218">String</span></span> | <span data-ttu-id="4954a-219">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="4954a-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="4954a-220">1.7</span><span class="sxs-lookup"><span data-stu-id="4954a-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4954a-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4954a-221">Requirements</span></span>

|<span data-ttu-id="4954a-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="4954a-222">Requirement</span></span>| <span data-ttu-id="4954a-223">Valor</span><span class="sxs-lookup"><span data-stu-id="4954a-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="4954a-224">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4954a-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4954a-225">1,5</span><span class="sxs-lookup"><span data-stu-id="4954a-225">1.5</span></span> |
|[<span data-ttu-id="4954a-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4954a-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4954a-227">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4954a-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="4954a-228">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4954a-228">SourceProperty: String</span></span>

<span data-ttu-id="4954a-229">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="4954a-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4954a-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="4954a-230">Type</span></span>

*   <span data-ttu-id="4954a-231">String</span><span class="sxs-lookup"><span data-stu-id="4954a-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4954a-232">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="4954a-232">Properties:</span></span>

|<span data-ttu-id="4954a-233">Nome</span><span class="sxs-lookup"><span data-stu-id="4954a-233">Name</span></span>| <span data-ttu-id="4954a-234">Tipo</span><span class="sxs-lookup"><span data-stu-id="4954a-234">Type</span></span>| <span data-ttu-id="4954a-235">Descrição</span><span class="sxs-lookup"><span data-stu-id="4954a-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4954a-236">String</span><span class="sxs-lookup"><span data-stu-id="4954a-236">String</span></span>|<span data-ttu-id="4954a-237">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4954a-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4954a-238">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="4954a-238">String</span></span>|<span data-ttu-id="4954a-239">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="4954a-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4954a-240">Requisitos</span><span class="sxs-lookup"><span data-stu-id="4954a-240">Requirements</span></span>

|<span data-ttu-id="4954a-241">Requisito</span><span class="sxs-lookup"><span data-stu-id="4954a-241">Requirement</span></span>| <span data-ttu-id="4954a-242">Valor</span><span class="sxs-lookup"><span data-stu-id="4954a-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="4954a-243">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="4954a-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4954a-244">1.1</span><span class="sxs-lookup"><span data-stu-id="4954a-244">1.1</span></span>|
|[<span data-ttu-id="4954a-245">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="4954a-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4954a-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="4954a-246">Compose or Read</span></span>|
