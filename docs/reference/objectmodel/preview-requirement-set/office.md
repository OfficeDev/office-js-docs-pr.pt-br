---
title: Namespace do Office – conjunto de requisitos de visualização
description: Membros do namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de visualização da API da caixa de correio.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 1e0f932106df462c7cd172327082992f6e4d9a58
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431119"
---
# <a name="office-mailbox-preview-requirement-set"></a><span data-ttu-id="3c76e-103">Office (conjunto de requisitos de visualização da caixa de correio)</span><span class="sxs-lookup"><span data-stu-id="3c76e-103">Office (Mailbox preview requirement set)</span></span>

<span data-ttu-id="3c76e-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="3c76e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3c76e-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3c76e-106">Requirements</span></span>

|<span data-ttu-id="3c76e-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="3c76e-107">Requirement</span></span>| <span data-ttu-id="3c76e-108">Valor</span><span class="sxs-lookup"><span data-stu-id="3c76e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c76e-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3c76e-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c76e-110">1.1</span><span class="sxs-lookup"><span data-stu-id="3c76e-110">1.1</span></span>|
|[<span data-ttu-id="3c76e-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3c76e-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c76e-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3c76e-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="3c76e-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="3c76e-113">Properties</span></span>

| <span data-ttu-id="3c76e-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="3c76e-114">Property</span></span> | <span data-ttu-id="3c76e-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="3c76e-115">Modes</span></span> | <span data-ttu-id="3c76e-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="3c76e-116">Return type</span></span> | <span data-ttu-id="3c76e-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="3c76e-117">Minimum</span></span><br><span data-ttu-id="3c76e-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="3c76e-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3c76e-119">context</span><span class="sxs-lookup"><span data-stu-id="3c76e-119">context</span></span>](office.context.md) | <span data-ttu-id="3c76e-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="3c76e-120">Compose</span></span><br><span data-ttu-id="3c76e-121">Ler</span><span class="sxs-lookup"><span data-stu-id="3c76e-121">Read</span></span> | [<span data-ttu-id="3c76e-122">Context</span><span class="sxs-lookup"><span data-stu-id="3c76e-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="3c76e-123">1.1</span><span class="sxs-lookup"><span data-stu-id="3c76e-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="3c76e-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="3c76e-124">Enumerations</span></span>

| <span data-ttu-id="3c76e-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="3c76e-125">Enumeration</span></span> | <span data-ttu-id="3c76e-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="3c76e-126">Modes</span></span> | <span data-ttu-id="3c76e-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="3c76e-127">Return type</span></span> | <span data-ttu-id="3c76e-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="3c76e-128">Minimum</span></span><br><span data-ttu-id="3c76e-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="3c76e-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3c76e-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3c76e-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3c76e-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="3c76e-131">Compose</span></span><br><span data-ttu-id="3c76e-132">Ler</span><span class="sxs-lookup"><span data-stu-id="3c76e-132">Read</span></span> | <span data-ttu-id="3c76e-133">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-133">String</span></span> | [<span data-ttu-id="3c76e-134">1.1</span><span class="sxs-lookup"><span data-stu-id="3c76e-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3c76e-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3c76e-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3c76e-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="3c76e-136">Compose</span></span><br><span data-ttu-id="3c76e-137">Ler</span><span class="sxs-lookup"><span data-stu-id="3c76e-137">Read</span></span> | <span data-ttu-id="3c76e-138">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-138">String</span></span> | [<span data-ttu-id="3c76e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3c76e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3c76e-140">EventType</span><span class="sxs-lookup"><span data-stu-id="3c76e-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="3c76e-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="3c76e-141">Compose</span></span><br><span data-ttu-id="3c76e-142">Ler</span><span class="sxs-lookup"><span data-stu-id="3c76e-142">Read</span></span> | <span data-ttu-id="3c76e-143">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-143">String</span></span> | [<span data-ttu-id="3c76e-144">1,5</span><span class="sxs-lookup"><span data-stu-id="3c76e-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="3c76e-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3c76e-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3c76e-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="3c76e-146">Compose</span></span><br><span data-ttu-id="3c76e-147">Ler</span><span class="sxs-lookup"><span data-stu-id="3c76e-147">Read</span></span> | <span data-ttu-id="3c76e-148">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-148">String</span></span> | [<span data-ttu-id="3c76e-149">1.1</span><span class="sxs-lookup"><span data-stu-id="3c76e-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="3c76e-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="3c76e-150">Namespaces</span></span>

<span data-ttu-id="3c76e-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): inclui uma série de enumerações específicas do Outlook, por exemplo,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` e `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="3c76e-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="3c76e-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="3c76e-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="3c76e-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3c76e-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="3c76e-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="3c76e-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3c76e-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="3c76e-155">Type</span></span>

*   <span data-ttu-id="3c76e-156">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3c76e-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3c76e-157">Properties:</span></span>

|<span data-ttu-id="3c76e-158">Nome</span><span class="sxs-lookup"><span data-stu-id="3c76e-158">Name</span></span>| <span data-ttu-id="3c76e-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="3c76e-159">Type</span></span>| <span data-ttu-id="3c76e-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="3c76e-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3c76e-161">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-161">String</span></span>|<span data-ttu-id="3c76e-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="3c76e-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3c76e-163">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-163">String</span></span>|<span data-ttu-id="3c76e-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="3c76e-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3c76e-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3c76e-165">Requirements</span></span>

|<span data-ttu-id="3c76e-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="3c76e-166">Requirement</span></span>| <span data-ttu-id="3c76e-167">Valor</span><span class="sxs-lookup"><span data-stu-id="3c76e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c76e-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3c76e-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c76e-169">1.1</span><span class="sxs-lookup"><span data-stu-id="3c76e-169">1.1</span></span>|
|[<span data-ttu-id="3c76e-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3c76e-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c76e-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3c76e-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="3c76e-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3c76e-172">CoercionType: String</span></span>

<span data-ttu-id="3c76e-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="3c76e-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3c76e-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="3c76e-174">Type</span></span>

*   <span data-ttu-id="3c76e-175">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3c76e-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3c76e-176">Properties:</span></span>

|<span data-ttu-id="3c76e-177">Nome</span><span class="sxs-lookup"><span data-stu-id="3c76e-177">Name</span></span>| <span data-ttu-id="3c76e-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="3c76e-178">Type</span></span>| <span data-ttu-id="3c76e-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="3c76e-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3c76e-180">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-180">String</span></span>|<span data-ttu-id="3c76e-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="3c76e-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3c76e-182">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-182">String</span></span>|<span data-ttu-id="3c76e-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="3c76e-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3c76e-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3c76e-184">Requirements</span></span>

|<span data-ttu-id="3c76e-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="3c76e-185">Requirement</span></span>| <span data-ttu-id="3c76e-186">Valor</span><span class="sxs-lookup"><span data-stu-id="3c76e-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c76e-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3c76e-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c76e-188">1.1</span><span class="sxs-lookup"><span data-stu-id="3c76e-188">1.1</span></span>|
|[<span data-ttu-id="3c76e-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3c76e-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c76e-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3c76e-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="3c76e-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3c76e-191">EventType: String</span></span>

<span data-ttu-id="3c76e-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="3c76e-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="3c76e-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="3c76e-193">Type</span></span>

*   <span data-ttu-id="3c76e-194">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3c76e-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3c76e-195">Properties:</span></span>

| <span data-ttu-id="3c76e-196">Nome</span><span class="sxs-lookup"><span data-stu-id="3c76e-196">Name</span></span> | <span data-ttu-id="3c76e-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="3c76e-197">Type</span></span> | <span data-ttu-id="3c76e-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="3c76e-198">Description</span></span> | <span data-ttu-id="3c76e-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="3c76e-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="3c76e-200">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-200">String</span></span> | <span data-ttu-id="3c76e-201">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="3c76e-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="3c76e-202">1.7</span><span class="sxs-lookup"><span data-stu-id="3c76e-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="3c76e-203">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-203">String</span></span> | <span data-ttu-id="3c76e-204">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="3c76e-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="3c76e-205">1,8</span><span class="sxs-lookup"><span data-stu-id="3c76e-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="3c76e-206">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-206">String</span></span> | <span data-ttu-id="3c76e-207">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="3c76e-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="3c76e-208">1,8</span><span class="sxs-lookup"><span data-stu-id="3c76e-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="3c76e-209">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-209">String</span></span> | <span data-ttu-id="3c76e-210">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="3c76e-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="3c76e-211">1,5</span><span class="sxs-lookup"><span data-stu-id="3c76e-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="3c76e-212">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-212">String</span></span> | <span data-ttu-id="3c76e-213">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="3c76e-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="3c76e-214">Visualização</span><span class="sxs-lookup"><span data-stu-id="3c76e-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="3c76e-215">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-215">String</span></span> | <span data-ttu-id="3c76e-216">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="3c76e-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="3c76e-217">1.7</span><span class="sxs-lookup"><span data-stu-id="3c76e-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="3c76e-218">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-218">String</span></span> | <span data-ttu-id="3c76e-219">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="3c76e-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="3c76e-220">1.7</span><span class="sxs-lookup"><span data-stu-id="3c76e-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3c76e-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3c76e-221">Requirements</span></span>

|<span data-ttu-id="3c76e-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="3c76e-222">Requirement</span></span>| <span data-ttu-id="3c76e-223">Valor</span><span class="sxs-lookup"><span data-stu-id="3c76e-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c76e-224">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3c76e-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c76e-225">1,5</span><span class="sxs-lookup"><span data-stu-id="3c76e-225">1.5</span></span> |
|[<span data-ttu-id="3c76e-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3c76e-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c76e-227">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3c76e-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="3c76e-228">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3c76e-228">SourceProperty: String</span></span>

<span data-ttu-id="3c76e-229">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="3c76e-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3c76e-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="3c76e-230">Type</span></span>

*   <span data-ttu-id="3c76e-231">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3c76e-232">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3c76e-232">Properties:</span></span>

|<span data-ttu-id="3c76e-233">Nome</span><span class="sxs-lookup"><span data-stu-id="3c76e-233">Name</span></span>| <span data-ttu-id="3c76e-234">Tipo</span><span class="sxs-lookup"><span data-stu-id="3c76e-234">Type</span></span>| <span data-ttu-id="3c76e-235">Descrição</span><span class="sxs-lookup"><span data-stu-id="3c76e-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3c76e-236">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-236">String</span></span>|<span data-ttu-id="3c76e-237">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="3c76e-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3c76e-238">String</span><span class="sxs-lookup"><span data-stu-id="3c76e-238">String</span></span>|<span data-ttu-id="3c76e-239">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="3c76e-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3c76e-240">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3c76e-240">Requirements</span></span>

|<span data-ttu-id="3c76e-241">Requisito</span><span class="sxs-lookup"><span data-stu-id="3c76e-241">Requirement</span></span>| <span data-ttu-id="3c76e-242">Valor</span><span class="sxs-lookup"><span data-stu-id="3c76e-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c76e-243">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3c76e-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c76e-244">1.1</span><span class="sxs-lookup"><span data-stu-id="3c76e-244">1.1</span></span>|
|[<span data-ttu-id="3c76e-245">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3c76e-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c76e-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3c76e-246">Compose or Read</span></span>|
