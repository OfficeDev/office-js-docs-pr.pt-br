---
title: Office namespace - conjunto de requisitos 1.10
description: Office namespace disponíveis para os Outlook que usam o conjunto de requisitos da API de Caixa de Correio 1.10.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e7b7ab9127ebf8ce9b7394d348144fe63b47de6c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592023"
---
# <a name="office-mailbox-requirement-set-110"></a><span data-ttu-id="c694f-103">Office (conjunto de requisitos de caixa de correio 1.10)</span><span class="sxs-lookup"><span data-stu-id="c694f-103">Office (Mailbox requirement set 1.10)</span></span>

<span data-ttu-id="c694f-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c694f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c694f-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c694f-106">Requirements</span></span>

|<span data-ttu-id="c694f-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="c694f-107">Requirement</span></span>| <span data-ttu-id="c694f-108">Valor</span><span class="sxs-lookup"><span data-stu-id="c694f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c694f-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c694f-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c694f-110">1.1</span><span class="sxs-lookup"><span data-stu-id="c694f-110">1.1</span></span>|
|[<span data-ttu-id="c694f-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c694f-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c694f-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c694f-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="c694f-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="c694f-113">Properties</span></span>

| <span data-ttu-id="c694f-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="c694f-114">Property</span></span> | <span data-ttu-id="c694f-115">Modos</span><span class="sxs-lookup"><span data-stu-id="c694f-115">Modes</span></span> | <span data-ttu-id="c694f-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="c694f-116">Return type</span></span> | <span data-ttu-id="c694f-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="c694f-117">Minimum</span></span><br><span data-ttu-id="c694f-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="c694f-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c694f-119">context</span><span class="sxs-lookup"><span data-stu-id="c694f-119">context</span></span>](office.context.md) | <span data-ttu-id="c694f-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="c694f-120">Compose</span></span><br><span data-ttu-id="c694f-121">Ler</span><span class="sxs-lookup"><span data-stu-id="c694f-121">Read</span></span> | [<span data-ttu-id="c694f-122">Context</span><span class="sxs-lookup"><span data-stu-id="c694f-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="c694f-123">1.1</span><span class="sxs-lookup"><span data-stu-id="c694f-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="c694f-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="c694f-124">Enumerations</span></span>

| <span data-ttu-id="c694f-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="c694f-125">Enumeration</span></span> | <span data-ttu-id="c694f-126">Modos</span><span class="sxs-lookup"><span data-stu-id="c694f-126">Modes</span></span> | <span data-ttu-id="c694f-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="c694f-127">Return type</span></span> | <span data-ttu-id="c694f-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="c694f-128">Minimum</span></span><br><span data-ttu-id="c694f-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="c694f-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c694f-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c694f-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c694f-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="c694f-131">Compose</span></span><br><span data-ttu-id="c694f-132">Ler</span><span class="sxs-lookup"><span data-stu-id="c694f-132">Read</span></span> | <span data-ttu-id="c694f-133">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-133">String</span></span> | [<span data-ttu-id="c694f-134">1.1</span><span class="sxs-lookup"><span data-stu-id="c694f-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c694f-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c694f-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c694f-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="c694f-136">Compose</span></span><br><span data-ttu-id="c694f-137">Ler</span><span class="sxs-lookup"><span data-stu-id="c694f-137">Read</span></span> | <span data-ttu-id="c694f-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-138">String</span></span> | [<span data-ttu-id="c694f-139">1.1</span><span class="sxs-lookup"><span data-stu-id="c694f-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c694f-140">EventType</span><span class="sxs-lookup"><span data-stu-id="c694f-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="c694f-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="c694f-141">Compose</span></span><br><span data-ttu-id="c694f-142">Ler</span><span class="sxs-lookup"><span data-stu-id="c694f-142">Read</span></span> | <span data-ttu-id="c694f-143">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-143">String</span></span> | [<span data-ttu-id="c694f-144">1.5</span><span class="sxs-lookup"><span data-stu-id="c694f-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="c694f-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c694f-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c694f-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="c694f-146">Compose</span></span><br><span data-ttu-id="c694f-147">Ler</span><span class="sxs-lookup"><span data-stu-id="c694f-147">Read</span></span> | <span data-ttu-id="c694f-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-148">String</span></span> | [<span data-ttu-id="c694f-149">1.1</span><span class="sxs-lookup"><span data-stu-id="c694f-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="c694f-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="c694f-150">Namespaces</span></span>

<span data-ttu-id="c694f-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.10&preserve-view=true): inclui várias enumerações específicas Outlook, por exemplo, `ItemType` , , , , , e `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="c694f-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.10&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="c694f-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="c694f-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="c694f-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="c694f-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="c694f-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="c694f-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c694f-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="c694f-155">Type</span></span>

*   <span data-ttu-id="c694f-156">String</span><span class="sxs-lookup"><span data-stu-id="c694f-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c694f-157">Propriedades</span><span class="sxs-lookup"><span data-stu-id="c694f-157">Properties</span></span>

|<span data-ttu-id="c694f-158">Nome</span><span class="sxs-lookup"><span data-stu-id="c694f-158">Name</span></span>| <span data-ttu-id="c694f-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="c694f-159">Type</span></span>| <span data-ttu-id="c694f-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="c694f-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c694f-161">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-161">String</span></span>|<span data-ttu-id="c694f-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="c694f-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c694f-163">String</span><span class="sxs-lookup"><span data-stu-id="c694f-163">String</span></span>|<span data-ttu-id="c694f-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="c694f-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c694f-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c694f-165">Requirements</span></span>

|<span data-ttu-id="c694f-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="c694f-166">Requirement</span></span>| <span data-ttu-id="c694f-167">Valor</span><span class="sxs-lookup"><span data-stu-id="c694f-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="c694f-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c694f-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c694f-169">1.1</span><span class="sxs-lookup"><span data-stu-id="c694f-169">1.1</span></span>|
|[<span data-ttu-id="c694f-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c694f-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c694f-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c694f-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="c694f-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="c694f-172">CoercionType: String</span></span>

<span data-ttu-id="c694f-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="c694f-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c694f-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="c694f-174">Type</span></span>

*   <span data-ttu-id="c694f-175">String</span><span class="sxs-lookup"><span data-stu-id="c694f-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c694f-176">Propriedades</span><span class="sxs-lookup"><span data-stu-id="c694f-176">Properties</span></span>

|<span data-ttu-id="c694f-177">Nome</span><span class="sxs-lookup"><span data-stu-id="c694f-177">Name</span></span>| <span data-ttu-id="c694f-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="c694f-178">Type</span></span>| <span data-ttu-id="c694f-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="c694f-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c694f-180">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-180">String</span></span>|<span data-ttu-id="c694f-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="c694f-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c694f-182">String</span><span class="sxs-lookup"><span data-stu-id="c694f-182">String</span></span>|<span data-ttu-id="c694f-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="c694f-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c694f-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c694f-184">Requirements</span></span>

|<span data-ttu-id="c694f-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="c694f-185">Requirement</span></span>| <span data-ttu-id="c694f-186">Valor</span><span class="sxs-lookup"><span data-stu-id="c694f-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="c694f-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c694f-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c694f-188">1.1</span><span class="sxs-lookup"><span data-stu-id="c694f-188">1.1</span></span>|
|[<span data-ttu-id="c694f-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c694f-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c694f-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c694f-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="c694f-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="c694f-191">EventType: String</span></span>

<span data-ttu-id="c694f-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="c694f-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="c694f-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="c694f-193">Type</span></span>

*   <span data-ttu-id="c694f-194">String</span><span class="sxs-lookup"><span data-stu-id="c694f-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c694f-195">Propriedades</span><span class="sxs-lookup"><span data-stu-id="c694f-195">Properties</span></span>

| <span data-ttu-id="c694f-196">Nome</span><span class="sxs-lookup"><span data-stu-id="c694f-196">Name</span></span> | <span data-ttu-id="c694f-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="c694f-197">Type</span></span> | <span data-ttu-id="c694f-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="c694f-198">Description</span></span> | <span data-ttu-id="c694f-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="c694f-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="c694f-200">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-200">String</span></span> | <span data-ttu-id="c694f-201">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="c694f-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="c694f-202">1.7</span><span class="sxs-lookup"><span data-stu-id="c694f-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="c694f-203">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-203">String</span></span> | <span data-ttu-id="c694f-204">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="c694f-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="c694f-205">1,8</span><span class="sxs-lookup"><span data-stu-id="c694f-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="c694f-206">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-206">String</span></span> | <span data-ttu-id="c694f-207">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="c694f-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="c694f-208">1,8</span><span class="sxs-lookup"><span data-stu-id="c694f-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="c694f-209">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-209">String</span></span> | <span data-ttu-id="c694f-210">Um item Outlook diferente é selecionado para exibição enquanto o painel de tarefas é fixado.</span><span class="sxs-lookup"><span data-stu-id="c694f-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="c694f-211">1,5</span><span class="sxs-lookup"><span data-stu-id="c694f-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="c694f-212">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-212">String</span></span> | <span data-ttu-id="c694f-213">O Office tema na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="c694f-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="c694f-214">1.10</span><span class="sxs-lookup"><span data-stu-id="c694f-214">1.10</span></span> |
|`RecipientsChanged`| <span data-ttu-id="c694f-215">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-215">String</span></span> | <span data-ttu-id="c694f-216">A lista de destinatários do item ou local do compromisso selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="c694f-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="c694f-217">1.7</span><span class="sxs-lookup"><span data-stu-id="c694f-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="c694f-218">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-218">String</span></span> | <span data-ttu-id="c694f-219">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="c694f-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="c694f-220">1.7</span><span class="sxs-lookup"><span data-stu-id="c694f-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c694f-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c694f-221">Requirements</span></span>

|<span data-ttu-id="c694f-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="c694f-222">Requirement</span></span>| <span data-ttu-id="c694f-223">Valor</span><span class="sxs-lookup"><span data-stu-id="c694f-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="c694f-224">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c694f-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c694f-225">1,5</span><span class="sxs-lookup"><span data-stu-id="c694f-225">1.5</span></span> |
|[<span data-ttu-id="c694f-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c694f-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c694f-227">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c694f-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="c694f-228">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="c694f-228">SourceProperty: String</span></span>

<span data-ttu-id="c694f-229">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="c694f-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c694f-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="c694f-230">Type</span></span>

*   <span data-ttu-id="c694f-231">String</span><span class="sxs-lookup"><span data-stu-id="c694f-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c694f-232">Propriedades</span><span class="sxs-lookup"><span data-stu-id="c694f-232">Properties</span></span>

|<span data-ttu-id="c694f-233">Nome</span><span class="sxs-lookup"><span data-stu-id="c694f-233">Name</span></span>| <span data-ttu-id="c694f-234">Tipo</span><span class="sxs-lookup"><span data-stu-id="c694f-234">Type</span></span>| <span data-ttu-id="c694f-235">Descrição</span><span class="sxs-lookup"><span data-stu-id="c694f-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c694f-236">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="c694f-236">String</span></span>|<span data-ttu-id="c694f-237">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="c694f-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c694f-238">String</span><span class="sxs-lookup"><span data-stu-id="c694f-238">String</span></span>|<span data-ttu-id="c694f-239">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="c694f-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c694f-240">Requisitos</span><span class="sxs-lookup"><span data-stu-id="c694f-240">Requirements</span></span>

|<span data-ttu-id="c694f-241">Requisito</span><span class="sxs-lookup"><span data-stu-id="c694f-241">Requirement</span></span>| <span data-ttu-id="c694f-242">Valor</span><span class="sxs-lookup"><span data-stu-id="c694f-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="c694f-243">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="c694f-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c694f-244">1.1</span><span class="sxs-lookup"><span data-stu-id="c694f-244">1.1</span></span>|
|[<span data-ttu-id="c694f-245">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="c694f-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c694f-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="c694f-246">Compose or Read</span></span>|
