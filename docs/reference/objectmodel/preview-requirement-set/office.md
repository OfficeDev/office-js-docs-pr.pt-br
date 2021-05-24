---
title: Office namespace - conjunto de requisitos de visualização
description: Office namespace disponíveis para os Outlook que usam conjunto de requisitos de visualização da API de Caixa de Correio.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 72e2300dd50ff01e26417efaca92906049358fc0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590880"
---
# <a name="office-mailbox-preview-requirement-set"></a><span data-ttu-id="06fbf-103">Office (conjunto de requisitos de visualização de caixa de correio)</span><span class="sxs-lookup"><span data-stu-id="06fbf-103">Office (Mailbox preview requirement set)</span></span>

<span data-ttu-id="06fbf-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="06fbf-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="06fbf-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="06fbf-106">Requirements</span></span>

|<span data-ttu-id="06fbf-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="06fbf-107">Requirement</span></span>| <span data-ttu-id="06fbf-108">Valor</span><span class="sxs-lookup"><span data-stu-id="06fbf-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="06fbf-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="06fbf-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="06fbf-110">1.1</span><span class="sxs-lookup"><span data-stu-id="06fbf-110">1.1</span></span>|
|[<span data-ttu-id="06fbf-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="06fbf-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="06fbf-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="06fbf-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="06fbf-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="06fbf-113">Properties</span></span>

| <span data-ttu-id="06fbf-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="06fbf-114">Property</span></span> | <span data-ttu-id="06fbf-115">Modos</span><span class="sxs-lookup"><span data-stu-id="06fbf-115">Modes</span></span> | <span data-ttu-id="06fbf-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="06fbf-116">Return type</span></span> | <span data-ttu-id="06fbf-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="06fbf-117">Minimum</span></span><br><span data-ttu-id="06fbf-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="06fbf-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="06fbf-119">context</span><span class="sxs-lookup"><span data-stu-id="06fbf-119">context</span></span>](office.context.md) | <span data-ttu-id="06fbf-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="06fbf-120">Compose</span></span><br><span data-ttu-id="06fbf-121">Ler</span><span class="sxs-lookup"><span data-stu-id="06fbf-121">Read</span></span> | [<span data-ttu-id="06fbf-122">Context</span><span class="sxs-lookup"><span data-stu-id="06fbf-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="06fbf-123">1.1</span><span class="sxs-lookup"><span data-stu-id="06fbf-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="06fbf-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="06fbf-124">Enumerations</span></span>

| <span data-ttu-id="06fbf-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="06fbf-125">Enumeration</span></span> | <span data-ttu-id="06fbf-126">Modos</span><span class="sxs-lookup"><span data-stu-id="06fbf-126">Modes</span></span> | <span data-ttu-id="06fbf-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="06fbf-127">Return type</span></span> | <span data-ttu-id="06fbf-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="06fbf-128">Minimum</span></span><br><span data-ttu-id="06fbf-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="06fbf-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="06fbf-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="06fbf-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="06fbf-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="06fbf-131">Compose</span></span><br><span data-ttu-id="06fbf-132">Ler</span><span class="sxs-lookup"><span data-stu-id="06fbf-132">Read</span></span> | <span data-ttu-id="06fbf-133">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-133">String</span></span> | [<span data-ttu-id="06fbf-134">1.1</span><span class="sxs-lookup"><span data-stu-id="06fbf-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="06fbf-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="06fbf-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="06fbf-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="06fbf-136">Compose</span></span><br><span data-ttu-id="06fbf-137">Ler</span><span class="sxs-lookup"><span data-stu-id="06fbf-137">Read</span></span> | <span data-ttu-id="06fbf-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-138">String</span></span> | [<span data-ttu-id="06fbf-139">1.1</span><span class="sxs-lookup"><span data-stu-id="06fbf-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="06fbf-140">EventType</span><span class="sxs-lookup"><span data-stu-id="06fbf-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="06fbf-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="06fbf-141">Compose</span></span><br><span data-ttu-id="06fbf-142">Ler</span><span class="sxs-lookup"><span data-stu-id="06fbf-142">Read</span></span> | <span data-ttu-id="06fbf-143">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-143">String</span></span> | [<span data-ttu-id="06fbf-144">1.5</span><span class="sxs-lookup"><span data-stu-id="06fbf-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="06fbf-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="06fbf-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="06fbf-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="06fbf-146">Compose</span></span><br><span data-ttu-id="06fbf-147">Ler</span><span class="sxs-lookup"><span data-stu-id="06fbf-147">Read</span></span> | <span data-ttu-id="06fbf-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-148">String</span></span> | [<span data-ttu-id="06fbf-149">1.1</span><span class="sxs-lookup"><span data-stu-id="06fbf-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="06fbf-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="06fbf-150">Namespaces</span></span>

<span data-ttu-id="06fbf-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): inclui várias enumerações específicas Outlook, por exemplo, `ItemType` , , , , , e `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="06fbf-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="06fbf-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="06fbf-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="06fbf-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="06fbf-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="06fbf-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="06fbf-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="06fbf-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="06fbf-155">Type</span></span>

*   <span data-ttu-id="06fbf-156">String</span><span class="sxs-lookup"><span data-stu-id="06fbf-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="06fbf-157">Propriedades</span><span class="sxs-lookup"><span data-stu-id="06fbf-157">Properties</span></span>

|<span data-ttu-id="06fbf-158">Nome</span><span class="sxs-lookup"><span data-stu-id="06fbf-158">Name</span></span>| <span data-ttu-id="06fbf-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="06fbf-159">Type</span></span>| <span data-ttu-id="06fbf-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="06fbf-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="06fbf-161">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-161">String</span></span>|<span data-ttu-id="06fbf-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="06fbf-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="06fbf-163">String</span><span class="sxs-lookup"><span data-stu-id="06fbf-163">String</span></span>|<span data-ttu-id="06fbf-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="06fbf-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="06fbf-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="06fbf-165">Requirements</span></span>

|<span data-ttu-id="06fbf-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="06fbf-166">Requirement</span></span>| <span data-ttu-id="06fbf-167">Valor</span><span class="sxs-lookup"><span data-stu-id="06fbf-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="06fbf-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="06fbf-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="06fbf-169">1.1</span><span class="sxs-lookup"><span data-stu-id="06fbf-169">1.1</span></span>|
|[<span data-ttu-id="06fbf-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="06fbf-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="06fbf-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="06fbf-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="06fbf-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="06fbf-172">CoercionType: String</span></span>

<span data-ttu-id="06fbf-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="06fbf-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="06fbf-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="06fbf-174">Type</span></span>

*   <span data-ttu-id="06fbf-175">String</span><span class="sxs-lookup"><span data-stu-id="06fbf-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="06fbf-176">Propriedades</span><span class="sxs-lookup"><span data-stu-id="06fbf-176">Properties</span></span>

|<span data-ttu-id="06fbf-177">Nome</span><span class="sxs-lookup"><span data-stu-id="06fbf-177">Name</span></span>| <span data-ttu-id="06fbf-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="06fbf-178">Type</span></span>| <span data-ttu-id="06fbf-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="06fbf-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="06fbf-180">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-180">String</span></span>|<span data-ttu-id="06fbf-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="06fbf-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="06fbf-182">String</span><span class="sxs-lookup"><span data-stu-id="06fbf-182">String</span></span>|<span data-ttu-id="06fbf-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="06fbf-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="06fbf-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="06fbf-184">Requirements</span></span>

|<span data-ttu-id="06fbf-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="06fbf-185">Requirement</span></span>| <span data-ttu-id="06fbf-186">Valor</span><span class="sxs-lookup"><span data-stu-id="06fbf-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="06fbf-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="06fbf-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="06fbf-188">1.1</span><span class="sxs-lookup"><span data-stu-id="06fbf-188">1.1</span></span>|
|[<span data-ttu-id="06fbf-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="06fbf-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="06fbf-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="06fbf-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="06fbf-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="06fbf-191">EventType: String</span></span>

<span data-ttu-id="06fbf-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="06fbf-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="06fbf-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="06fbf-193">Type</span></span>

*   <span data-ttu-id="06fbf-194">String</span><span class="sxs-lookup"><span data-stu-id="06fbf-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="06fbf-195">Propriedades</span><span class="sxs-lookup"><span data-stu-id="06fbf-195">Properties</span></span>

| <span data-ttu-id="06fbf-196">Nome</span><span class="sxs-lookup"><span data-stu-id="06fbf-196">Name</span></span> | <span data-ttu-id="06fbf-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="06fbf-197">Type</span></span> | <span data-ttu-id="06fbf-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="06fbf-198">Description</span></span> | <span data-ttu-id="06fbf-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="06fbf-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="06fbf-200">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-200">String</span></span> | <span data-ttu-id="06fbf-201">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="06fbf-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="06fbf-202">1.7</span><span class="sxs-lookup"><span data-stu-id="06fbf-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="06fbf-203">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-203">String</span></span> | <span data-ttu-id="06fbf-204">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="06fbf-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="06fbf-205">1,8</span><span class="sxs-lookup"><span data-stu-id="06fbf-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="06fbf-206">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-206">String</span></span> | <span data-ttu-id="06fbf-207">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="06fbf-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="06fbf-208">1,8</span><span class="sxs-lookup"><span data-stu-id="06fbf-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="06fbf-209">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-209">String</span></span> | <span data-ttu-id="06fbf-210">Um item Outlook diferente é selecionado para exibição enquanto o painel de tarefas é fixado.</span><span class="sxs-lookup"><span data-stu-id="06fbf-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="06fbf-211">1,5</span><span class="sxs-lookup"><span data-stu-id="06fbf-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="06fbf-212">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-212">String</span></span> | <span data-ttu-id="06fbf-213">O Office tema na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="06fbf-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="06fbf-214">Visualização</span><span class="sxs-lookup"><span data-stu-id="06fbf-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="06fbf-215">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-215">String</span></span> | <span data-ttu-id="06fbf-216">A lista de destinatários do item ou local do compromisso selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="06fbf-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="06fbf-217">1.7</span><span class="sxs-lookup"><span data-stu-id="06fbf-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="06fbf-218">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-218">String</span></span> | <span data-ttu-id="06fbf-219">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="06fbf-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="06fbf-220">1.7</span><span class="sxs-lookup"><span data-stu-id="06fbf-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="06fbf-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="06fbf-221">Requirements</span></span>

|<span data-ttu-id="06fbf-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="06fbf-222">Requirement</span></span>| <span data-ttu-id="06fbf-223">Valor</span><span class="sxs-lookup"><span data-stu-id="06fbf-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="06fbf-224">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="06fbf-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="06fbf-225">1,5</span><span class="sxs-lookup"><span data-stu-id="06fbf-225">1.5</span></span> |
|[<span data-ttu-id="06fbf-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="06fbf-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="06fbf-227">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="06fbf-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="06fbf-228">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="06fbf-228">SourceProperty: String</span></span>

<span data-ttu-id="06fbf-229">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="06fbf-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="06fbf-230">Tipo</span><span class="sxs-lookup"><span data-stu-id="06fbf-230">Type</span></span>

*   <span data-ttu-id="06fbf-231">String</span><span class="sxs-lookup"><span data-stu-id="06fbf-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="06fbf-232">Propriedades</span><span class="sxs-lookup"><span data-stu-id="06fbf-232">Properties</span></span>

|<span data-ttu-id="06fbf-233">Nome</span><span class="sxs-lookup"><span data-stu-id="06fbf-233">Name</span></span>| <span data-ttu-id="06fbf-234">Tipo</span><span class="sxs-lookup"><span data-stu-id="06fbf-234">Type</span></span>| <span data-ttu-id="06fbf-235">Descrição</span><span class="sxs-lookup"><span data-stu-id="06fbf-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="06fbf-236">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="06fbf-236">String</span></span>|<span data-ttu-id="06fbf-237">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="06fbf-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="06fbf-238">String</span><span class="sxs-lookup"><span data-stu-id="06fbf-238">String</span></span>|<span data-ttu-id="06fbf-239">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="06fbf-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="06fbf-240">Requisitos</span><span class="sxs-lookup"><span data-stu-id="06fbf-240">Requirements</span></span>

|<span data-ttu-id="06fbf-241">Requisito</span><span class="sxs-lookup"><span data-stu-id="06fbf-241">Requirement</span></span>| <span data-ttu-id="06fbf-242">Valor</span><span class="sxs-lookup"><span data-stu-id="06fbf-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="06fbf-243">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="06fbf-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="06fbf-244">1.1</span><span class="sxs-lookup"><span data-stu-id="06fbf-244">1.1</span></span>|
|[<span data-ttu-id="06fbf-245">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="06fbf-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="06fbf-246">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="06fbf-246">Compose or Read</span></span>|
