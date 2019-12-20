---
title: Namespace do Office – conjunto de requisitos 1,7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9bfff9c45cb157d2dcd42997a01f5ada40aecfa0
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814567"
---
# <a name="office"></a><span data-ttu-id="fd4ca-102">Office</span><span class="sxs-lookup"><span data-stu-id="fd4ca-102">Office</span></span>

<span data-ttu-id="fd4ca-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="fd4ca-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd4ca-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fd4ca-105">Requirements</span></span>

|<span data-ttu-id="fd4ca-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="fd4ca-106">Requirement</span></span>| <span data-ttu-id="fd4ca-107">Valor</span><span class="sxs-lookup"><span data-stu-id="fd4ca-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd4ca-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fd4ca-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fd4ca-109">1.1</span><span class="sxs-lookup"><span data-stu-id="fd4ca-109">1.1</span></span>|
|[<span data-ttu-id="fd4ca-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fd4ca-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd4ca-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fd4ca-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="fd4ca-112">Propriedades</span><span class="sxs-lookup"><span data-stu-id="fd4ca-112">Properties</span></span>

| <span data-ttu-id="fd4ca-113">Propriedade</span><span class="sxs-lookup"><span data-stu-id="fd4ca-113">Property</span></span> | <span data-ttu-id="fd4ca-114">Modelos</span><span class="sxs-lookup"><span data-stu-id="fd4ca-114">Modes</span></span> | <span data-ttu-id="fd4ca-115">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="fd4ca-115">Return type</span></span> | <span data-ttu-id="fd4ca-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="fd4ca-116">Minimum</span></span><br><span data-ttu-id="fd4ca-117">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="fd4ca-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="fd4ca-118">context</span><span class="sxs-lookup"><span data-stu-id="fd4ca-118">context</span></span>](office.context.md) | <span data-ttu-id="fd4ca-119">Escrever</span><span class="sxs-lookup"><span data-stu-id="fd4ca-119">Compose</span></span><br><span data-ttu-id="fd4ca-120">Leitura</span><span class="sxs-lookup"><span data-stu-id="fd4ca-120">Read</span></span> | [<span data-ttu-id="fd4ca-121">Context</span><span class="sxs-lookup"><span data-stu-id="fd4ca-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="fd4ca-122">1.1</span><span class="sxs-lookup"><span data-stu-id="fd4ca-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="fd4ca-123">Enumerações</span><span class="sxs-lookup"><span data-stu-id="fd4ca-123">Enumerations</span></span>

| <span data-ttu-id="fd4ca-124">Enumeração</span><span class="sxs-lookup"><span data-stu-id="fd4ca-124">Enumeration</span></span> | <span data-ttu-id="fd4ca-125">Modelos</span><span class="sxs-lookup"><span data-stu-id="fd4ca-125">Modes</span></span> | <span data-ttu-id="fd4ca-126">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="fd4ca-126">Return type</span></span> | <span data-ttu-id="fd4ca-127">Mínimo</span><span class="sxs-lookup"><span data-stu-id="fd4ca-127">Minimum</span></span><br><span data-ttu-id="fd4ca-128">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="fd4ca-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="fd4ca-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="fd4ca-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="fd4ca-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="fd4ca-130">Compose</span></span><br><span data-ttu-id="fd4ca-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="fd4ca-131">Read</span></span> | <span data-ttu-id="fd4ca-132">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-132">String</span></span> | [<span data-ttu-id="fd4ca-133">1.1</span><span class="sxs-lookup"><span data-stu-id="fd4ca-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="fd4ca-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="fd4ca-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="fd4ca-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="fd4ca-135">Compose</span></span><br><span data-ttu-id="fd4ca-136">Leitura</span><span class="sxs-lookup"><span data-stu-id="fd4ca-136">Read</span></span> | <span data-ttu-id="fd4ca-137">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-137">String</span></span> | [<span data-ttu-id="fd4ca-138">1.1</span><span class="sxs-lookup"><span data-stu-id="fd4ca-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="fd4ca-139">EventType</span><span class="sxs-lookup"><span data-stu-id="fd4ca-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="fd4ca-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="fd4ca-140">Compose</span></span><br><span data-ttu-id="fd4ca-141">Leitura</span><span class="sxs-lookup"><span data-stu-id="fd4ca-141">Read</span></span> | <span data-ttu-id="fd4ca-142">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-142">String</span></span> | [<span data-ttu-id="fd4ca-143">1,5</span><span class="sxs-lookup"><span data-stu-id="fd4ca-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="fd4ca-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="fd4ca-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="fd4ca-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="fd4ca-145">Compose</span></span><br><span data-ttu-id="fd4ca-146">Leitura</span><span class="sxs-lookup"><span data-stu-id="fd4ca-146">Read</span></span> | <span data-ttu-id="fd4ca-147">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-147">String</span></span> | [<span data-ttu-id="fd4ca-148">1.1</span><span class="sxs-lookup"><span data-stu-id="fd4ca-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="fd4ca-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="fd4ca-149">Namespaces</span></span>

<span data-ttu-id="fd4ca-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="fd4ca-151">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="fd4ca-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="fd4ca-152">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fd4ca-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="fd4ca-153">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="fd4ca-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="fd4ca-154">Type</span></span>

*   <span data-ttu-id="fd4ca-155">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fd4ca-156">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="fd4ca-156">Properties:</span></span>

|<span data-ttu-id="fd4ca-157">Nome</span><span class="sxs-lookup"><span data-stu-id="fd4ca-157">Name</span></span>| <span data-ttu-id="fd4ca-158">Tipo</span><span class="sxs-lookup"><span data-stu-id="fd4ca-158">Type</span></span>| <span data-ttu-id="fd4ca-159">Descrição</span><span class="sxs-lookup"><span data-stu-id="fd4ca-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="fd4ca-160">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-160">String</span></span>|<span data-ttu-id="fd4ca-161">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="fd4ca-162">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-162">String</span></span>|<span data-ttu-id="fd4ca-163">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd4ca-164">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fd4ca-164">Requirements</span></span>

|<span data-ttu-id="fd4ca-165">Requisito</span><span class="sxs-lookup"><span data-stu-id="fd4ca-165">Requirement</span></span>| <span data-ttu-id="fd4ca-166">Valor</span><span class="sxs-lookup"><span data-stu-id="fd4ca-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd4ca-167">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fd4ca-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fd4ca-168">1.1</span><span class="sxs-lookup"><span data-stu-id="fd4ca-168">1.1</span></span>|
|[<span data-ttu-id="fd4ca-169">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fd4ca-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd4ca-170">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fd4ca-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="fd4ca-171">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fd4ca-171">CoercionType: String</span></span>

<span data-ttu-id="fd4ca-172">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="fd4ca-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="fd4ca-173">Type</span></span>

*   <span data-ttu-id="fd4ca-174">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fd4ca-175">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="fd4ca-175">Properties:</span></span>

|<span data-ttu-id="fd4ca-176">Nome</span><span class="sxs-lookup"><span data-stu-id="fd4ca-176">Name</span></span>| <span data-ttu-id="fd4ca-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="fd4ca-177">Type</span></span>| <span data-ttu-id="fd4ca-178">Descrição</span><span class="sxs-lookup"><span data-stu-id="fd4ca-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="fd4ca-179">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-179">String</span></span>|<span data-ttu-id="fd4ca-180">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="fd4ca-181">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-181">String</span></span>|<span data-ttu-id="fd4ca-182">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd4ca-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fd4ca-183">Requirements</span></span>

|<span data-ttu-id="fd4ca-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="fd4ca-184">Requirement</span></span>| <span data-ttu-id="fd4ca-185">Valor</span><span class="sxs-lookup"><span data-stu-id="fd4ca-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd4ca-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fd4ca-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fd4ca-187">1.1</span><span class="sxs-lookup"><span data-stu-id="fd4ca-187">1.1</span></span>|
|[<span data-ttu-id="fd4ca-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fd4ca-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd4ca-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fd4ca-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="fd4ca-190">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fd4ca-190">EventType: String</span></span>

<span data-ttu-id="fd4ca-191">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="fd4ca-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="fd4ca-192">Type</span></span>

*   <span data-ttu-id="fd4ca-193">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fd4ca-194">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="fd4ca-194">Properties:</span></span>

| <span data-ttu-id="fd4ca-195">Nome</span><span class="sxs-lookup"><span data-stu-id="fd4ca-195">Name</span></span> | <span data-ttu-id="fd4ca-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="fd4ca-196">Type</span></span> | <span data-ttu-id="fd4ca-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="fd4ca-197">Description</span></span> | <span data-ttu-id="fd4ca-198">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="fd4ca-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="fd4ca-199">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-199">String</span></span> | <span data-ttu-id="fd4ca-200">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="fd4ca-201">1.7</span><span class="sxs-lookup"><span data-stu-id="fd4ca-201">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="fd4ca-202">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-202">String</span></span> | <span data-ttu-id="fd4ca-203">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-203">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="fd4ca-204">1,5</span><span class="sxs-lookup"><span data-stu-id="fd4ca-204">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="fd4ca-205">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-205">String</span></span> | <span data-ttu-id="fd4ca-206">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-206">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="fd4ca-207">1.7</span><span class="sxs-lookup"><span data-stu-id="fd4ca-207">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="fd4ca-208">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-208">String</span></span> | <span data-ttu-id="fd4ca-209">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-209">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="fd4ca-210">1.7</span><span class="sxs-lookup"><span data-stu-id="fd4ca-210">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fd4ca-211">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fd4ca-211">Requirements</span></span>

|<span data-ttu-id="fd4ca-212">Requisito</span><span class="sxs-lookup"><span data-stu-id="fd4ca-212">Requirement</span></span>| <span data-ttu-id="fd4ca-213">Valor</span><span class="sxs-lookup"><span data-stu-id="fd4ca-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd4ca-214">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fd4ca-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fd4ca-215">1,5</span><span class="sxs-lookup"><span data-stu-id="fd4ca-215">1.5</span></span> |
|[<span data-ttu-id="fd4ca-216">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fd4ca-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd4ca-217">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fd4ca-217">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="fd4ca-218">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fd4ca-218">SourceProperty: String</span></span>

<span data-ttu-id="fd4ca-219">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-219">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="fd4ca-220">Tipo</span><span class="sxs-lookup"><span data-stu-id="fd4ca-220">Type</span></span>

*   <span data-ttu-id="fd4ca-221">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-221">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fd4ca-222">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="fd4ca-222">Properties:</span></span>

|<span data-ttu-id="fd4ca-223">Nome</span><span class="sxs-lookup"><span data-stu-id="fd4ca-223">Name</span></span>| <span data-ttu-id="fd4ca-224">Tipo</span><span class="sxs-lookup"><span data-stu-id="fd4ca-224">Type</span></span>| <span data-ttu-id="fd4ca-225">Descrição</span><span class="sxs-lookup"><span data-stu-id="fd4ca-225">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="fd4ca-226">String</span><span class="sxs-lookup"><span data-stu-id="fd4ca-226">String</span></span>|<span data-ttu-id="fd4ca-227">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-227">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="fd4ca-228">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fd4ca-228">String</span></span>|<span data-ttu-id="fd4ca-229">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fd4ca-229">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd4ca-230">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fd4ca-230">Requirements</span></span>

|<span data-ttu-id="fd4ca-231">Requisito</span><span class="sxs-lookup"><span data-stu-id="fd4ca-231">Requirement</span></span>| <span data-ttu-id="fd4ca-232">Valor</span><span class="sxs-lookup"><span data-stu-id="fd4ca-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd4ca-233">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fd4ca-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="fd4ca-234">1.1</span><span class="sxs-lookup"><span data-stu-id="fd4ca-234">1.1</span></span>|
|[<span data-ttu-id="fd4ca-235">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fd4ca-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd4ca-236">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="fd4ca-236">Compose or Read</span></span>|
