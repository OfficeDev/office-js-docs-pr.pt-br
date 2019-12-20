---
title: Namespace do Office – conjunto de requisitos 1,6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: e15f01db9423a9df38608f18098d2c808f5d944b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814665"
---
# <a name="office"></a><span data-ttu-id="ac57f-102">Office</span><span class="sxs-lookup"><span data-stu-id="ac57f-102">Office</span></span>

<span data-ttu-id="ac57f-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ac57f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac57f-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac57f-105">Requirements</span></span>

|<span data-ttu-id="ac57f-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac57f-106">Requirement</span></span>| <span data-ttu-id="ac57f-107">Valor</span><span class="sxs-lookup"><span data-stu-id="ac57f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac57f-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac57f-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac57f-109">1.1</span><span class="sxs-lookup"><span data-stu-id="ac57f-109">1.1</span></span>|
|[<span data-ttu-id="ac57f-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac57f-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac57f-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac57f-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ac57f-112">Propriedades</span><span class="sxs-lookup"><span data-stu-id="ac57f-112">Properties</span></span>

| <span data-ttu-id="ac57f-113">Propriedade</span><span class="sxs-lookup"><span data-stu-id="ac57f-113">Property</span></span> | <span data-ttu-id="ac57f-114">Modelos</span><span class="sxs-lookup"><span data-stu-id="ac57f-114">Modes</span></span> | <span data-ttu-id="ac57f-115">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="ac57f-115">Return type</span></span> | <span data-ttu-id="ac57f-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="ac57f-116">Minimum</span></span><br><span data-ttu-id="ac57f-117">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="ac57f-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ac57f-118">context</span><span class="sxs-lookup"><span data-stu-id="ac57f-118">context</span></span>](office.context.md) | <span data-ttu-id="ac57f-119">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac57f-119">Compose</span></span><br><span data-ttu-id="ac57f-120">Leitura</span><span class="sxs-lookup"><span data-stu-id="ac57f-120">Read</span></span> | [<span data-ttu-id="ac57f-121">Context</span><span class="sxs-lookup"><span data-stu-id="ac57f-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="ac57f-122">1.1</span><span class="sxs-lookup"><span data-stu-id="ac57f-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="ac57f-123">Enumerações</span><span class="sxs-lookup"><span data-stu-id="ac57f-123">Enumerations</span></span>

| <span data-ttu-id="ac57f-124">Enumeração</span><span class="sxs-lookup"><span data-stu-id="ac57f-124">Enumeration</span></span> | <span data-ttu-id="ac57f-125">Modelos</span><span class="sxs-lookup"><span data-stu-id="ac57f-125">Modes</span></span> | <span data-ttu-id="ac57f-126">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="ac57f-126">Return type</span></span> | <span data-ttu-id="ac57f-127">Mínimo</span><span class="sxs-lookup"><span data-stu-id="ac57f-127">Minimum</span></span><br><span data-ttu-id="ac57f-128">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="ac57f-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ac57f-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ac57f-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ac57f-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac57f-130">Compose</span></span><br><span data-ttu-id="ac57f-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="ac57f-131">Read</span></span> | <span data-ttu-id="ac57f-132">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-132">String</span></span> | [<span data-ttu-id="ac57f-133">1.1</span><span class="sxs-lookup"><span data-stu-id="ac57f-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac57f-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ac57f-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ac57f-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac57f-135">Compose</span></span><br><span data-ttu-id="ac57f-136">Leitura</span><span class="sxs-lookup"><span data-stu-id="ac57f-136">Read</span></span> | <span data-ttu-id="ac57f-137">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-137">String</span></span> | [<span data-ttu-id="ac57f-138">1.1</span><span class="sxs-lookup"><span data-stu-id="ac57f-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac57f-139">EventType</span><span class="sxs-lookup"><span data-stu-id="ac57f-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="ac57f-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac57f-140">Compose</span></span><br><span data-ttu-id="ac57f-141">Leitura</span><span class="sxs-lookup"><span data-stu-id="ac57f-141">Read</span></span> | <span data-ttu-id="ac57f-142">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-142">String</span></span> | [<span data-ttu-id="ac57f-143">1,5</span><span class="sxs-lookup"><span data-stu-id="ac57f-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="ac57f-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ac57f-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ac57f-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="ac57f-145">Compose</span></span><br><span data-ttu-id="ac57f-146">Leitura</span><span class="sxs-lookup"><span data-stu-id="ac57f-146">Read</span></span> | <span data-ttu-id="ac57f-147">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-147">String</span></span> | [<span data-ttu-id="ac57f-148">1.1</span><span class="sxs-lookup"><span data-stu-id="ac57f-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="ac57f-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="ac57f-149">Namespaces</span></span>

<span data-ttu-id="ac57f-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="ac57f-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="ac57f-151">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="ac57f-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="ac57f-152">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ac57f-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="ac57f-153">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="ac57f-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ac57f-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac57f-154">Type</span></span>

*   <span data-ttu-id="ac57f-155">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ac57f-156">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ac57f-156">Properties:</span></span>

|<span data-ttu-id="ac57f-157">Nome</span><span class="sxs-lookup"><span data-stu-id="ac57f-157">Name</span></span>| <span data-ttu-id="ac57f-158">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac57f-158">Type</span></span>| <span data-ttu-id="ac57f-159">Descrição</span><span class="sxs-lookup"><span data-stu-id="ac57f-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ac57f-160">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-160">String</span></span>|<span data-ttu-id="ac57f-161">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="ac57f-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ac57f-162">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-162">String</span></span>|<span data-ttu-id="ac57f-163">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="ac57f-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ac57f-164">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac57f-164">Requirements</span></span>

|<span data-ttu-id="ac57f-165">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac57f-165">Requirement</span></span>| <span data-ttu-id="ac57f-166">Valor</span><span class="sxs-lookup"><span data-stu-id="ac57f-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac57f-167">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac57f-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac57f-168">1.1</span><span class="sxs-lookup"><span data-stu-id="ac57f-168">1.1</span></span>|
|[<span data-ttu-id="ac57f-169">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac57f-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac57f-170">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac57f-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="ac57f-171">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ac57f-171">CoercionType: String</span></span>

<span data-ttu-id="ac57f-172">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="ac57f-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ac57f-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac57f-173">Type</span></span>

*   <span data-ttu-id="ac57f-174">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ac57f-175">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ac57f-175">Properties:</span></span>

|<span data-ttu-id="ac57f-176">Nome</span><span class="sxs-lookup"><span data-stu-id="ac57f-176">Name</span></span>| <span data-ttu-id="ac57f-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac57f-177">Type</span></span>| <span data-ttu-id="ac57f-178">Descrição</span><span class="sxs-lookup"><span data-stu-id="ac57f-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ac57f-179">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-179">String</span></span>|<span data-ttu-id="ac57f-180">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="ac57f-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ac57f-181">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-181">String</span></span>|<span data-ttu-id="ac57f-182">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="ac57f-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ac57f-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac57f-183">Requirements</span></span>

|<span data-ttu-id="ac57f-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac57f-184">Requirement</span></span>| <span data-ttu-id="ac57f-185">Valor</span><span class="sxs-lookup"><span data-stu-id="ac57f-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac57f-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac57f-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac57f-187">1.1</span><span class="sxs-lookup"><span data-stu-id="ac57f-187">1.1</span></span>|
|[<span data-ttu-id="ac57f-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac57f-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac57f-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac57f-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="ac57f-190">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ac57f-190">EventType: String</span></span>

<span data-ttu-id="ac57f-191">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="ac57f-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="ac57f-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac57f-192">Type</span></span>

*   <span data-ttu-id="ac57f-193">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ac57f-194">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ac57f-194">Properties:</span></span>

| <span data-ttu-id="ac57f-195">Nome</span><span class="sxs-lookup"><span data-stu-id="ac57f-195">Name</span></span> | <span data-ttu-id="ac57f-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac57f-196">Type</span></span> | <span data-ttu-id="ac57f-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="ac57f-197">Description</span></span> | <span data-ttu-id="ac57f-198">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="ac57f-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="ac57f-199">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-199">String</span></span> | <span data-ttu-id="ac57f-200">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="ac57f-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="ac57f-201">1,5</span><span class="sxs-lookup"><span data-stu-id="ac57f-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ac57f-202">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac57f-202">Requirements</span></span>

|<span data-ttu-id="ac57f-203">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac57f-203">Requirement</span></span>| <span data-ttu-id="ac57f-204">Valor</span><span class="sxs-lookup"><span data-stu-id="ac57f-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac57f-205">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac57f-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac57f-206">1,5</span><span class="sxs-lookup"><span data-stu-id="ac57f-206">1.5</span></span> |
|[<span data-ttu-id="ac57f-207">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac57f-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac57f-208">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac57f-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="ac57f-209">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ac57f-209">SourceProperty: String</span></span>

<span data-ttu-id="ac57f-210">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="ac57f-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ac57f-211">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac57f-211">Type</span></span>

*   <span data-ttu-id="ac57f-212">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ac57f-213">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ac57f-213">Properties:</span></span>

|<span data-ttu-id="ac57f-214">Nome</span><span class="sxs-lookup"><span data-stu-id="ac57f-214">Name</span></span>| <span data-ttu-id="ac57f-215">Tipo</span><span class="sxs-lookup"><span data-stu-id="ac57f-215">Type</span></span>| <span data-ttu-id="ac57f-216">Descrição</span><span class="sxs-lookup"><span data-stu-id="ac57f-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ac57f-217">String</span><span class="sxs-lookup"><span data-stu-id="ac57f-217">String</span></span>|<span data-ttu-id="ac57f-218">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ac57f-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ac57f-219">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ac57f-219">String</span></span>|<span data-ttu-id="ac57f-220">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ac57f-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ac57f-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ac57f-221">Requirements</span></span>

|<span data-ttu-id="ac57f-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="ac57f-222">Requirement</span></span>| <span data-ttu-id="ac57f-223">Valor</span><span class="sxs-lookup"><span data-stu-id="ac57f-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac57f-224">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ac57f-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac57f-225">1.1</span><span class="sxs-lookup"><span data-stu-id="ac57f-225">1.1</span></span>|
|[<span data-ttu-id="ac57f-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ac57f-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac57f-227">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ac57f-227">Compose or Read</span></span>|
