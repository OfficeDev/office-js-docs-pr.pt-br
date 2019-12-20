---
title: Namespace do Office – conjunto de requisitos 1,5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 63dbb3ac10492ac6e2019353b8cb057227e4c1e6
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814749"
---
# <a name="office"></a><span data-ttu-id="2459b-102">Office</span><span class="sxs-lookup"><span data-stu-id="2459b-102">Office</span></span>

<span data-ttu-id="2459b-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="2459b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2459b-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2459b-105">Requirements</span></span>

|<span data-ttu-id="2459b-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="2459b-106">Requirement</span></span>| <span data-ttu-id="2459b-107">Valor</span><span class="sxs-lookup"><span data-stu-id="2459b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2459b-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2459b-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2459b-109">1.1</span><span class="sxs-lookup"><span data-stu-id="2459b-109">1.1</span></span>|
|[<span data-ttu-id="2459b-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2459b-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2459b-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2459b-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2459b-112">Propriedades</span><span class="sxs-lookup"><span data-stu-id="2459b-112">Properties</span></span>

| <span data-ttu-id="2459b-113">Propriedade</span><span class="sxs-lookup"><span data-stu-id="2459b-113">Property</span></span> | <span data-ttu-id="2459b-114">Modelos</span><span class="sxs-lookup"><span data-stu-id="2459b-114">Modes</span></span> | <span data-ttu-id="2459b-115">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="2459b-115">Return type</span></span> | <span data-ttu-id="2459b-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="2459b-116">Minimum</span></span><br><span data-ttu-id="2459b-117">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="2459b-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2459b-118">context</span><span class="sxs-lookup"><span data-stu-id="2459b-118">context</span></span>](office.context.md) | <span data-ttu-id="2459b-119">Escrever</span><span class="sxs-lookup"><span data-stu-id="2459b-119">Compose</span></span><br><span data-ttu-id="2459b-120">Leitura</span><span class="sxs-lookup"><span data-stu-id="2459b-120">Read</span></span> | [<span data-ttu-id="2459b-121">Context</span><span class="sxs-lookup"><span data-stu-id="2459b-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="2459b-122">1.1</span><span class="sxs-lookup"><span data-stu-id="2459b-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="2459b-123">Enumerações</span><span class="sxs-lookup"><span data-stu-id="2459b-123">Enumerations</span></span>

| <span data-ttu-id="2459b-124">Enumeração</span><span class="sxs-lookup"><span data-stu-id="2459b-124">Enumeration</span></span> | <span data-ttu-id="2459b-125">Modelos</span><span class="sxs-lookup"><span data-stu-id="2459b-125">Modes</span></span> | <span data-ttu-id="2459b-126">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="2459b-126">Return type</span></span> | <span data-ttu-id="2459b-127">Mínimo</span><span class="sxs-lookup"><span data-stu-id="2459b-127">Minimum</span></span><br><span data-ttu-id="2459b-128">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="2459b-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2459b-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2459b-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2459b-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="2459b-130">Compose</span></span><br><span data-ttu-id="2459b-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="2459b-131">Read</span></span> | <span data-ttu-id="2459b-132">String</span><span class="sxs-lookup"><span data-stu-id="2459b-132">String</span></span> | [<span data-ttu-id="2459b-133">1.1</span><span class="sxs-lookup"><span data-stu-id="2459b-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2459b-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2459b-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2459b-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="2459b-135">Compose</span></span><br><span data-ttu-id="2459b-136">Leitura</span><span class="sxs-lookup"><span data-stu-id="2459b-136">Read</span></span> | <span data-ttu-id="2459b-137">String</span><span class="sxs-lookup"><span data-stu-id="2459b-137">String</span></span> | [<span data-ttu-id="2459b-138">1.1</span><span class="sxs-lookup"><span data-stu-id="2459b-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2459b-139">EventType</span><span class="sxs-lookup"><span data-stu-id="2459b-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2459b-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="2459b-140">Compose</span></span><br><span data-ttu-id="2459b-141">Leitura</span><span class="sxs-lookup"><span data-stu-id="2459b-141">Read</span></span> | <span data-ttu-id="2459b-142">String</span><span class="sxs-lookup"><span data-stu-id="2459b-142">String</span></span> | [<span data-ttu-id="2459b-143">1,5</span><span class="sxs-lookup"><span data-stu-id="2459b-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="2459b-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2459b-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2459b-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="2459b-145">Compose</span></span><br><span data-ttu-id="2459b-146">Leitura</span><span class="sxs-lookup"><span data-stu-id="2459b-146">Read</span></span> | <span data-ttu-id="2459b-147">String</span><span class="sxs-lookup"><span data-stu-id="2459b-147">String</span></span> | [<span data-ttu-id="2459b-148">1.1</span><span class="sxs-lookup"><span data-stu-id="2459b-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="2459b-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="2459b-149">Namespaces</span></span>

<span data-ttu-id="2459b-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="2459b-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="2459b-151">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="2459b-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="2459b-152">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2459b-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="2459b-153">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="2459b-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2459b-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="2459b-154">Type</span></span>

*   <span data-ttu-id="2459b-155">String</span><span class="sxs-lookup"><span data-stu-id="2459b-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2459b-156">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="2459b-156">Properties:</span></span>

|<span data-ttu-id="2459b-157">Nome</span><span class="sxs-lookup"><span data-stu-id="2459b-157">Name</span></span>| <span data-ttu-id="2459b-158">Tipo</span><span class="sxs-lookup"><span data-stu-id="2459b-158">Type</span></span>| <span data-ttu-id="2459b-159">Descrição</span><span class="sxs-lookup"><span data-stu-id="2459b-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2459b-160">String</span><span class="sxs-lookup"><span data-stu-id="2459b-160">String</span></span>|<span data-ttu-id="2459b-161">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="2459b-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2459b-162">String</span><span class="sxs-lookup"><span data-stu-id="2459b-162">String</span></span>|<span data-ttu-id="2459b-163">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="2459b-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2459b-164">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2459b-164">Requirements</span></span>

|<span data-ttu-id="2459b-165">Requisito</span><span class="sxs-lookup"><span data-stu-id="2459b-165">Requirement</span></span>| <span data-ttu-id="2459b-166">Valor</span><span class="sxs-lookup"><span data-stu-id="2459b-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="2459b-167">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2459b-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2459b-168">1.1</span><span class="sxs-lookup"><span data-stu-id="2459b-168">1.1</span></span>|
|[<span data-ttu-id="2459b-169">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2459b-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2459b-170">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2459b-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="2459b-171">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2459b-171">CoercionType: String</span></span>

<span data-ttu-id="2459b-172">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="2459b-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2459b-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="2459b-173">Type</span></span>

*   <span data-ttu-id="2459b-174">String</span><span class="sxs-lookup"><span data-stu-id="2459b-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2459b-175">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="2459b-175">Properties:</span></span>

|<span data-ttu-id="2459b-176">Nome</span><span class="sxs-lookup"><span data-stu-id="2459b-176">Name</span></span>| <span data-ttu-id="2459b-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="2459b-177">Type</span></span>| <span data-ttu-id="2459b-178">Descrição</span><span class="sxs-lookup"><span data-stu-id="2459b-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2459b-179">String</span><span class="sxs-lookup"><span data-stu-id="2459b-179">String</span></span>|<span data-ttu-id="2459b-180">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="2459b-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2459b-181">String</span><span class="sxs-lookup"><span data-stu-id="2459b-181">String</span></span>|<span data-ttu-id="2459b-182">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="2459b-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2459b-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2459b-183">Requirements</span></span>

|<span data-ttu-id="2459b-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="2459b-184">Requirement</span></span>| <span data-ttu-id="2459b-185">Valor</span><span class="sxs-lookup"><span data-stu-id="2459b-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="2459b-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2459b-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2459b-187">1.1</span><span class="sxs-lookup"><span data-stu-id="2459b-187">1.1</span></span>|
|[<span data-ttu-id="2459b-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2459b-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2459b-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2459b-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="2459b-190">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2459b-190">EventType: String</span></span>

<span data-ttu-id="2459b-191">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="2459b-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2459b-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="2459b-192">Type</span></span>

*   <span data-ttu-id="2459b-193">String</span><span class="sxs-lookup"><span data-stu-id="2459b-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2459b-194">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="2459b-194">Properties:</span></span>

| <span data-ttu-id="2459b-195">Nome</span><span class="sxs-lookup"><span data-stu-id="2459b-195">Name</span></span> | <span data-ttu-id="2459b-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="2459b-196">Type</span></span> | <span data-ttu-id="2459b-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="2459b-197">Description</span></span> | <span data-ttu-id="2459b-198">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="2459b-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="2459b-199">String</span><span class="sxs-lookup"><span data-stu-id="2459b-199">String</span></span> | <span data-ttu-id="2459b-200">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="2459b-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="2459b-201">1,5</span><span class="sxs-lookup"><span data-stu-id="2459b-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2459b-202">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2459b-202">Requirements</span></span>

|<span data-ttu-id="2459b-203">Requisito</span><span class="sxs-lookup"><span data-stu-id="2459b-203">Requirement</span></span>| <span data-ttu-id="2459b-204">Valor</span><span class="sxs-lookup"><span data-stu-id="2459b-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="2459b-205">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2459b-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2459b-206">1,5</span><span class="sxs-lookup"><span data-stu-id="2459b-206">1.5</span></span> |
|[<span data-ttu-id="2459b-207">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2459b-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2459b-208">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2459b-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="2459b-209">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2459b-209">SourceProperty: String</span></span>

<span data-ttu-id="2459b-210">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="2459b-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2459b-211">Tipo</span><span class="sxs-lookup"><span data-stu-id="2459b-211">Type</span></span>

*   <span data-ttu-id="2459b-212">String</span><span class="sxs-lookup"><span data-stu-id="2459b-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2459b-213">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="2459b-213">Properties:</span></span>

|<span data-ttu-id="2459b-214">Nome</span><span class="sxs-lookup"><span data-stu-id="2459b-214">Name</span></span>| <span data-ttu-id="2459b-215">Tipo</span><span class="sxs-lookup"><span data-stu-id="2459b-215">Type</span></span>| <span data-ttu-id="2459b-216">Descrição</span><span class="sxs-lookup"><span data-stu-id="2459b-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2459b-217">String</span><span class="sxs-lookup"><span data-stu-id="2459b-217">String</span></span>|<span data-ttu-id="2459b-218">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="2459b-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2459b-219">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2459b-219">String</span></span>|<span data-ttu-id="2459b-220">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="2459b-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2459b-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2459b-221">Requirements</span></span>

|<span data-ttu-id="2459b-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="2459b-222">Requirement</span></span>| <span data-ttu-id="2459b-223">Valor</span><span class="sxs-lookup"><span data-stu-id="2459b-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="2459b-224">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2459b-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2459b-225">1.1</span><span class="sxs-lookup"><span data-stu-id="2459b-225">1.1</span></span>|
|[<span data-ttu-id="2459b-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2459b-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2459b-227">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="2459b-227">Compose or Read</span></span>|
