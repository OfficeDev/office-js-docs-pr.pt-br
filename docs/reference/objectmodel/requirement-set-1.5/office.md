---
title: Namespace do Office – conjunto de requisitos 1,5
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,5.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 848aa30c07b936c8454b2833d5dce3e1d15ee193
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891345"
---
# <a name="office-mailbox-requirement-set-15"></a><span data-ttu-id="3839a-103">Office (conjunto de requisitos de caixa de correio 1,5)</span><span class="sxs-lookup"><span data-stu-id="3839a-103">Office (Mailbox requirement set 1.5)</span></span>

<span data-ttu-id="3839a-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="3839a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3839a-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3839a-106">Requirements</span></span>

|<span data-ttu-id="3839a-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="3839a-107">Requirement</span></span>| <span data-ttu-id="3839a-108">Valor</span><span class="sxs-lookup"><span data-stu-id="3839a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="3839a-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3839a-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3839a-110">1.1</span><span class="sxs-lookup"><span data-stu-id="3839a-110">1.1</span></span>|
|[<span data-ttu-id="3839a-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3839a-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3839a-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3839a-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="3839a-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="3839a-113">Properties</span></span>

| <span data-ttu-id="3839a-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="3839a-114">Property</span></span> | <span data-ttu-id="3839a-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="3839a-115">Modes</span></span> | <span data-ttu-id="3839a-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="3839a-116">Return type</span></span> | <span data-ttu-id="3839a-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="3839a-117">Minimum</span></span><br><span data-ttu-id="3839a-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="3839a-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3839a-119">context</span><span class="sxs-lookup"><span data-stu-id="3839a-119">context</span></span>](office.context.md) | <span data-ttu-id="3839a-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="3839a-120">Compose</span></span><br><span data-ttu-id="3839a-121">Ler</span><span class="sxs-lookup"><span data-stu-id="3839a-121">Read</span></span> | [<span data-ttu-id="3839a-122">Context</span><span class="sxs-lookup"><span data-stu-id="3839a-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="3839a-123">1.1</span><span class="sxs-lookup"><span data-stu-id="3839a-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="3839a-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="3839a-124">Enumerations</span></span>

| <span data-ttu-id="3839a-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="3839a-125">Enumeration</span></span> | <span data-ttu-id="3839a-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="3839a-126">Modes</span></span> | <span data-ttu-id="3839a-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="3839a-127">Return type</span></span> | <span data-ttu-id="3839a-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="3839a-128">Minimum</span></span><br><span data-ttu-id="3839a-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="3839a-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3839a-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3839a-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3839a-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="3839a-131">Compose</span></span><br><span data-ttu-id="3839a-132">Ler</span><span class="sxs-lookup"><span data-stu-id="3839a-132">Read</span></span> | <span data-ttu-id="3839a-133">String</span><span class="sxs-lookup"><span data-stu-id="3839a-133">String</span></span> | [<span data-ttu-id="3839a-134">1.1</span><span class="sxs-lookup"><span data-stu-id="3839a-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3839a-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3839a-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3839a-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="3839a-136">Compose</span></span><br><span data-ttu-id="3839a-137">Ler</span><span class="sxs-lookup"><span data-stu-id="3839a-137">Read</span></span> | <span data-ttu-id="3839a-138">String</span><span class="sxs-lookup"><span data-stu-id="3839a-138">String</span></span> | [<span data-ttu-id="3839a-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3839a-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3839a-140">EventType</span><span class="sxs-lookup"><span data-stu-id="3839a-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="3839a-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="3839a-141">Compose</span></span><br><span data-ttu-id="3839a-142">Ler</span><span class="sxs-lookup"><span data-stu-id="3839a-142">Read</span></span> | <span data-ttu-id="3839a-143">String</span><span class="sxs-lookup"><span data-stu-id="3839a-143">String</span></span> | [<span data-ttu-id="3839a-144">1,5</span><span class="sxs-lookup"><span data-stu-id="3839a-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="3839a-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3839a-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3839a-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="3839a-146">Compose</span></span><br><span data-ttu-id="3839a-147">Ler</span><span class="sxs-lookup"><span data-stu-id="3839a-147">Read</span></span> | <span data-ttu-id="3839a-148">String</span><span class="sxs-lookup"><span data-stu-id="3839a-148">String</span></span> | [<span data-ttu-id="3839a-149">1.1</span><span class="sxs-lookup"><span data-stu-id="3839a-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="3839a-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="3839a-150">Namespaces</span></span>

<span data-ttu-id="3839a-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="3839a-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="3839a-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="3839a-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="3839a-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3839a-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="3839a-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="3839a-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3839a-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="3839a-155">Type</span></span>

*   <span data-ttu-id="3839a-156">String</span><span class="sxs-lookup"><span data-stu-id="3839a-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3839a-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3839a-157">Properties:</span></span>

|<span data-ttu-id="3839a-158">Nome</span><span class="sxs-lookup"><span data-stu-id="3839a-158">Name</span></span>| <span data-ttu-id="3839a-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="3839a-159">Type</span></span>| <span data-ttu-id="3839a-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="3839a-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3839a-161">String</span><span class="sxs-lookup"><span data-stu-id="3839a-161">String</span></span>|<span data-ttu-id="3839a-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="3839a-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3839a-163">String</span><span class="sxs-lookup"><span data-stu-id="3839a-163">String</span></span>|<span data-ttu-id="3839a-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="3839a-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3839a-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3839a-165">Requirements</span></span>

|<span data-ttu-id="3839a-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="3839a-166">Requirement</span></span>| <span data-ttu-id="3839a-167">Valor</span><span class="sxs-lookup"><span data-stu-id="3839a-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="3839a-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3839a-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3839a-169">1.1</span><span class="sxs-lookup"><span data-stu-id="3839a-169">1.1</span></span>|
|[<span data-ttu-id="3839a-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3839a-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3839a-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3839a-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="3839a-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3839a-172">CoercionType: String</span></span>

<span data-ttu-id="3839a-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="3839a-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3839a-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="3839a-174">Type</span></span>

*   <span data-ttu-id="3839a-175">String</span><span class="sxs-lookup"><span data-stu-id="3839a-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3839a-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3839a-176">Properties:</span></span>

|<span data-ttu-id="3839a-177">Nome</span><span class="sxs-lookup"><span data-stu-id="3839a-177">Name</span></span>| <span data-ttu-id="3839a-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="3839a-178">Type</span></span>| <span data-ttu-id="3839a-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="3839a-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3839a-180">String</span><span class="sxs-lookup"><span data-stu-id="3839a-180">String</span></span>|<span data-ttu-id="3839a-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="3839a-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3839a-182">String</span><span class="sxs-lookup"><span data-stu-id="3839a-182">String</span></span>|<span data-ttu-id="3839a-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="3839a-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3839a-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3839a-184">Requirements</span></span>

|<span data-ttu-id="3839a-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="3839a-185">Requirement</span></span>| <span data-ttu-id="3839a-186">Valor</span><span class="sxs-lookup"><span data-stu-id="3839a-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="3839a-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3839a-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3839a-188">1.1</span><span class="sxs-lookup"><span data-stu-id="3839a-188">1.1</span></span>|
|[<span data-ttu-id="3839a-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3839a-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3839a-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3839a-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="3839a-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3839a-191">EventType: String</span></span>

<span data-ttu-id="3839a-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="3839a-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="3839a-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="3839a-193">Type</span></span>

*   <span data-ttu-id="3839a-194">String</span><span class="sxs-lookup"><span data-stu-id="3839a-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3839a-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3839a-195">Properties:</span></span>

| <span data-ttu-id="3839a-196">Nome</span><span class="sxs-lookup"><span data-stu-id="3839a-196">Name</span></span> | <span data-ttu-id="3839a-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="3839a-197">Type</span></span> | <span data-ttu-id="3839a-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="3839a-198">Description</span></span> | <span data-ttu-id="3839a-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="3839a-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="3839a-200">String</span><span class="sxs-lookup"><span data-stu-id="3839a-200">String</span></span> | <span data-ttu-id="3839a-201">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="3839a-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="3839a-202">1,5</span><span class="sxs-lookup"><span data-stu-id="3839a-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3839a-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3839a-203">Requirements</span></span>

|<span data-ttu-id="3839a-204">Requisito</span><span class="sxs-lookup"><span data-stu-id="3839a-204">Requirement</span></span>| <span data-ttu-id="3839a-205">Valor</span><span class="sxs-lookup"><span data-stu-id="3839a-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="3839a-206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3839a-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3839a-207">1,5</span><span class="sxs-lookup"><span data-stu-id="3839a-207">1.5</span></span> |
|[<span data-ttu-id="3839a-208">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3839a-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3839a-209">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3839a-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="3839a-210">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3839a-210">SourceProperty: String</span></span>

<span data-ttu-id="3839a-211">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="3839a-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3839a-212">Tipo</span><span class="sxs-lookup"><span data-stu-id="3839a-212">Type</span></span>

*   <span data-ttu-id="3839a-213">String</span><span class="sxs-lookup"><span data-stu-id="3839a-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3839a-214">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="3839a-214">Properties:</span></span>

|<span data-ttu-id="3839a-215">Nome</span><span class="sxs-lookup"><span data-stu-id="3839a-215">Name</span></span>| <span data-ttu-id="3839a-216">Tipo</span><span class="sxs-lookup"><span data-stu-id="3839a-216">Type</span></span>| <span data-ttu-id="3839a-217">Descrição</span><span class="sxs-lookup"><span data-stu-id="3839a-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3839a-218">String</span><span class="sxs-lookup"><span data-stu-id="3839a-218">String</span></span>|<span data-ttu-id="3839a-219">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="3839a-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3839a-220">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="3839a-220">String</span></span>|<span data-ttu-id="3839a-221">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="3839a-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3839a-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="3839a-222">Requirements</span></span>

|<span data-ttu-id="3839a-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="3839a-223">Requirement</span></span>| <span data-ttu-id="3839a-224">Valor</span><span class="sxs-lookup"><span data-stu-id="3839a-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="3839a-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="3839a-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3839a-226">1.1</span><span class="sxs-lookup"><span data-stu-id="3839a-226">1.1</span></span>|
|[<span data-ttu-id="3839a-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="3839a-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3839a-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="3839a-228">Compose or Read</span></span>|
