---
title: Namespace do Office – conjunto de requisitos 1,5
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,5.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 35c42d5c134bbeeb7eab4595b94ed9c721c04884
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612042"
---
# <a name="office-mailbox-requirement-set-15"></a><span data-ttu-id="1954b-103">Office (conjunto de requisitos de caixa de correio 1,5)</span><span class="sxs-lookup"><span data-stu-id="1954b-103">Office (Mailbox requirement set 1.5)</span></span>

<span data-ttu-id="1954b-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="1954b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1954b-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1954b-106">Requirements</span></span>

|<span data-ttu-id="1954b-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="1954b-107">Requirement</span></span>| <span data-ttu-id="1954b-108">Valor</span><span class="sxs-lookup"><span data-stu-id="1954b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1954b-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1954b-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1954b-110">1.1</span><span class="sxs-lookup"><span data-stu-id="1954b-110">1.1</span></span>|
|[<span data-ttu-id="1954b-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1954b-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1954b-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1954b-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="1954b-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="1954b-113">Properties</span></span>

| <span data-ttu-id="1954b-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="1954b-114">Property</span></span> | <span data-ttu-id="1954b-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="1954b-115">Modes</span></span> | <span data-ttu-id="1954b-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="1954b-116">Return type</span></span> | <span data-ttu-id="1954b-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="1954b-117">Minimum</span></span><br><span data-ttu-id="1954b-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="1954b-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1954b-119">context</span><span class="sxs-lookup"><span data-stu-id="1954b-119">context</span></span>](office.context.md) | <span data-ttu-id="1954b-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="1954b-120">Compose</span></span><br><span data-ttu-id="1954b-121">Read</span><span class="sxs-lookup"><span data-stu-id="1954b-121">Read</span></span> | [<span data-ttu-id="1954b-122">Context</span><span class="sxs-lookup"><span data-stu-id="1954b-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="1954b-123">1.1</span><span class="sxs-lookup"><span data-stu-id="1954b-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="1954b-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="1954b-124">Enumerations</span></span>

| <span data-ttu-id="1954b-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="1954b-125">Enumeration</span></span> | <span data-ttu-id="1954b-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="1954b-126">Modes</span></span> | <span data-ttu-id="1954b-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="1954b-127">Return type</span></span> | <span data-ttu-id="1954b-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="1954b-128">Minimum</span></span><br><span data-ttu-id="1954b-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="1954b-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1954b-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="1954b-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="1954b-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="1954b-131">Compose</span></span><br><span data-ttu-id="1954b-132">Read</span><span class="sxs-lookup"><span data-stu-id="1954b-132">Read</span></span> | <span data-ttu-id="1954b-133">String</span><span class="sxs-lookup"><span data-stu-id="1954b-133">String</span></span> | [<span data-ttu-id="1954b-134">1.1</span><span class="sxs-lookup"><span data-stu-id="1954b-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1954b-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="1954b-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="1954b-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="1954b-136">Compose</span></span><br><span data-ttu-id="1954b-137">Read</span><span class="sxs-lookup"><span data-stu-id="1954b-137">Read</span></span> | <span data-ttu-id="1954b-138">String</span><span class="sxs-lookup"><span data-stu-id="1954b-138">String</span></span> | [<span data-ttu-id="1954b-139">1.1</span><span class="sxs-lookup"><span data-stu-id="1954b-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1954b-140">EventType</span><span class="sxs-lookup"><span data-stu-id="1954b-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="1954b-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="1954b-141">Compose</span></span><br><span data-ttu-id="1954b-142">Read</span><span class="sxs-lookup"><span data-stu-id="1954b-142">Read</span></span> | <span data-ttu-id="1954b-143">String</span><span class="sxs-lookup"><span data-stu-id="1954b-143">String</span></span> | [<span data-ttu-id="1954b-144">1,5</span><span class="sxs-lookup"><span data-stu-id="1954b-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="1954b-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="1954b-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="1954b-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="1954b-146">Compose</span></span><br><span data-ttu-id="1954b-147">Read</span><span class="sxs-lookup"><span data-stu-id="1954b-147">Read</span></span> | <span data-ttu-id="1954b-148">String</span><span class="sxs-lookup"><span data-stu-id="1954b-148">String</span></span> | [<span data-ttu-id="1954b-149">1.1</span><span class="sxs-lookup"><span data-stu-id="1954b-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="1954b-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="1954b-150">Namespaces</span></span>

<span data-ttu-id="1954b-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): inclui uma série de enumerações específicas do Outlook, por exemplo,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` e `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="1954b-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="1954b-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="1954b-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="1954b-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1954b-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="1954b-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="1954b-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="1954b-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="1954b-155">Type</span></span>

*   <span data-ttu-id="1954b-156">String</span><span class="sxs-lookup"><span data-stu-id="1954b-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1954b-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1954b-157">Properties:</span></span>

|<span data-ttu-id="1954b-158">Nome</span><span class="sxs-lookup"><span data-stu-id="1954b-158">Name</span></span>| <span data-ttu-id="1954b-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="1954b-159">Type</span></span>| <span data-ttu-id="1954b-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="1954b-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="1954b-161">String</span><span class="sxs-lookup"><span data-stu-id="1954b-161">String</span></span>|<span data-ttu-id="1954b-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="1954b-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="1954b-163">String</span><span class="sxs-lookup"><span data-stu-id="1954b-163">String</span></span>|<span data-ttu-id="1954b-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="1954b-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1954b-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1954b-165">Requirements</span></span>

|<span data-ttu-id="1954b-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="1954b-166">Requirement</span></span>| <span data-ttu-id="1954b-167">Valor</span><span class="sxs-lookup"><span data-stu-id="1954b-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="1954b-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1954b-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1954b-169">1.1</span><span class="sxs-lookup"><span data-stu-id="1954b-169">1.1</span></span>|
|[<span data-ttu-id="1954b-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1954b-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1954b-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1954b-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="1954b-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1954b-172">CoercionType: String</span></span>

<span data-ttu-id="1954b-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="1954b-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1954b-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="1954b-174">Type</span></span>

*   <span data-ttu-id="1954b-175">String</span><span class="sxs-lookup"><span data-stu-id="1954b-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1954b-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1954b-176">Properties:</span></span>

|<span data-ttu-id="1954b-177">Nome</span><span class="sxs-lookup"><span data-stu-id="1954b-177">Name</span></span>| <span data-ttu-id="1954b-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="1954b-178">Type</span></span>| <span data-ttu-id="1954b-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="1954b-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="1954b-180">String</span><span class="sxs-lookup"><span data-stu-id="1954b-180">String</span></span>|<span data-ttu-id="1954b-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="1954b-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="1954b-182">String</span><span class="sxs-lookup"><span data-stu-id="1954b-182">String</span></span>|<span data-ttu-id="1954b-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="1954b-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1954b-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1954b-184">Requirements</span></span>

|<span data-ttu-id="1954b-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="1954b-185">Requirement</span></span>| <span data-ttu-id="1954b-186">Valor</span><span class="sxs-lookup"><span data-stu-id="1954b-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="1954b-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1954b-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1954b-188">1.1</span><span class="sxs-lookup"><span data-stu-id="1954b-188">1.1</span></span>|
|[<span data-ttu-id="1954b-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1954b-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1954b-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1954b-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="1954b-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1954b-191">EventType: String</span></span>

<span data-ttu-id="1954b-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="1954b-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="1954b-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="1954b-193">Type</span></span>

*   <span data-ttu-id="1954b-194">String</span><span class="sxs-lookup"><span data-stu-id="1954b-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1954b-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1954b-195">Properties:</span></span>

| <span data-ttu-id="1954b-196">Nome</span><span class="sxs-lookup"><span data-stu-id="1954b-196">Name</span></span> | <span data-ttu-id="1954b-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="1954b-197">Type</span></span> | <span data-ttu-id="1954b-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="1954b-198">Description</span></span> | <span data-ttu-id="1954b-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="1954b-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="1954b-200">String</span><span class="sxs-lookup"><span data-stu-id="1954b-200">String</span></span> | <span data-ttu-id="1954b-201">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="1954b-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="1954b-202">1,5</span><span class="sxs-lookup"><span data-stu-id="1954b-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1954b-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1954b-203">Requirements</span></span>

|<span data-ttu-id="1954b-204">Requisito</span><span class="sxs-lookup"><span data-stu-id="1954b-204">Requirement</span></span>| <span data-ttu-id="1954b-205">Valor</span><span class="sxs-lookup"><span data-stu-id="1954b-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="1954b-206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1954b-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1954b-207">1,5</span><span class="sxs-lookup"><span data-stu-id="1954b-207">1.5</span></span> |
|[<span data-ttu-id="1954b-208">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1954b-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1954b-209">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1954b-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="1954b-210">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1954b-210">SourceProperty: String</span></span>

<span data-ttu-id="1954b-211">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="1954b-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1954b-212">Tipo</span><span class="sxs-lookup"><span data-stu-id="1954b-212">Type</span></span>

*   <span data-ttu-id="1954b-213">String</span><span class="sxs-lookup"><span data-stu-id="1954b-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1954b-214">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1954b-214">Properties:</span></span>

|<span data-ttu-id="1954b-215">Nome</span><span class="sxs-lookup"><span data-stu-id="1954b-215">Name</span></span>| <span data-ttu-id="1954b-216">Tipo</span><span class="sxs-lookup"><span data-stu-id="1954b-216">Type</span></span>| <span data-ttu-id="1954b-217">Descrição</span><span class="sxs-lookup"><span data-stu-id="1954b-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="1954b-218">String</span><span class="sxs-lookup"><span data-stu-id="1954b-218">String</span></span>|<span data-ttu-id="1954b-219">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1954b-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="1954b-220">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1954b-220">String</span></span>|<span data-ttu-id="1954b-221">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1954b-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1954b-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1954b-222">Requirements</span></span>

|<span data-ttu-id="1954b-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="1954b-223">Requirement</span></span>| <span data-ttu-id="1954b-224">Valor</span><span class="sxs-lookup"><span data-stu-id="1954b-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="1954b-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1954b-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1954b-226">1.1</span><span class="sxs-lookup"><span data-stu-id="1954b-226">1.1</span></span>|
|[<span data-ttu-id="1954b-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1954b-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1954b-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1954b-228">Compose or Read</span></span>|
