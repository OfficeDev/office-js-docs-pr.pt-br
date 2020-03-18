---
title: Namespace do Office – conjunto de requisitos 1,6
description: O modelo de objeto para o namespace de nível superior da API de suplementos do Outlook (versão da API de caixa de correio 1,6).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ae2f863e054016636ebffc3ff3925cee018036a1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717646"
---
# <a name="office"></a><span data-ttu-id="366ce-103">Office</span><span class="sxs-lookup"><span data-stu-id="366ce-103">Office</span></span>

<span data-ttu-id="366ce-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="366ce-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="366ce-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="366ce-106">Requirements</span></span>

|<span data-ttu-id="366ce-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="366ce-107">Requirement</span></span>| <span data-ttu-id="366ce-108">Valor</span><span class="sxs-lookup"><span data-stu-id="366ce-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="366ce-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="366ce-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="366ce-110">1.1</span><span class="sxs-lookup"><span data-stu-id="366ce-110">1.1</span></span>|
|[<span data-ttu-id="366ce-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="366ce-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="366ce-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="366ce-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="366ce-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="366ce-113">Properties</span></span>

| <span data-ttu-id="366ce-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="366ce-114">Property</span></span> | <span data-ttu-id="366ce-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="366ce-115">Modes</span></span> | <span data-ttu-id="366ce-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="366ce-116">Return type</span></span> | <span data-ttu-id="366ce-117">Mínimo</span><span class="sxs-lookup"><span data-stu-id="366ce-117">Minimum</span></span><br><span data-ttu-id="366ce-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="366ce-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="366ce-119">context</span><span class="sxs-lookup"><span data-stu-id="366ce-119">context</span></span>](office.context.md) | <span data-ttu-id="366ce-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="366ce-120">Compose</span></span><br><span data-ttu-id="366ce-121">Ler</span><span class="sxs-lookup"><span data-stu-id="366ce-121">Read</span></span> | [<span data-ttu-id="366ce-122">Context</span><span class="sxs-lookup"><span data-stu-id="366ce-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="366ce-123">1.1</span><span class="sxs-lookup"><span data-stu-id="366ce-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="366ce-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="366ce-124">Enumerations</span></span>

| <span data-ttu-id="366ce-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="366ce-125">Enumeration</span></span> | <span data-ttu-id="366ce-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="366ce-126">Modes</span></span> | <span data-ttu-id="366ce-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="366ce-127">Return type</span></span> | <span data-ttu-id="366ce-128">Mínimo</span><span class="sxs-lookup"><span data-stu-id="366ce-128">Minimum</span></span><br><span data-ttu-id="366ce-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="366ce-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="366ce-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="366ce-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="366ce-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="366ce-131">Compose</span></span><br><span data-ttu-id="366ce-132">Ler</span><span class="sxs-lookup"><span data-stu-id="366ce-132">Read</span></span> | <span data-ttu-id="366ce-133">String</span><span class="sxs-lookup"><span data-stu-id="366ce-133">String</span></span> | [<span data-ttu-id="366ce-134">1.1</span><span class="sxs-lookup"><span data-stu-id="366ce-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="366ce-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="366ce-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="366ce-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="366ce-136">Compose</span></span><br><span data-ttu-id="366ce-137">Ler</span><span class="sxs-lookup"><span data-stu-id="366ce-137">Read</span></span> | <span data-ttu-id="366ce-138">String</span><span class="sxs-lookup"><span data-stu-id="366ce-138">String</span></span> | [<span data-ttu-id="366ce-139">1.1</span><span class="sxs-lookup"><span data-stu-id="366ce-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="366ce-140">EventType</span><span class="sxs-lookup"><span data-stu-id="366ce-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="366ce-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="366ce-141">Compose</span></span><br><span data-ttu-id="366ce-142">Ler</span><span class="sxs-lookup"><span data-stu-id="366ce-142">Read</span></span> | <span data-ttu-id="366ce-143">String</span><span class="sxs-lookup"><span data-stu-id="366ce-143">String</span></span> | [<span data-ttu-id="366ce-144">1,5</span><span class="sxs-lookup"><span data-stu-id="366ce-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="366ce-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="366ce-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="366ce-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="366ce-146">Compose</span></span><br><span data-ttu-id="366ce-147">Ler</span><span class="sxs-lookup"><span data-stu-id="366ce-147">Read</span></span> | <span data-ttu-id="366ce-148">String</span><span class="sxs-lookup"><span data-stu-id="366ce-148">String</span></span> | [<span data-ttu-id="366ce-149">1.1</span><span class="sxs-lookup"><span data-stu-id="366ce-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="366ce-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="366ce-150">Namespaces</span></span>

<span data-ttu-id="366ce-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="366ce-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="366ce-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="366ce-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="366ce-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="366ce-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="366ce-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="366ce-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="366ce-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="366ce-155">Type</span></span>

*   <span data-ttu-id="366ce-156">String</span><span class="sxs-lookup"><span data-stu-id="366ce-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="366ce-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="366ce-157">Properties:</span></span>

|<span data-ttu-id="366ce-158">Nome</span><span class="sxs-lookup"><span data-stu-id="366ce-158">Name</span></span>| <span data-ttu-id="366ce-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="366ce-159">Type</span></span>| <span data-ttu-id="366ce-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="366ce-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="366ce-161">String</span><span class="sxs-lookup"><span data-stu-id="366ce-161">String</span></span>|<span data-ttu-id="366ce-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="366ce-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="366ce-163">String</span><span class="sxs-lookup"><span data-stu-id="366ce-163">String</span></span>|<span data-ttu-id="366ce-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="366ce-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="366ce-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="366ce-165">Requirements</span></span>

|<span data-ttu-id="366ce-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="366ce-166">Requirement</span></span>| <span data-ttu-id="366ce-167">Valor</span><span class="sxs-lookup"><span data-stu-id="366ce-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="366ce-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="366ce-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="366ce-169">1.1</span><span class="sxs-lookup"><span data-stu-id="366ce-169">1.1</span></span>|
|[<span data-ttu-id="366ce-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="366ce-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="366ce-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="366ce-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="366ce-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="366ce-172">CoercionType: String</span></span>

<span data-ttu-id="366ce-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="366ce-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="366ce-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="366ce-174">Type</span></span>

*   <span data-ttu-id="366ce-175">String</span><span class="sxs-lookup"><span data-stu-id="366ce-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="366ce-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="366ce-176">Properties:</span></span>

|<span data-ttu-id="366ce-177">Nome</span><span class="sxs-lookup"><span data-stu-id="366ce-177">Name</span></span>| <span data-ttu-id="366ce-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="366ce-178">Type</span></span>| <span data-ttu-id="366ce-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="366ce-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="366ce-180">String</span><span class="sxs-lookup"><span data-stu-id="366ce-180">String</span></span>|<span data-ttu-id="366ce-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="366ce-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="366ce-182">String</span><span class="sxs-lookup"><span data-stu-id="366ce-182">String</span></span>|<span data-ttu-id="366ce-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="366ce-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="366ce-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="366ce-184">Requirements</span></span>

|<span data-ttu-id="366ce-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="366ce-185">Requirement</span></span>| <span data-ttu-id="366ce-186">Valor</span><span class="sxs-lookup"><span data-stu-id="366ce-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="366ce-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="366ce-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="366ce-188">1.1</span><span class="sxs-lookup"><span data-stu-id="366ce-188">1.1</span></span>|
|[<span data-ttu-id="366ce-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="366ce-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="366ce-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="366ce-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="366ce-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="366ce-191">EventType: String</span></span>

<span data-ttu-id="366ce-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="366ce-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="366ce-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="366ce-193">Type</span></span>

*   <span data-ttu-id="366ce-194">String</span><span class="sxs-lookup"><span data-stu-id="366ce-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="366ce-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="366ce-195">Properties:</span></span>

| <span data-ttu-id="366ce-196">Nome</span><span class="sxs-lookup"><span data-stu-id="366ce-196">Name</span></span> | <span data-ttu-id="366ce-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="366ce-197">Type</span></span> | <span data-ttu-id="366ce-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="366ce-198">Description</span></span> | <span data-ttu-id="366ce-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="366ce-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="366ce-200">String</span><span class="sxs-lookup"><span data-stu-id="366ce-200">String</span></span> | <span data-ttu-id="366ce-201">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="366ce-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="366ce-202">1,5</span><span class="sxs-lookup"><span data-stu-id="366ce-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="366ce-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="366ce-203">Requirements</span></span>

|<span data-ttu-id="366ce-204">Requisito</span><span class="sxs-lookup"><span data-stu-id="366ce-204">Requirement</span></span>| <span data-ttu-id="366ce-205">Valor</span><span class="sxs-lookup"><span data-stu-id="366ce-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="366ce-206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="366ce-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="366ce-207">1,5</span><span class="sxs-lookup"><span data-stu-id="366ce-207">1.5</span></span> |
|[<span data-ttu-id="366ce-208">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="366ce-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="366ce-209">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="366ce-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="366ce-210">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="366ce-210">SourceProperty: String</span></span>

<span data-ttu-id="366ce-211">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="366ce-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="366ce-212">Tipo</span><span class="sxs-lookup"><span data-stu-id="366ce-212">Type</span></span>

*   <span data-ttu-id="366ce-213">String</span><span class="sxs-lookup"><span data-stu-id="366ce-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="366ce-214">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="366ce-214">Properties:</span></span>

|<span data-ttu-id="366ce-215">Nome</span><span class="sxs-lookup"><span data-stu-id="366ce-215">Name</span></span>| <span data-ttu-id="366ce-216">Tipo</span><span class="sxs-lookup"><span data-stu-id="366ce-216">Type</span></span>| <span data-ttu-id="366ce-217">Descrição</span><span class="sxs-lookup"><span data-stu-id="366ce-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="366ce-218">String</span><span class="sxs-lookup"><span data-stu-id="366ce-218">String</span></span>|<span data-ttu-id="366ce-219">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="366ce-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="366ce-220">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="366ce-220">String</span></span>|<span data-ttu-id="366ce-221">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="366ce-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="366ce-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="366ce-222">Requirements</span></span>

|<span data-ttu-id="366ce-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="366ce-223">Requirement</span></span>| <span data-ttu-id="366ce-224">Valor</span><span class="sxs-lookup"><span data-stu-id="366ce-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="366ce-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="366ce-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="366ce-226">1.1</span><span class="sxs-lookup"><span data-stu-id="366ce-226">1.1</span></span>|
|[<span data-ttu-id="366ce-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="366ce-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="366ce-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="366ce-228">Compose or Read</span></span>|
