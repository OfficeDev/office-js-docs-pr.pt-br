---
title: Namespace do Office – conjunto de requisitos 1,5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 7cc8e6acc60c28b44ec7a2b91bb5e388b2618a31
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/06/2020
ms.locfileid: "42554717"
---
# <a name="office"></a><span data-ttu-id="719a3-102">Office</span><span class="sxs-lookup"><span data-stu-id="719a3-102">Office</span></span>

<span data-ttu-id="719a3-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="719a3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="719a3-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="719a3-105">Requirements</span></span>

|<span data-ttu-id="719a3-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="719a3-106">Requirement</span></span>| <span data-ttu-id="719a3-107">Valor</span><span class="sxs-lookup"><span data-stu-id="719a3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="719a3-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="719a3-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="719a3-109">1.1</span><span class="sxs-lookup"><span data-stu-id="719a3-109">1.1</span></span>|
|[<span data-ttu-id="719a3-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="719a3-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="719a3-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="719a3-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="719a3-112">Propriedades</span><span class="sxs-lookup"><span data-stu-id="719a3-112">Properties</span></span>

| <span data-ttu-id="719a3-113">Propriedade</span><span class="sxs-lookup"><span data-stu-id="719a3-113">Property</span></span> | <span data-ttu-id="719a3-114">Modelos</span><span class="sxs-lookup"><span data-stu-id="719a3-114">Modes</span></span> | <span data-ttu-id="719a3-115">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="719a3-115">Return type</span></span> | <span data-ttu-id="719a3-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="719a3-116">Minimum</span></span><br><span data-ttu-id="719a3-117">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="719a3-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="719a3-118">context</span><span class="sxs-lookup"><span data-stu-id="719a3-118">context</span></span>](office.context.md) | <span data-ttu-id="719a3-119">Escrever</span><span class="sxs-lookup"><span data-stu-id="719a3-119">Compose</span></span><br><span data-ttu-id="719a3-120">Ler</span><span class="sxs-lookup"><span data-stu-id="719a3-120">Read</span></span> | [<span data-ttu-id="719a3-121">Context</span><span class="sxs-lookup"><span data-stu-id="719a3-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="719a3-122">1.1</span><span class="sxs-lookup"><span data-stu-id="719a3-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="719a3-123">Enumerações</span><span class="sxs-lookup"><span data-stu-id="719a3-123">Enumerations</span></span>

| <span data-ttu-id="719a3-124">Enumeração</span><span class="sxs-lookup"><span data-stu-id="719a3-124">Enumeration</span></span> | <span data-ttu-id="719a3-125">Modelos</span><span class="sxs-lookup"><span data-stu-id="719a3-125">Modes</span></span> | <span data-ttu-id="719a3-126">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="719a3-126">Return type</span></span> | <span data-ttu-id="719a3-127">Mínimo</span><span class="sxs-lookup"><span data-stu-id="719a3-127">Minimum</span></span><br><span data-ttu-id="719a3-128">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="719a3-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="719a3-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="719a3-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="719a3-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="719a3-130">Compose</span></span><br><span data-ttu-id="719a3-131">Ler</span><span class="sxs-lookup"><span data-stu-id="719a3-131">Read</span></span> | <span data-ttu-id="719a3-132">String</span><span class="sxs-lookup"><span data-stu-id="719a3-132">String</span></span> | [<span data-ttu-id="719a3-133">1.1</span><span class="sxs-lookup"><span data-stu-id="719a3-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="719a3-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="719a3-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="719a3-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="719a3-135">Compose</span></span><br><span data-ttu-id="719a3-136">Ler</span><span class="sxs-lookup"><span data-stu-id="719a3-136">Read</span></span> | <span data-ttu-id="719a3-137">String</span><span class="sxs-lookup"><span data-stu-id="719a3-137">String</span></span> | [<span data-ttu-id="719a3-138">1.1</span><span class="sxs-lookup"><span data-stu-id="719a3-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="719a3-139">EventType</span><span class="sxs-lookup"><span data-stu-id="719a3-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="719a3-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="719a3-140">Compose</span></span><br><span data-ttu-id="719a3-141">Ler</span><span class="sxs-lookup"><span data-stu-id="719a3-141">Read</span></span> | <span data-ttu-id="719a3-142">String</span><span class="sxs-lookup"><span data-stu-id="719a3-142">String</span></span> | [<span data-ttu-id="719a3-143">1,5</span><span class="sxs-lookup"><span data-stu-id="719a3-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="719a3-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="719a3-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="719a3-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="719a3-145">Compose</span></span><br><span data-ttu-id="719a3-146">Ler</span><span class="sxs-lookup"><span data-stu-id="719a3-146">Read</span></span> | <span data-ttu-id="719a3-147">String</span><span class="sxs-lookup"><span data-stu-id="719a3-147">String</span></span> | [<span data-ttu-id="719a3-148">1.1</span><span class="sxs-lookup"><span data-stu-id="719a3-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="719a3-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="719a3-149">Namespaces</span></span>

<span data-ttu-id="719a3-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="719a3-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="719a3-151">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="719a3-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="719a3-152">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="719a3-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="719a3-153">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="719a3-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="719a3-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="719a3-154">Type</span></span>

*   <span data-ttu-id="719a3-155">String</span><span class="sxs-lookup"><span data-stu-id="719a3-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="719a3-156">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="719a3-156">Properties:</span></span>

|<span data-ttu-id="719a3-157">Nome</span><span class="sxs-lookup"><span data-stu-id="719a3-157">Name</span></span>| <span data-ttu-id="719a3-158">Tipo</span><span class="sxs-lookup"><span data-stu-id="719a3-158">Type</span></span>| <span data-ttu-id="719a3-159">Descrição</span><span class="sxs-lookup"><span data-stu-id="719a3-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="719a3-160">String</span><span class="sxs-lookup"><span data-stu-id="719a3-160">String</span></span>|<span data-ttu-id="719a3-161">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="719a3-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="719a3-162">String</span><span class="sxs-lookup"><span data-stu-id="719a3-162">String</span></span>|<span data-ttu-id="719a3-163">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="719a3-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="719a3-164">Requisitos</span><span class="sxs-lookup"><span data-stu-id="719a3-164">Requirements</span></span>

|<span data-ttu-id="719a3-165">Requisito</span><span class="sxs-lookup"><span data-stu-id="719a3-165">Requirement</span></span>| <span data-ttu-id="719a3-166">Valor</span><span class="sxs-lookup"><span data-stu-id="719a3-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="719a3-167">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="719a3-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="719a3-168">1.1</span><span class="sxs-lookup"><span data-stu-id="719a3-168">1.1</span></span>|
|[<span data-ttu-id="719a3-169">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="719a3-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="719a3-170">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="719a3-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="719a3-171">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="719a3-171">CoercionType: String</span></span>

<span data-ttu-id="719a3-172">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="719a3-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="719a3-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="719a3-173">Type</span></span>

*   <span data-ttu-id="719a3-174">String</span><span class="sxs-lookup"><span data-stu-id="719a3-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="719a3-175">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="719a3-175">Properties:</span></span>

|<span data-ttu-id="719a3-176">Nome</span><span class="sxs-lookup"><span data-stu-id="719a3-176">Name</span></span>| <span data-ttu-id="719a3-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="719a3-177">Type</span></span>| <span data-ttu-id="719a3-178">Descrição</span><span class="sxs-lookup"><span data-stu-id="719a3-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="719a3-179">String</span><span class="sxs-lookup"><span data-stu-id="719a3-179">String</span></span>|<span data-ttu-id="719a3-180">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="719a3-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="719a3-181">String</span><span class="sxs-lookup"><span data-stu-id="719a3-181">String</span></span>|<span data-ttu-id="719a3-182">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="719a3-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="719a3-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="719a3-183">Requirements</span></span>

|<span data-ttu-id="719a3-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="719a3-184">Requirement</span></span>| <span data-ttu-id="719a3-185">Valor</span><span class="sxs-lookup"><span data-stu-id="719a3-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="719a3-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="719a3-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="719a3-187">1.1</span><span class="sxs-lookup"><span data-stu-id="719a3-187">1.1</span></span>|
|[<span data-ttu-id="719a3-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="719a3-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="719a3-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="719a3-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="719a3-190">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="719a3-190">EventType: String</span></span>

<span data-ttu-id="719a3-191">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="719a3-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="719a3-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="719a3-192">Type</span></span>

*   <span data-ttu-id="719a3-193">String</span><span class="sxs-lookup"><span data-stu-id="719a3-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="719a3-194">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="719a3-194">Properties:</span></span>

| <span data-ttu-id="719a3-195">Nome</span><span class="sxs-lookup"><span data-stu-id="719a3-195">Name</span></span> | <span data-ttu-id="719a3-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="719a3-196">Type</span></span> | <span data-ttu-id="719a3-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="719a3-197">Description</span></span> | <span data-ttu-id="719a3-198">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="719a3-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="719a3-199">String</span><span class="sxs-lookup"><span data-stu-id="719a3-199">String</span></span> | <span data-ttu-id="719a3-200">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="719a3-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="719a3-201">1,5</span><span class="sxs-lookup"><span data-stu-id="719a3-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="719a3-202">Requisitos</span><span class="sxs-lookup"><span data-stu-id="719a3-202">Requirements</span></span>

|<span data-ttu-id="719a3-203">Requisito</span><span class="sxs-lookup"><span data-stu-id="719a3-203">Requirement</span></span>| <span data-ttu-id="719a3-204">Valor</span><span class="sxs-lookup"><span data-stu-id="719a3-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="719a3-205">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="719a3-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="719a3-206">1,5</span><span class="sxs-lookup"><span data-stu-id="719a3-206">1.5</span></span> |
|[<span data-ttu-id="719a3-207">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="719a3-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="719a3-208">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="719a3-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="719a3-209">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="719a3-209">SourceProperty: String</span></span>

<span data-ttu-id="719a3-210">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="719a3-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="719a3-211">Tipo</span><span class="sxs-lookup"><span data-stu-id="719a3-211">Type</span></span>

*   <span data-ttu-id="719a3-212">String</span><span class="sxs-lookup"><span data-stu-id="719a3-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="719a3-213">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="719a3-213">Properties:</span></span>

|<span data-ttu-id="719a3-214">Nome</span><span class="sxs-lookup"><span data-stu-id="719a3-214">Name</span></span>| <span data-ttu-id="719a3-215">Tipo</span><span class="sxs-lookup"><span data-stu-id="719a3-215">Type</span></span>| <span data-ttu-id="719a3-216">Descrição</span><span class="sxs-lookup"><span data-stu-id="719a3-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="719a3-217">String</span><span class="sxs-lookup"><span data-stu-id="719a3-217">String</span></span>|<span data-ttu-id="719a3-218">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="719a3-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="719a3-219">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="719a3-219">String</span></span>|<span data-ttu-id="719a3-220">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="719a3-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="719a3-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="719a3-221">Requirements</span></span>

|<span data-ttu-id="719a3-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="719a3-222">Requirement</span></span>| <span data-ttu-id="719a3-223">Valor</span><span class="sxs-lookup"><span data-stu-id="719a3-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="719a3-224">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="719a3-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="719a3-225">1.1</span><span class="sxs-lookup"><span data-stu-id="719a3-225">1.1</span></span>|
|[<span data-ttu-id="719a3-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="719a3-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="719a3-227">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="719a3-227">Compose or Read</span></span>|
