---
title: Namespace do Office – conjunto de requisitos 1,6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0a6360ff7f4e397b878d9a3f744bdbe58347c558
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163658"
---
# <a name="office"></a><span data-ttu-id="590d5-102">Office</span><span class="sxs-lookup"><span data-stu-id="590d5-102">Office</span></span>

<span data-ttu-id="590d5-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="590d5-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="590d5-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="590d5-105">Requirements</span></span>

|<span data-ttu-id="590d5-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="590d5-106">Requirement</span></span>| <span data-ttu-id="590d5-107">Valor</span><span class="sxs-lookup"><span data-stu-id="590d5-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="590d5-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="590d5-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="590d5-109">1.1</span><span class="sxs-lookup"><span data-stu-id="590d5-109">1.1</span></span>|
|[<span data-ttu-id="590d5-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="590d5-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="590d5-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="590d5-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="590d5-112">Propriedades</span><span class="sxs-lookup"><span data-stu-id="590d5-112">Properties</span></span>

| <span data-ttu-id="590d5-113">Propriedade</span><span class="sxs-lookup"><span data-stu-id="590d5-113">Property</span></span> | <span data-ttu-id="590d5-114">Modelos</span><span class="sxs-lookup"><span data-stu-id="590d5-114">Modes</span></span> | <span data-ttu-id="590d5-115">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="590d5-115">Return type</span></span> | <span data-ttu-id="590d5-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="590d5-116">Minimum</span></span><br><span data-ttu-id="590d5-117">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="590d5-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="590d5-118">context</span><span class="sxs-lookup"><span data-stu-id="590d5-118">context</span></span>](office.context.md) | <span data-ttu-id="590d5-119">Escrever</span><span class="sxs-lookup"><span data-stu-id="590d5-119">Compose</span></span><br><span data-ttu-id="590d5-120">Ler</span><span class="sxs-lookup"><span data-stu-id="590d5-120">Read</span></span> | [<span data-ttu-id="590d5-121">Context</span><span class="sxs-lookup"><span data-stu-id="590d5-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="590d5-122">1.1</span><span class="sxs-lookup"><span data-stu-id="590d5-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="590d5-123">Enumerações</span><span class="sxs-lookup"><span data-stu-id="590d5-123">Enumerations</span></span>

| <span data-ttu-id="590d5-124">Enumeração</span><span class="sxs-lookup"><span data-stu-id="590d5-124">Enumeration</span></span> | <span data-ttu-id="590d5-125">Modelos</span><span class="sxs-lookup"><span data-stu-id="590d5-125">Modes</span></span> | <span data-ttu-id="590d5-126">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="590d5-126">Return type</span></span> | <span data-ttu-id="590d5-127">Mínimo</span><span class="sxs-lookup"><span data-stu-id="590d5-127">Minimum</span></span><br><span data-ttu-id="590d5-128">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="590d5-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="590d5-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="590d5-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="590d5-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="590d5-130">Compose</span></span><br><span data-ttu-id="590d5-131">Ler</span><span class="sxs-lookup"><span data-stu-id="590d5-131">Read</span></span> | <span data-ttu-id="590d5-132">String</span><span class="sxs-lookup"><span data-stu-id="590d5-132">String</span></span> | [<span data-ttu-id="590d5-133">1.1</span><span class="sxs-lookup"><span data-stu-id="590d5-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="590d5-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="590d5-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="590d5-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="590d5-135">Compose</span></span><br><span data-ttu-id="590d5-136">Ler</span><span class="sxs-lookup"><span data-stu-id="590d5-136">Read</span></span> | <span data-ttu-id="590d5-137">String</span><span class="sxs-lookup"><span data-stu-id="590d5-137">String</span></span> | [<span data-ttu-id="590d5-138">1.1</span><span class="sxs-lookup"><span data-stu-id="590d5-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="590d5-139">EventType</span><span class="sxs-lookup"><span data-stu-id="590d5-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="590d5-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="590d5-140">Compose</span></span><br><span data-ttu-id="590d5-141">Ler</span><span class="sxs-lookup"><span data-stu-id="590d5-141">Read</span></span> | <span data-ttu-id="590d5-142">String</span><span class="sxs-lookup"><span data-stu-id="590d5-142">String</span></span> | [<span data-ttu-id="590d5-143">1,5</span><span class="sxs-lookup"><span data-stu-id="590d5-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="590d5-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="590d5-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="590d5-145">Escrever</span><span class="sxs-lookup"><span data-stu-id="590d5-145">Compose</span></span><br><span data-ttu-id="590d5-146">Ler</span><span class="sxs-lookup"><span data-stu-id="590d5-146">Read</span></span> | <span data-ttu-id="590d5-147">String</span><span class="sxs-lookup"><span data-stu-id="590d5-147">String</span></span> | [<span data-ttu-id="590d5-148">1.1</span><span class="sxs-lookup"><span data-stu-id="590d5-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="590d5-149">Namespaces</span><span class="sxs-lookup"><span data-stu-id="590d5-149">Namespaces</span></span>

<span data-ttu-id="590d5-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="590d5-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="590d5-151">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="590d5-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="590d5-152">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="590d5-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="590d5-153">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="590d5-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="590d5-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="590d5-154">Type</span></span>

*   <span data-ttu-id="590d5-155">String</span><span class="sxs-lookup"><span data-stu-id="590d5-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="590d5-156">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="590d5-156">Properties:</span></span>

|<span data-ttu-id="590d5-157">Nome</span><span class="sxs-lookup"><span data-stu-id="590d5-157">Name</span></span>| <span data-ttu-id="590d5-158">Tipo</span><span class="sxs-lookup"><span data-stu-id="590d5-158">Type</span></span>| <span data-ttu-id="590d5-159">Descrição</span><span class="sxs-lookup"><span data-stu-id="590d5-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="590d5-160">String</span><span class="sxs-lookup"><span data-stu-id="590d5-160">String</span></span>|<span data-ttu-id="590d5-161">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="590d5-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="590d5-162">String</span><span class="sxs-lookup"><span data-stu-id="590d5-162">String</span></span>|<span data-ttu-id="590d5-163">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="590d5-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="590d5-164">Requisitos</span><span class="sxs-lookup"><span data-stu-id="590d5-164">Requirements</span></span>

|<span data-ttu-id="590d5-165">Requisito</span><span class="sxs-lookup"><span data-stu-id="590d5-165">Requirement</span></span>| <span data-ttu-id="590d5-166">Valor</span><span class="sxs-lookup"><span data-stu-id="590d5-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="590d5-167">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="590d5-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="590d5-168">1.1</span><span class="sxs-lookup"><span data-stu-id="590d5-168">1.1</span></span>|
|[<span data-ttu-id="590d5-169">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="590d5-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="590d5-170">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="590d5-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="590d5-171">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="590d5-171">CoercionType: String</span></span>

<span data-ttu-id="590d5-172">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="590d5-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="590d5-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="590d5-173">Type</span></span>

*   <span data-ttu-id="590d5-174">String</span><span class="sxs-lookup"><span data-stu-id="590d5-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="590d5-175">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="590d5-175">Properties:</span></span>

|<span data-ttu-id="590d5-176">Nome</span><span class="sxs-lookup"><span data-stu-id="590d5-176">Name</span></span>| <span data-ttu-id="590d5-177">Tipo</span><span class="sxs-lookup"><span data-stu-id="590d5-177">Type</span></span>| <span data-ttu-id="590d5-178">Descrição</span><span class="sxs-lookup"><span data-stu-id="590d5-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="590d5-179">String</span><span class="sxs-lookup"><span data-stu-id="590d5-179">String</span></span>|<span data-ttu-id="590d5-180">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="590d5-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="590d5-181">String</span><span class="sxs-lookup"><span data-stu-id="590d5-181">String</span></span>|<span data-ttu-id="590d5-182">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="590d5-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="590d5-183">Requisitos</span><span class="sxs-lookup"><span data-stu-id="590d5-183">Requirements</span></span>

|<span data-ttu-id="590d5-184">Requisito</span><span class="sxs-lookup"><span data-stu-id="590d5-184">Requirement</span></span>| <span data-ttu-id="590d5-185">Valor</span><span class="sxs-lookup"><span data-stu-id="590d5-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="590d5-186">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="590d5-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="590d5-187">1.1</span><span class="sxs-lookup"><span data-stu-id="590d5-187">1.1</span></span>|
|[<span data-ttu-id="590d5-188">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="590d5-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="590d5-189">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="590d5-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="590d5-190">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="590d5-190">EventType: String</span></span>

<span data-ttu-id="590d5-191">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="590d5-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="590d5-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="590d5-192">Type</span></span>

*   <span data-ttu-id="590d5-193">String</span><span class="sxs-lookup"><span data-stu-id="590d5-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="590d5-194">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="590d5-194">Properties:</span></span>

| <span data-ttu-id="590d5-195">Nome</span><span class="sxs-lookup"><span data-stu-id="590d5-195">Name</span></span> | <span data-ttu-id="590d5-196">Tipo</span><span class="sxs-lookup"><span data-stu-id="590d5-196">Type</span></span> | <span data-ttu-id="590d5-197">Descrição</span><span class="sxs-lookup"><span data-stu-id="590d5-197">Description</span></span> | <span data-ttu-id="590d5-198">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="590d5-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="590d5-199">String</span><span class="sxs-lookup"><span data-stu-id="590d5-199">String</span></span> | <span data-ttu-id="590d5-200">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="590d5-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="590d5-201">1,5</span><span class="sxs-lookup"><span data-stu-id="590d5-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="590d5-202">Requisitos</span><span class="sxs-lookup"><span data-stu-id="590d5-202">Requirements</span></span>

|<span data-ttu-id="590d5-203">Requisito</span><span class="sxs-lookup"><span data-stu-id="590d5-203">Requirement</span></span>| <span data-ttu-id="590d5-204">Valor</span><span class="sxs-lookup"><span data-stu-id="590d5-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="590d5-205">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="590d5-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="590d5-206">1,5</span><span class="sxs-lookup"><span data-stu-id="590d5-206">1.5</span></span> |
|[<span data-ttu-id="590d5-207">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="590d5-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="590d5-208">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="590d5-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="590d5-209">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="590d5-209">SourceProperty: String</span></span>

<span data-ttu-id="590d5-210">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="590d5-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="590d5-211">Tipo</span><span class="sxs-lookup"><span data-stu-id="590d5-211">Type</span></span>

*   <span data-ttu-id="590d5-212">String</span><span class="sxs-lookup"><span data-stu-id="590d5-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="590d5-213">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="590d5-213">Properties:</span></span>

|<span data-ttu-id="590d5-214">Nome</span><span class="sxs-lookup"><span data-stu-id="590d5-214">Name</span></span>| <span data-ttu-id="590d5-215">Tipo</span><span class="sxs-lookup"><span data-stu-id="590d5-215">Type</span></span>| <span data-ttu-id="590d5-216">Descrição</span><span class="sxs-lookup"><span data-stu-id="590d5-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="590d5-217">String</span><span class="sxs-lookup"><span data-stu-id="590d5-217">String</span></span>|<span data-ttu-id="590d5-218">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="590d5-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="590d5-219">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="590d5-219">String</span></span>|<span data-ttu-id="590d5-220">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="590d5-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="590d5-221">Requisitos</span><span class="sxs-lookup"><span data-stu-id="590d5-221">Requirements</span></span>

|<span data-ttu-id="590d5-222">Requisito</span><span class="sxs-lookup"><span data-stu-id="590d5-222">Requirement</span></span>| <span data-ttu-id="590d5-223">Valor</span><span class="sxs-lookup"><span data-stu-id="590d5-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="590d5-224">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="590d5-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="590d5-225">1.1</span><span class="sxs-lookup"><span data-stu-id="590d5-225">1.1</span></span>|
|[<span data-ttu-id="590d5-226">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="590d5-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="590d5-227">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="590d5-227">Compose or Read</span></span>|
