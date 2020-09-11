---
title: Namespace do Office – conjunto de requisitos 1,5
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,5.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 141fd124ba5778a5ae576c7b4cd2c749a9c4bd6f
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430594"
---
# <a name="office-mailbox-requirement-set-15"></a><span data-ttu-id="514c1-103">Office (conjunto de requisitos de caixa de correio 1,5)</span><span class="sxs-lookup"><span data-stu-id="514c1-103">Office (Mailbox requirement set 1.5)</span></span>

<span data-ttu-id="514c1-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="514c1-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="514c1-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="514c1-106">Requirements</span></span>

|<span data-ttu-id="514c1-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="514c1-107">Requirement</span></span>| <span data-ttu-id="514c1-108">Valor</span><span class="sxs-lookup"><span data-stu-id="514c1-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="514c1-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="514c1-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="514c1-110">1.1</span><span class="sxs-lookup"><span data-stu-id="514c1-110">1.1</span></span>|
|[<span data-ttu-id="514c1-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="514c1-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="514c1-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="514c1-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="514c1-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="514c1-113">Properties</span></span>

| <span data-ttu-id="514c1-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="514c1-114">Property</span></span> | <span data-ttu-id="514c1-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="514c1-115">Modes</span></span> | <span data-ttu-id="514c1-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="514c1-116">Return type</span></span> | <span data-ttu-id="514c1-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="514c1-117">Minimum</span></span><br><span data-ttu-id="514c1-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="514c1-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="514c1-119">context</span><span class="sxs-lookup"><span data-stu-id="514c1-119">context</span></span>](office.context.md) | <span data-ttu-id="514c1-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="514c1-120">Compose</span></span><br><span data-ttu-id="514c1-121">Ler</span><span class="sxs-lookup"><span data-stu-id="514c1-121">Read</span></span> | [<span data-ttu-id="514c1-122">Context</span><span class="sxs-lookup"><span data-stu-id="514c1-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="514c1-123">1.1</span><span class="sxs-lookup"><span data-stu-id="514c1-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="514c1-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="514c1-124">Enumerations</span></span>

| <span data-ttu-id="514c1-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="514c1-125">Enumeration</span></span> | <span data-ttu-id="514c1-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="514c1-126">Modes</span></span> | <span data-ttu-id="514c1-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="514c1-127">Return type</span></span> | <span data-ttu-id="514c1-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="514c1-128">Minimum</span></span><br><span data-ttu-id="514c1-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="514c1-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="514c1-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="514c1-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="514c1-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="514c1-131">Compose</span></span><br><span data-ttu-id="514c1-132">Ler</span><span class="sxs-lookup"><span data-stu-id="514c1-132">Read</span></span> | <span data-ttu-id="514c1-133">String</span><span class="sxs-lookup"><span data-stu-id="514c1-133">String</span></span> | [<span data-ttu-id="514c1-134">1.1</span><span class="sxs-lookup"><span data-stu-id="514c1-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="514c1-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="514c1-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="514c1-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="514c1-136">Compose</span></span><br><span data-ttu-id="514c1-137">Ler</span><span class="sxs-lookup"><span data-stu-id="514c1-137">Read</span></span> | <span data-ttu-id="514c1-138">String</span><span class="sxs-lookup"><span data-stu-id="514c1-138">String</span></span> | [<span data-ttu-id="514c1-139">1.1</span><span class="sxs-lookup"><span data-stu-id="514c1-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="514c1-140">EventType</span><span class="sxs-lookup"><span data-stu-id="514c1-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="514c1-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="514c1-141">Compose</span></span><br><span data-ttu-id="514c1-142">Ler</span><span class="sxs-lookup"><span data-stu-id="514c1-142">Read</span></span> | <span data-ttu-id="514c1-143">String</span><span class="sxs-lookup"><span data-stu-id="514c1-143">String</span></span> | [<span data-ttu-id="514c1-144">1,5</span><span class="sxs-lookup"><span data-stu-id="514c1-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="514c1-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="514c1-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="514c1-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="514c1-146">Compose</span></span><br><span data-ttu-id="514c1-147">Ler</span><span class="sxs-lookup"><span data-stu-id="514c1-147">Read</span></span> | <span data-ttu-id="514c1-148">String</span><span class="sxs-lookup"><span data-stu-id="514c1-148">String</span></span> | [<span data-ttu-id="514c1-149">1.1</span><span class="sxs-lookup"><span data-stu-id="514c1-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="514c1-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="514c1-150">Namespaces</span></span>

<span data-ttu-id="514c1-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true): inclui uma série de enumerações específicas do Outlook, por exemplo,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` e `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="514c1-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="514c1-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="514c1-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="514c1-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="514c1-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="514c1-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="514c1-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="514c1-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="514c1-155">Type</span></span>

*   <span data-ttu-id="514c1-156">String</span><span class="sxs-lookup"><span data-stu-id="514c1-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="514c1-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="514c1-157">Properties:</span></span>

|<span data-ttu-id="514c1-158">Nome</span><span class="sxs-lookup"><span data-stu-id="514c1-158">Name</span></span>| <span data-ttu-id="514c1-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="514c1-159">Type</span></span>| <span data-ttu-id="514c1-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="514c1-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="514c1-161">String</span><span class="sxs-lookup"><span data-stu-id="514c1-161">String</span></span>|<span data-ttu-id="514c1-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="514c1-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="514c1-163">String</span><span class="sxs-lookup"><span data-stu-id="514c1-163">String</span></span>|<span data-ttu-id="514c1-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="514c1-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="514c1-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="514c1-165">Requirements</span></span>

|<span data-ttu-id="514c1-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="514c1-166">Requirement</span></span>| <span data-ttu-id="514c1-167">Valor</span><span class="sxs-lookup"><span data-stu-id="514c1-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="514c1-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="514c1-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="514c1-169">1.1</span><span class="sxs-lookup"><span data-stu-id="514c1-169">1.1</span></span>|
|[<span data-ttu-id="514c1-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="514c1-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="514c1-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="514c1-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="514c1-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="514c1-172">CoercionType: String</span></span>

<span data-ttu-id="514c1-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="514c1-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="514c1-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="514c1-174">Type</span></span>

*   <span data-ttu-id="514c1-175">String</span><span class="sxs-lookup"><span data-stu-id="514c1-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="514c1-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="514c1-176">Properties:</span></span>

|<span data-ttu-id="514c1-177">Nome</span><span class="sxs-lookup"><span data-stu-id="514c1-177">Name</span></span>| <span data-ttu-id="514c1-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="514c1-178">Type</span></span>| <span data-ttu-id="514c1-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="514c1-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="514c1-180">String</span><span class="sxs-lookup"><span data-stu-id="514c1-180">String</span></span>|<span data-ttu-id="514c1-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="514c1-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="514c1-182">String</span><span class="sxs-lookup"><span data-stu-id="514c1-182">String</span></span>|<span data-ttu-id="514c1-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="514c1-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="514c1-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="514c1-184">Requirements</span></span>

|<span data-ttu-id="514c1-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="514c1-185">Requirement</span></span>| <span data-ttu-id="514c1-186">Valor</span><span class="sxs-lookup"><span data-stu-id="514c1-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="514c1-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="514c1-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="514c1-188">1.1</span><span class="sxs-lookup"><span data-stu-id="514c1-188">1.1</span></span>|
|[<span data-ttu-id="514c1-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="514c1-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="514c1-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="514c1-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="514c1-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="514c1-191">EventType: String</span></span>

<span data-ttu-id="514c1-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="514c1-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="514c1-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="514c1-193">Type</span></span>

*   <span data-ttu-id="514c1-194">String</span><span class="sxs-lookup"><span data-stu-id="514c1-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="514c1-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="514c1-195">Properties:</span></span>

| <span data-ttu-id="514c1-196">Nome</span><span class="sxs-lookup"><span data-stu-id="514c1-196">Name</span></span> | <span data-ttu-id="514c1-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="514c1-197">Type</span></span> | <span data-ttu-id="514c1-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="514c1-198">Description</span></span> | <span data-ttu-id="514c1-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="514c1-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="514c1-200">String</span><span class="sxs-lookup"><span data-stu-id="514c1-200">String</span></span> | <span data-ttu-id="514c1-201">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="514c1-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="514c1-202">1,5</span><span class="sxs-lookup"><span data-stu-id="514c1-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="514c1-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="514c1-203">Requirements</span></span>

|<span data-ttu-id="514c1-204">Requisito</span><span class="sxs-lookup"><span data-stu-id="514c1-204">Requirement</span></span>| <span data-ttu-id="514c1-205">Valor</span><span class="sxs-lookup"><span data-stu-id="514c1-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="514c1-206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="514c1-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="514c1-207">1,5</span><span class="sxs-lookup"><span data-stu-id="514c1-207">1.5</span></span> |
|[<span data-ttu-id="514c1-208">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="514c1-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="514c1-209">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="514c1-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="514c1-210">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="514c1-210">SourceProperty: String</span></span>

<span data-ttu-id="514c1-211">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="514c1-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="514c1-212">Tipo</span><span class="sxs-lookup"><span data-stu-id="514c1-212">Type</span></span>

*   <span data-ttu-id="514c1-213">String</span><span class="sxs-lookup"><span data-stu-id="514c1-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="514c1-214">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="514c1-214">Properties:</span></span>

|<span data-ttu-id="514c1-215">Nome</span><span class="sxs-lookup"><span data-stu-id="514c1-215">Name</span></span>| <span data-ttu-id="514c1-216">Tipo</span><span class="sxs-lookup"><span data-stu-id="514c1-216">Type</span></span>| <span data-ttu-id="514c1-217">Descrição</span><span class="sxs-lookup"><span data-stu-id="514c1-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="514c1-218">String</span><span class="sxs-lookup"><span data-stu-id="514c1-218">String</span></span>|<span data-ttu-id="514c1-219">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="514c1-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="514c1-220">String</span><span class="sxs-lookup"><span data-stu-id="514c1-220">String</span></span>|<span data-ttu-id="514c1-221">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="514c1-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="514c1-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="514c1-222">Requirements</span></span>

|<span data-ttu-id="514c1-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="514c1-223">Requirement</span></span>| <span data-ttu-id="514c1-224">Valor</span><span class="sxs-lookup"><span data-stu-id="514c1-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="514c1-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="514c1-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="514c1-226">1.1</span><span class="sxs-lookup"><span data-stu-id="514c1-226">1.1</span></span>|
|[<span data-ttu-id="514c1-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="514c1-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="514c1-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="514c1-228">Compose or Read</span></span>|
