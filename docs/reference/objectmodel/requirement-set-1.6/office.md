---
title: Namespace do Office – conjunto de requisitos 1,6
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,6.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 97b866a11ad96dbbbebdde6c5ed46c67406441fd
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431441"
---
# <a name="office-mailbox-requirement-set-16"></a><span data-ttu-id="b251e-103">Office (conjunto de requisitos de caixa de correio 1,6)</span><span class="sxs-lookup"><span data-stu-id="b251e-103">Office (Mailbox requirement set 1.6)</span></span>

<span data-ttu-id="b251e-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="b251e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b251e-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b251e-106">Requirements</span></span>

|<span data-ttu-id="b251e-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="b251e-107">Requirement</span></span>| <span data-ttu-id="b251e-108">Valor</span><span class="sxs-lookup"><span data-stu-id="b251e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b251e-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b251e-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b251e-110">1.1</span><span class="sxs-lookup"><span data-stu-id="b251e-110">1.1</span></span>|
|[<span data-ttu-id="b251e-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b251e-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b251e-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b251e-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b251e-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="b251e-113">Properties</span></span>

| <span data-ttu-id="b251e-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="b251e-114">Property</span></span> | <span data-ttu-id="b251e-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="b251e-115">Modes</span></span> | <span data-ttu-id="b251e-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="b251e-116">Return type</span></span> | <span data-ttu-id="b251e-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="b251e-117">Minimum</span></span><br><span data-ttu-id="b251e-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="b251e-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b251e-119">context</span><span class="sxs-lookup"><span data-stu-id="b251e-119">context</span></span>](office.context.md) | <span data-ttu-id="b251e-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="b251e-120">Compose</span></span><br><span data-ttu-id="b251e-121">Ler</span><span class="sxs-lookup"><span data-stu-id="b251e-121">Read</span></span> | [<span data-ttu-id="b251e-122">Context</span><span class="sxs-lookup"><span data-stu-id="b251e-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="b251e-123">1.1</span><span class="sxs-lookup"><span data-stu-id="b251e-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="b251e-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="b251e-124">Enumerations</span></span>

| <span data-ttu-id="b251e-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="b251e-125">Enumeration</span></span> | <span data-ttu-id="b251e-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="b251e-126">Modes</span></span> | <span data-ttu-id="b251e-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="b251e-127">Return type</span></span> | <span data-ttu-id="b251e-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="b251e-128">Minimum</span></span><br><span data-ttu-id="b251e-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="b251e-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b251e-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b251e-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b251e-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="b251e-131">Compose</span></span><br><span data-ttu-id="b251e-132">Ler</span><span class="sxs-lookup"><span data-stu-id="b251e-132">Read</span></span> | <span data-ttu-id="b251e-133">String</span><span class="sxs-lookup"><span data-stu-id="b251e-133">String</span></span> | [<span data-ttu-id="b251e-134">1.1</span><span class="sxs-lookup"><span data-stu-id="b251e-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b251e-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b251e-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b251e-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="b251e-136">Compose</span></span><br><span data-ttu-id="b251e-137">Ler</span><span class="sxs-lookup"><span data-stu-id="b251e-137">Read</span></span> | <span data-ttu-id="b251e-138">String</span><span class="sxs-lookup"><span data-stu-id="b251e-138">String</span></span> | [<span data-ttu-id="b251e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b251e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b251e-140">EventType</span><span class="sxs-lookup"><span data-stu-id="b251e-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="b251e-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="b251e-141">Compose</span></span><br><span data-ttu-id="b251e-142">Ler</span><span class="sxs-lookup"><span data-stu-id="b251e-142">Read</span></span> | <span data-ttu-id="b251e-143">String</span><span class="sxs-lookup"><span data-stu-id="b251e-143">String</span></span> | [<span data-ttu-id="b251e-144">1,5</span><span class="sxs-lookup"><span data-stu-id="b251e-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="b251e-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b251e-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b251e-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="b251e-146">Compose</span></span><br><span data-ttu-id="b251e-147">Ler</span><span class="sxs-lookup"><span data-stu-id="b251e-147">Read</span></span> | <span data-ttu-id="b251e-148">String</span><span class="sxs-lookup"><span data-stu-id="b251e-148">String</span></span> | [<span data-ttu-id="b251e-149">1.1</span><span class="sxs-lookup"><span data-stu-id="b251e-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="b251e-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="b251e-150">Namespaces</span></span>

<span data-ttu-id="b251e-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6&preserve-view=true): inclui uma série de enumerações específicas do Outlook, por exemplo,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` e `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="b251e-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="b251e-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="b251e-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="b251e-153">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b251e-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="b251e-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="b251e-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b251e-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="b251e-155">Type</span></span>

*   <span data-ttu-id="b251e-156">String</span><span class="sxs-lookup"><span data-stu-id="b251e-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b251e-157">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b251e-157">Properties:</span></span>

|<span data-ttu-id="b251e-158">Nome</span><span class="sxs-lookup"><span data-stu-id="b251e-158">Name</span></span>| <span data-ttu-id="b251e-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="b251e-159">Type</span></span>| <span data-ttu-id="b251e-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="b251e-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b251e-161">String</span><span class="sxs-lookup"><span data-stu-id="b251e-161">String</span></span>|<span data-ttu-id="b251e-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="b251e-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b251e-163">String</span><span class="sxs-lookup"><span data-stu-id="b251e-163">String</span></span>|<span data-ttu-id="b251e-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="b251e-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b251e-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b251e-165">Requirements</span></span>

|<span data-ttu-id="b251e-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="b251e-166">Requirement</span></span>| <span data-ttu-id="b251e-167">Valor</span><span class="sxs-lookup"><span data-stu-id="b251e-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="b251e-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b251e-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b251e-169">1.1</span><span class="sxs-lookup"><span data-stu-id="b251e-169">1.1</span></span>|
|[<span data-ttu-id="b251e-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b251e-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b251e-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b251e-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="b251e-172">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b251e-172">CoercionType: String</span></span>

<span data-ttu-id="b251e-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="b251e-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b251e-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="b251e-174">Type</span></span>

*   <span data-ttu-id="b251e-175">String</span><span class="sxs-lookup"><span data-stu-id="b251e-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b251e-176">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b251e-176">Properties:</span></span>

|<span data-ttu-id="b251e-177">Nome</span><span class="sxs-lookup"><span data-stu-id="b251e-177">Name</span></span>| <span data-ttu-id="b251e-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="b251e-178">Type</span></span>| <span data-ttu-id="b251e-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="b251e-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b251e-180">String</span><span class="sxs-lookup"><span data-stu-id="b251e-180">String</span></span>|<span data-ttu-id="b251e-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="b251e-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b251e-182">String</span><span class="sxs-lookup"><span data-stu-id="b251e-182">String</span></span>|<span data-ttu-id="b251e-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="b251e-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b251e-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b251e-184">Requirements</span></span>

|<span data-ttu-id="b251e-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="b251e-185">Requirement</span></span>| <span data-ttu-id="b251e-186">Valor</span><span class="sxs-lookup"><span data-stu-id="b251e-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="b251e-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b251e-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b251e-188">1.1</span><span class="sxs-lookup"><span data-stu-id="b251e-188">1.1</span></span>|
|[<span data-ttu-id="b251e-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b251e-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b251e-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b251e-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="b251e-191">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b251e-191">EventType: String</span></span>

<span data-ttu-id="b251e-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="b251e-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="b251e-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="b251e-193">Type</span></span>

*   <span data-ttu-id="b251e-194">String</span><span class="sxs-lookup"><span data-stu-id="b251e-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b251e-195">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b251e-195">Properties:</span></span>

| <span data-ttu-id="b251e-196">Nome</span><span class="sxs-lookup"><span data-stu-id="b251e-196">Name</span></span> | <span data-ttu-id="b251e-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="b251e-197">Type</span></span> | <span data-ttu-id="b251e-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="b251e-198">Description</span></span> | <span data-ttu-id="b251e-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="b251e-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="b251e-200">String</span><span class="sxs-lookup"><span data-stu-id="b251e-200">String</span></span> | <span data-ttu-id="b251e-201">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="b251e-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="b251e-202">1,5</span><span class="sxs-lookup"><span data-stu-id="b251e-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b251e-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b251e-203">Requirements</span></span>

|<span data-ttu-id="b251e-204">Requisito</span><span class="sxs-lookup"><span data-stu-id="b251e-204">Requirement</span></span>| <span data-ttu-id="b251e-205">Valor</span><span class="sxs-lookup"><span data-stu-id="b251e-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="b251e-206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b251e-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b251e-207">1,5</span><span class="sxs-lookup"><span data-stu-id="b251e-207">1.5</span></span> |
|[<span data-ttu-id="b251e-208">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b251e-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b251e-209">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b251e-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="b251e-210">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b251e-210">SourceProperty: String</span></span>

<span data-ttu-id="b251e-211">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="b251e-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b251e-212">Tipo</span><span class="sxs-lookup"><span data-stu-id="b251e-212">Type</span></span>

*   <span data-ttu-id="b251e-213">String</span><span class="sxs-lookup"><span data-stu-id="b251e-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b251e-214">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b251e-214">Properties:</span></span>

|<span data-ttu-id="b251e-215">Nome</span><span class="sxs-lookup"><span data-stu-id="b251e-215">Name</span></span>| <span data-ttu-id="b251e-216">Tipo</span><span class="sxs-lookup"><span data-stu-id="b251e-216">Type</span></span>| <span data-ttu-id="b251e-217">Descrição</span><span class="sxs-lookup"><span data-stu-id="b251e-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b251e-218">String</span><span class="sxs-lookup"><span data-stu-id="b251e-218">String</span></span>|<span data-ttu-id="b251e-219">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b251e-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b251e-220">String</span><span class="sxs-lookup"><span data-stu-id="b251e-220">String</span></span>|<span data-ttu-id="b251e-221">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b251e-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b251e-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b251e-222">Requirements</span></span>

|<span data-ttu-id="b251e-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="b251e-223">Requirement</span></span>| <span data-ttu-id="b251e-224">Valor</span><span class="sxs-lookup"><span data-stu-id="b251e-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="b251e-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b251e-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b251e-226">1.1</span><span class="sxs-lookup"><span data-stu-id="b251e-226">1.1</span></span>|
|[<span data-ttu-id="b251e-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b251e-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b251e-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b251e-228">Compose or Read</span></span>|
