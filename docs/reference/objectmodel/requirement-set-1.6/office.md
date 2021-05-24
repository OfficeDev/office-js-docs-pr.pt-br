---
title: Office namespace - conjunto de requisitos 1.6
description: Office namespace disponíveis para os Outlook que usam o conjunto de requisitos da API de Caixa de Correio 1.6.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 40cdb7de0678007b93b9251e7f1e2921ed857338
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590831"
---
# <a name="office-mailbox-requirement-set-16"></a><span data-ttu-id="84223-103">Office (conjunto de requisitos de caixa de correio 1.6)</span><span class="sxs-lookup"><span data-stu-id="84223-103">Office (Mailbox requirement set 1.6)</span></span>

<span data-ttu-id="84223-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="84223-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="84223-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="84223-106">Requirements</span></span>

|<span data-ttu-id="84223-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="84223-107">Requirement</span></span>| <span data-ttu-id="84223-108">Valor</span><span class="sxs-lookup"><span data-stu-id="84223-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="84223-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="84223-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="84223-110">1.1</span><span class="sxs-lookup"><span data-stu-id="84223-110">1.1</span></span>|
|[<span data-ttu-id="84223-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="84223-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="84223-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="84223-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="84223-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="84223-113">Properties</span></span>

| <span data-ttu-id="84223-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="84223-114">Property</span></span> | <span data-ttu-id="84223-115">Modos</span><span class="sxs-lookup"><span data-stu-id="84223-115">Modes</span></span> | <span data-ttu-id="84223-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="84223-116">Return type</span></span> | <span data-ttu-id="84223-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="84223-117">Minimum</span></span><br><span data-ttu-id="84223-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="84223-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="84223-119">context</span><span class="sxs-lookup"><span data-stu-id="84223-119">context</span></span>](office.context.md) | <span data-ttu-id="84223-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="84223-120">Compose</span></span><br><span data-ttu-id="84223-121">Ler</span><span class="sxs-lookup"><span data-stu-id="84223-121">Read</span></span> | [<span data-ttu-id="84223-122">Context</span><span class="sxs-lookup"><span data-stu-id="84223-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="84223-123">1.1</span><span class="sxs-lookup"><span data-stu-id="84223-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="84223-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="84223-124">Enumerations</span></span>

| <span data-ttu-id="84223-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="84223-125">Enumeration</span></span> | <span data-ttu-id="84223-126">Modos</span><span class="sxs-lookup"><span data-stu-id="84223-126">Modes</span></span> | <span data-ttu-id="84223-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="84223-127">Return type</span></span> | <span data-ttu-id="84223-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="84223-128">Minimum</span></span><br><span data-ttu-id="84223-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="84223-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="84223-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="84223-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="84223-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="84223-131">Compose</span></span><br><span data-ttu-id="84223-132">Ler</span><span class="sxs-lookup"><span data-stu-id="84223-132">Read</span></span> | <span data-ttu-id="84223-133">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="84223-133">String</span></span> | [<span data-ttu-id="84223-134">1.1</span><span class="sxs-lookup"><span data-stu-id="84223-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="84223-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="84223-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="84223-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="84223-136">Compose</span></span><br><span data-ttu-id="84223-137">Ler</span><span class="sxs-lookup"><span data-stu-id="84223-137">Read</span></span> | <span data-ttu-id="84223-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="84223-138">String</span></span> | [<span data-ttu-id="84223-139">1.1</span><span class="sxs-lookup"><span data-stu-id="84223-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="84223-140">EventType</span><span class="sxs-lookup"><span data-stu-id="84223-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="84223-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="84223-141">Compose</span></span><br><span data-ttu-id="84223-142">Ler</span><span class="sxs-lookup"><span data-stu-id="84223-142">Read</span></span> | <span data-ttu-id="84223-143">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="84223-143">String</span></span> | [<span data-ttu-id="84223-144">1.5</span><span class="sxs-lookup"><span data-stu-id="84223-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="84223-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="84223-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="84223-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="84223-146">Compose</span></span><br><span data-ttu-id="84223-147">Ler</span><span class="sxs-lookup"><span data-stu-id="84223-147">Read</span></span> | <span data-ttu-id="84223-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="84223-148">String</span></span> | [<span data-ttu-id="84223-149">1.1</span><span class="sxs-lookup"><span data-stu-id="84223-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="84223-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="84223-150">Namespaces</span></span>

<span data-ttu-id="84223-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6&preserve-view=true): inclui várias enumerações específicas Outlook, por exemplo, `ItemType` , , , , , e `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="84223-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="84223-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="84223-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="84223-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="84223-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="84223-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="84223-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="84223-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="84223-155">Type</span></span>

*   <span data-ttu-id="84223-156">String</span><span class="sxs-lookup"><span data-stu-id="84223-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="84223-157">Propriedades</span><span class="sxs-lookup"><span data-stu-id="84223-157">Properties</span></span>

|<span data-ttu-id="84223-158">Nome</span><span class="sxs-lookup"><span data-stu-id="84223-158">Name</span></span>| <span data-ttu-id="84223-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="84223-159">Type</span></span>| <span data-ttu-id="84223-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="84223-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="84223-161">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="84223-161">String</span></span>|<span data-ttu-id="84223-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="84223-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="84223-163">String</span><span class="sxs-lookup"><span data-stu-id="84223-163">String</span></span>|<span data-ttu-id="84223-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="84223-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="84223-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="84223-165">Requirements</span></span>

|<span data-ttu-id="84223-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="84223-166">Requirement</span></span>| <span data-ttu-id="84223-167">Valor</span><span class="sxs-lookup"><span data-stu-id="84223-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="84223-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="84223-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="84223-169">1.1</span><span class="sxs-lookup"><span data-stu-id="84223-169">1.1</span></span>|
|[<span data-ttu-id="84223-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="84223-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="84223-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="84223-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="84223-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="84223-172">CoercionType: String</span></span>

<span data-ttu-id="84223-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="84223-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="84223-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="84223-174">Type</span></span>

*   <span data-ttu-id="84223-175">String</span><span class="sxs-lookup"><span data-stu-id="84223-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="84223-176">Propriedades</span><span class="sxs-lookup"><span data-stu-id="84223-176">Properties</span></span>

|<span data-ttu-id="84223-177">Nome</span><span class="sxs-lookup"><span data-stu-id="84223-177">Name</span></span>| <span data-ttu-id="84223-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="84223-178">Type</span></span>| <span data-ttu-id="84223-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="84223-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="84223-180">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="84223-180">String</span></span>|<span data-ttu-id="84223-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="84223-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="84223-182">String</span><span class="sxs-lookup"><span data-stu-id="84223-182">String</span></span>|<span data-ttu-id="84223-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="84223-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="84223-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="84223-184">Requirements</span></span>

|<span data-ttu-id="84223-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="84223-185">Requirement</span></span>| <span data-ttu-id="84223-186">Valor</span><span class="sxs-lookup"><span data-stu-id="84223-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="84223-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="84223-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="84223-188">1.1</span><span class="sxs-lookup"><span data-stu-id="84223-188">1.1</span></span>|
|[<span data-ttu-id="84223-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="84223-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="84223-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="84223-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="84223-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="84223-191">EventType: String</span></span>

<span data-ttu-id="84223-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="84223-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="84223-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="84223-193">Type</span></span>

*   <span data-ttu-id="84223-194">String</span><span class="sxs-lookup"><span data-stu-id="84223-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="84223-195">Propriedades</span><span class="sxs-lookup"><span data-stu-id="84223-195">Properties</span></span>

| <span data-ttu-id="84223-196">Nome</span><span class="sxs-lookup"><span data-stu-id="84223-196">Name</span></span> | <span data-ttu-id="84223-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="84223-197">Type</span></span> | <span data-ttu-id="84223-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="84223-198">Description</span></span> | <span data-ttu-id="84223-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="84223-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="84223-200">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="84223-200">String</span></span> | <span data-ttu-id="84223-201">Um item Outlook diferente é selecionado para exibição enquanto o painel de tarefas é fixado.</span><span class="sxs-lookup"><span data-stu-id="84223-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="84223-202">1,5</span><span class="sxs-lookup"><span data-stu-id="84223-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="84223-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="84223-203">Requirements</span></span>

|<span data-ttu-id="84223-204">Requisito</span><span class="sxs-lookup"><span data-stu-id="84223-204">Requirement</span></span>| <span data-ttu-id="84223-205">Valor</span><span class="sxs-lookup"><span data-stu-id="84223-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="84223-206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="84223-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="84223-207">1,5</span><span class="sxs-lookup"><span data-stu-id="84223-207">1.5</span></span> |
|[<span data-ttu-id="84223-208">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="84223-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="84223-209">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="84223-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="84223-210">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="84223-210">SourceProperty: String</span></span>

<span data-ttu-id="84223-211">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="84223-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="84223-212">Tipo</span><span class="sxs-lookup"><span data-stu-id="84223-212">Type</span></span>

*   <span data-ttu-id="84223-213">String</span><span class="sxs-lookup"><span data-stu-id="84223-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="84223-214">Propriedades</span><span class="sxs-lookup"><span data-stu-id="84223-214">Properties</span></span>

|<span data-ttu-id="84223-215">Nome</span><span class="sxs-lookup"><span data-stu-id="84223-215">Name</span></span>| <span data-ttu-id="84223-216">Tipo</span><span class="sxs-lookup"><span data-stu-id="84223-216">Type</span></span>| <span data-ttu-id="84223-217">Descrição</span><span class="sxs-lookup"><span data-stu-id="84223-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="84223-218">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="84223-218">String</span></span>|<span data-ttu-id="84223-219">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="84223-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="84223-220">String</span><span class="sxs-lookup"><span data-stu-id="84223-220">String</span></span>|<span data-ttu-id="84223-221">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="84223-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="84223-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="84223-222">Requirements</span></span>

|<span data-ttu-id="84223-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="84223-223">Requirement</span></span>| <span data-ttu-id="84223-224">Valor</span><span class="sxs-lookup"><span data-stu-id="84223-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="84223-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="84223-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="84223-226">1.1</span><span class="sxs-lookup"><span data-stu-id="84223-226">1.1</span></span>|
|[<span data-ttu-id="84223-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="84223-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="84223-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="84223-228">Compose or Read</span></span>|
