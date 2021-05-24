---
title: Office namespace - conjunto de requisitos 1.5
description: Office namespace disponíveis para os Outlook que usam o conjunto de requisitos da API de Caixa de Correio 1.5.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 46b70185ce983721c75093351e47a02eb8b9e7cd
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590852"
---
# <a name="office-mailbox-requirement-set-15"></a><span data-ttu-id="6a871-103">Office (conjunto de requisitos de caixa de correio 1.5)</span><span class="sxs-lookup"><span data-stu-id="6a871-103">Office (Mailbox requirement set 1.5)</span></span>

<span data-ttu-id="6a871-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="6a871-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6a871-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6a871-106">Requirements</span></span>

|<span data-ttu-id="6a871-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="6a871-107">Requirement</span></span>| <span data-ttu-id="6a871-108">Valor</span><span class="sxs-lookup"><span data-stu-id="6a871-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="6a871-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6a871-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6a871-110">1.1</span><span class="sxs-lookup"><span data-stu-id="6a871-110">1.1</span></span>|
|[<span data-ttu-id="6a871-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6a871-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6a871-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6a871-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="6a871-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="6a871-113">Properties</span></span>

| <span data-ttu-id="6a871-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="6a871-114">Property</span></span> | <span data-ttu-id="6a871-115">Modos</span><span class="sxs-lookup"><span data-stu-id="6a871-115">Modes</span></span> | <span data-ttu-id="6a871-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="6a871-116">Return type</span></span> | <span data-ttu-id="6a871-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="6a871-117">Minimum</span></span><br><span data-ttu-id="6a871-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="6a871-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6a871-119">context</span><span class="sxs-lookup"><span data-stu-id="6a871-119">context</span></span>](office.context.md) | <span data-ttu-id="6a871-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="6a871-120">Compose</span></span><br><span data-ttu-id="6a871-121">Ler</span><span class="sxs-lookup"><span data-stu-id="6a871-121">Read</span></span> | [<span data-ttu-id="6a871-122">Context</span><span class="sxs-lookup"><span data-stu-id="6a871-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="6a871-123">1.1</span><span class="sxs-lookup"><span data-stu-id="6a871-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="6a871-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="6a871-124">Enumerations</span></span>

| <span data-ttu-id="6a871-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="6a871-125">Enumeration</span></span> | <span data-ttu-id="6a871-126">Modos</span><span class="sxs-lookup"><span data-stu-id="6a871-126">Modes</span></span> | <span data-ttu-id="6a871-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="6a871-127">Return type</span></span> | <span data-ttu-id="6a871-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="6a871-128">Minimum</span></span><br><span data-ttu-id="6a871-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="6a871-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6a871-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="6a871-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="6a871-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="6a871-131">Compose</span></span><br><span data-ttu-id="6a871-132">Ler</span><span class="sxs-lookup"><span data-stu-id="6a871-132">Read</span></span> | <span data-ttu-id="6a871-133">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a871-133">String</span></span> | [<span data-ttu-id="6a871-134">1.1</span><span class="sxs-lookup"><span data-stu-id="6a871-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6a871-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="6a871-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="6a871-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="6a871-136">Compose</span></span><br><span data-ttu-id="6a871-137">Ler</span><span class="sxs-lookup"><span data-stu-id="6a871-137">Read</span></span> | <span data-ttu-id="6a871-138">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a871-138">String</span></span> | [<span data-ttu-id="6a871-139">1.1</span><span class="sxs-lookup"><span data-stu-id="6a871-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6a871-140">EventType</span><span class="sxs-lookup"><span data-stu-id="6a871-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="6a871-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="6a871-141">Compose</span></span><br><span data-ttu-id="6a871-142">Ler</span><span class="sxs-lookup"><span data-stu-id="6a871-142">Read</span></span> | <span data-ttu-id="6a871-143">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a871-143">String</span></span> | [<span data-ttu-id="6a871-144">1.5</span><span class="sxs-lookup"><span data-stu-id="6a871-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="6a871-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="6a871-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="6a871-146">Escrever</span><span class="sxs-lookup"><span data-stu-id="6a871-146">Compose</span></span><br><span data-ttu-id="6a871-147">Ler</span><span class="sxs-lookup"><span data-stu-id="6a871-147">Read</span></span> | <span data-ttu-id="6a871-148">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a871-148">String</span></span> | [<span data-ttu-id="6a871-149">1.1</span><span class="sxs-lookup"><span data-stu-id="6a871-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="6a871-150">Namespaces</span><span class="sxs-lookup"><span data-stu-id="6a871-150">Namespaces</span></span>

<span data-ttu-id="6a871-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true): inclui várias enumerações específicas Outlook, por exemplo, `ItemType` , , , , , e `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="6a871-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="6a871-152">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="6a871-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="6a871-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="6a871-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="6a871-154">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="6a871-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="6a871-155">Tipo</span><span class="sxs-lookup"><span data-stu-id="6a871-155">Type</span></span>

*   <span data-ttu-id="6a871-156">String</span><span class="sxs-lookup"><span data-stu-id="6a871-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6a871-157">Propriedades</span><span class="sxs-lookup"><span data-stu-id="6a871-157">Properties</span></span>

|<span data-ttu-id="6a871-158">Nome</span><span class="sxs-lookup"><span data-stu-id="6a871-158">Name</span></span>| <span data-ttu-id="6a871-159">Tipo</span><span class="sxs-lookup"><span data-stu-id="6a871-159">Type</span></span>| <span data-ttu-id="6a871-160">Descrição</span><span class="sxs-lookup"><span data-stu-id="6a871-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="6a871-161">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a871-161">String</span></span>|<span data-ttu-id="6a871-162">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="6a871-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="6a871-163">String</span><span class="sxs-lookup"><span data-stu-id="6a871-163">String</span></span>|<span data-ttu-id="6a871-164">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="6a871-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6a871-165">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6a871-165">Requirements</span></span>

|<span data-ttu-id="6a871-166">Requisito</span><span class="sxs-lookup"><span data-stu-id="6a871-166">Requirement</span></span>| <span data-ttu-id="6a871-167">Valor</span><span class="sxs-lookup"><span data-stu-id="6a871-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="6a871-168">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6a871-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6a871-169">1.1</span><span class="sxs-lookup"><span data-stu-id="6a871-169">1.1</span></span>|
|[<span data-ttu-id="6a871-170">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6a871-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6a871-171">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6a871-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="6a871-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="6a871-172">CoercionType: String</span></span>

<span data-ttu-id="6a871-173">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="6a871-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6a871-174">Tipo</span><span class="sxs-lookup"><span data-stu-id="6a871-174">Type</span></span>

*   <span data-ttu-id="6a871-175">String</span><span class="sxs-lookup"><span data-stu-id="6a871-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6a871-176">Propriedades</span><span class="sxs-lookup"><span data-stu-id="6a871-176">Properties</span></span>

|<span data-ttu-id="6a871-177">Nome</span><span class="sxs-lookup"><span data-stu-id="6a871-177">Name</span></span>| <span data-ttu-id="6a871-178">Tipo</span><span class="sxs-lookup"><span data-stu-id="6a871-178">Type</span></span>| <span data-ttu-id="6a871-179">Descrição</span><span class="sxs-lookup"><span data-stu-id="6a871-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="6a871-180">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a871-180">String</span></span>|<span data-ttu-id="6a871-181">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="6a871-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="6a871-182">String</span><span class="sxs-lookup"><span data-stu-id="6a871-182">String</span></span>|<span data-ttu-id="6a871-183">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="6a871-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6a871-184">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6a871-184">Requirements</span></span>

|<span data-ttu-id="6a871-185">Requisito</span><span class="sxs-lookup"><span data-stu-id="6a871-185">Requirement</span></span>| <span data-ttu-id="6a871-186">Valor</span><span class="sxs-lookup"><span data-stu-id="6a871-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="6a871-187">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6a871-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6a871-188">1.1</span><span class="sxs-lookup"><span data-stu-id="6a871-188">1.1</span></span>|
|[<span data-ttu-id="6a871-189">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6a871-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6a871-190">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6a871-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="6a871-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="6a871-191">EventType: String</span></span>

<span data-ttu-id="6a871-192">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="6a871-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="6a871-193">Tipo</span><span class="sxs-lookup"><span data-stu-id="6a871-193">Type</span></span>

*   <span data-ttu-id="6a871-194">String</span><span class="sxs-lookup"><span data-stu-id="6a871-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6a871-195">Propriedades</span><span class="sxs-lookup"><span data-stu-id="6a871-195">Properties</span></span>

| <span data-ttu-id="6a871-196">Nome</span><span class="sxs-lookup"><span data-stu-id="6a871-196">Name</span></span> | <span data-ttu-id="6a871-197">Tipo</span><span class="sxs-lookup"><span data-stu-id="6a871-197">Type</span></span> | <span data-ttu-id="6a871-198">Descrição</span><span class="sxs-lookup"><span data-stu-id="6a871-198">Description</span></span> | <span data-ttu-id="6a871-199">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="6a871-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="6a871-200">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a871-200">String</span></span> | <span data-ttu-id="6a871-201">Um item Outlook diferente é selecionado para exibição enquanto o painel de tarefas é fixado.</span><span class="sxs-lookup"><span data-stu-id="6a871-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="6a871-202">1,5</span><span class="sxs-lookup"><span data-stu-id="6a871-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6a871-203">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6a871-203">Requirements</span></span>

|<span data-ttu-id="6a871-204">Requisito</span><span class="sxs-lookup"><span data-stu-id="6a871-204">Requirement</span></span>| <span data-ttu-id="6a871-205">Valor</span><span class="sxs-lookup"><span data-stu-id="6a871-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="6a871-206">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6a871-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6a871-207">1,5</span><span class="sxs-lookup"><span data-stu-id="6a871-207">1.5</span></span> |
|[<span data-ttu-id="6a871-208">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6a871-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6a871-209">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6a871-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="6a871-210">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="6a871-210">SourceProperty: String</span></span>

<span data-ttu-id="6a871-211">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="6a871-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6a871-212">Tipo</span><span class="sxs-lookup"><span data-stu-id="6a871-212">Type</span></span>

*   <span data-ttu-id="6a871-213">String</span><span class="sxs-lookup"><span data-stu-id="6a871-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6a871-214">Propriedades</span><span class="sxs-lookup"><span data-stu-id="6a871-214">Properties</span></span>

|<span data-ttu-id="6a871-215">Nome</span><span class="sxs-lookup"><span data-stu-id="6a871-215">Name</span></span>| <span data-ttu-id="6a871-216">Tipo</span><span class="sxs-lookup"><span data-stu-id="6a871-216">Type</span></span>| <span data-ttu-id="6a871-217">Descrição</span><span class="sxs-lookup"><span data-stu-id="6a871-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="6a871-218">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="6a871-218">String</span></span>|<span data-ttu-id="6a871-219">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="6a871-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="6a871-220">String</span><span class="sxs-lookup"><span data-stu-id="6a871-220">String</span></span>|<span data-ttu-id="6a871-221">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="6a871-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6a871-222">Requisitos</span><span class="sxs-lookup"><span data-stu-id="6a871-222">Requirements</span></span>

|<span data-ttu-id="6a871-223">Requisito</span><span class="sxs-lookup"><span data-stu-id="6a871-223">Requirement</span></span>| <span data-ttu-id="6a871-224">Valor</span><span class="sxs-lookup"><span data-stu-id="6a871-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="6a871-225">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="6a871-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6a871-226">1.1</span><span class="sxs-lookup"><span data-stu-id="6a871-226">1.1</span></span>|
|[<span data-ttu-id="6a871-227">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="6a871-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6a871-228">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="6a871-228">Compose or Read</span></span>|
