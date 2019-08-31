---
title: Namespace do Office – conjunto de requisitos 1,7
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 8d22ce8400916dffe12a15bba35f70ceca4db510
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695865"
---
# <a name="office"></a><span data-ttu-id="752de-102">Office</span><span class="sxs-lookup"><span data-stu-id="752de-102">Office</span></span>

<span data-ttu-id="752de-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="752de-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="752de-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="752de-105">Requirements</span></span>

|<span data-ttu-id="752de-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="752de-106">Requirement</span></span>| <span data-ttu-id="752de-107">Valor</span><span class="sxs-lookup"><span data-stu-id="752de-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="752de-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="752de-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="752de-109">1.0</span><span class="sxs-lookup"><span data-stu-id="752de-109">1.0</span></span>|
|[<span data-ttu-id="752de-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="752de-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="752de-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="752de-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="752de-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="752de-112">Members and methods</span></span>

| <span data-ttu-id="752de-113">Membro</span><span class="sxs-lookup"><span data-stu-id="752de-113">Member</span></span> | <span data-ttu-id="752de-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="752de-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="752de-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="752de-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="752de-116">Membro</span><span class="sxs-lookup"><span data-stu-id="752de-116">Member</span></span> |
| [<span data-ttu-id="752de-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="752de-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="752de-118">Membro</span><span class="sxs-lookup"><span data-stu-id="752de-118">Member</span></span> |
| [<span data-ttu-id="752de-119">EventType</span><span class="sxs-lookup"><span data-stu-id="752de-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="752de-120">Membro</span><span class="sxs-lookup"><span data-stu-id="752de-120">Member</span></span> |
| [<span data-ttu-id="752de-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="752de-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="752de-122">Membro</span><span class="sxs-lookup"><span data-stu-id="752de-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="752de-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="752de-123">Namespaces</span></span>

<span data-ttu-id="752de-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="752de-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="752de-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): inclui um número de enumerações, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="752de-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="752de-126">Members</span><span class="sxs-lookup"><span data-stu-id="752de-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="752de-127">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="752de-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="752de-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="752de-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="752de-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="752de-129">Type</span></span>

*   <span data-ttu-id="752de-130">String</span><span class="sxs-lookup"><span data-stu-id="752de-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="752de-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="752de-131">Properties:</span></span>

|<span data-ttu-id="752de-132">Nome</span><span class="sxs-lookup"><span data-stu-id="752de-132">Name</span></span>| <span data-ttu-id="752de-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="752de-133">Type</span></span>| <span data-ttu-id="752de-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="752de-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="752de-135">String</span><span class="sxs-lookup"><span data-stu-id="752de-135">String</span></span>|<span data-ttu-id="752de-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="752de-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="752de-137">String</span><span class="sxs-lookup"><span data-stu-id="752de-137">String</span></span>|<span data-ttu-id="752de-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="752de-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="752de-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="752de-139">Requirements</span></span>

|<span data-ttu-id="752de-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="752de-140">Requirement</span></span>| <span data-ttu-id="752de-141">Valor</span><span class="sxs-lookup"><span data-stu-id="752de-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="752de-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="752de-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="752de-143">1.0</span><span class="sxs-lookup"><span data-stu-id="752de-143">1.0</span></span>|
|[<span data-ttu-id="752de-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="752de-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="752de-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="752de-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="752de-146">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="752de-146">CoercionType: String</span></span>

<span data-ttu-id="752de-147">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="752de-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="752de-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="752de-148">Type</span></span>

*   <span data-ttu-id="752de-149">String</span><span class="sxs-lookup"><span data-stu-id="752de-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="752de-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="752de-150">Properties:</span></span>

|<span data-ttu-id="752de-151">Nome</span><span class="sxs-lookup"><span data-stu-id="752de-151">Name</span></span>| <span data-ttu-id="752de-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="752de-152">Type</span></span>| <span data-ttu-id="752de-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="752de-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="752de-154">String</span><span class="sxs-lookup"><span data-stu-id="752de-154">String</span></span>|<span data-ttu-id="752de-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="752de-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="752de-156">String</span><span class="sxs-lookup"><span data-stu-id="752de-156">String</span></span>|<span data-ttu-id="752de-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="752de-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="752de-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="752de-158">Requirements</span></span>

|<span data-ttu-id="752de-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="752de-159">Requirement</span></span>| <span data-ttu-id="752de-160">Valor</span><span class="sxs-lookup"><span data-stu-id="752de-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="752de-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="752de-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="752de-162">1.0</span><span class="sxs-lookup"><span data-stu-id="752de-162">1.0</span></span>|
|[<span data-ttu-id="752de-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="752de-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="752de-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="752de-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="752de-165">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="752de-165">EventType: String</span></span>

<span data-ttu-id="752de-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="752de-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="752de-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="752de-167">Type</span></span>

*   <span data-ttu-id="752de-168">String</span><span class="sxs-lookup"><span data-stu-id="752de-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="752de-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="752de-169">Properties:</span></span>

| <span data-ttu-id="752de-170">Nome</span><span class="sxs-lookup"><span data-stu-id="752de-170">Name</span></span> | <span data-ttu-id="752de-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="752de-171">Type</span></span> | <span data-ttu-id="752de-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="752de-172">Description</span></span> | <span data-ttu-id="752de-173">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="752de-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="752de-174">String</span><span class="sxs-lookup"><span data-stu-id="752de-174">String</span></span> | <span data-ttu-id="752de-175">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="752de-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="752de-176">1.7</span><span class="sxs-lookup"><span data-stu-id="752de-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="752de-177">String</span><span class="sxs-lookup"><span data-stu-id="752de-177">String</span></span> | <span data-ttu-id="752de-178">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="752de-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="752de-179">1,5</span><span class="sxs-lookup"><span data-stu-id="752de-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="752de-180">String</span><span class="sxs-lookup"><span data-stu-id="752de-180">String</span></span> | <span data-ttu-id="752de-181">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="752de-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="752de-182">1.7</span><span class="sxs-lookup"><span data-stu-id="752de-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="752de-183">String</span><span class="sxs-lookup"><span data-stu-id="752de-183">String</span></span> | <span data-ttu-id="752de-184">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="752de-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="752de-185">1.7</span><span class="sxs-lookup"><span data-stu-id="752de-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="752de-186">Requisitos</span><span class="sxs-lookup"><span data-stu-id="752de-186">Requirements</span></span>

|<span data-ttu-id="752de-187">Requisito</span><span class="sxs-lookup"><span data-stu-id="752de-187">Requirement</span></span>| <span data-ttu-id="752de-188">Valor</span><span class="sxs-lookup"><span data-stu-id="752de-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="752de-189">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="752de-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="752de-190">1,5</span><span class="sxs-lookup"><span data-stu-id="752de-190">1.5</span></span> |
|[<span data-ttu-id="752de-191">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="752de-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="752de-192">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="752de-192">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="752de-193">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="752de-193">SourceProperty: String</span></span>

<span data-ttu-id="752de-194">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="752de-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="752de-195">Tipo</span><span class="sxs-lookup"><span data-stu-id="752de-195">Type</span></span>

*   <span data-ttu-id="752de-196">String</span><span class="sxs-lookup"><span data-stu-id="752de-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="752de-197">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="752de-197">Properties:</span></span>

|<span data-ttu-id="752de-198">Nome</span><span class="sxs-lookup"><span data-stu-id="752de-198">Name</span></span>| <span data-ttu-id="752de-199">Tipo</span><span class="sxs-lookup"><span data-stu-id="752de-199">Type</span></span>| <span data-ttu-id="752de-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="752de-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="752de-201">String</span><span class="sxs-lookup"><span data-stu-id="752de-201">String</span></span>|<span data-ttu-id="752de-202">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="752de-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="752de-203">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="752de-203">String</span></span>|<span data-ttu-id="752de-204">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="752de-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="752de-205">Requisitos</span><span class="sxs-lookup"><span data-stu-id="752de-205">Requirements</span></span>

|<span data-ttu-id="752de-206">Requisito</span><span class="sxs-lookup"><span data-stu-id="752de-206">Requirement</span></span>| <span data-ttu-id="752de-207">Valor</span><span class="sxs-lookup"><span data-stu-id="752de-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="752de-208">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="752de-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="752de-209">1.0</span><span class="sxs-lookup"><span data-stu-id="752de-209">1.0</span></span>|
|[<span data-ttu-id="752de-210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="752de-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="752de-211">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="752de-211">Compose or Read</span></span>|
