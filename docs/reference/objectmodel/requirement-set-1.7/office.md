---
title: Namespace do Office – conjunto de requisitos 1,7
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: be0223e7ed274abf0e742be13f258c14f6dccf91
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395691"
---
# <a name="office"></a><span data-ttu-id="156f8-102">Office</span><span class="sxs-lookup"><span data-stu-id="156f8-102">Office</span></span>

<span data-ttu-id="156f8-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="156f8-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="156f8-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="156f8-105">Requirements</span></span>

|<span data-ttu-id="156f8-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="156f8-106">Requirement</span></span>| <span data-ttu-id="156f8-107">Valor</span><span class="sxs-lookup"><span data-stu-id="156f8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="156f8-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="156f8-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="156f8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="156f8-109">1.0</span></span>|
|[<span data-ttu-id="156f8-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="156f8-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="156f8-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="156f8-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="156f8-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="156f8-112">Members and methods</span></span>

| <span data-ttu-id="156f8-113">Membro</span><span class="sxs-lookup"><span data-stu-id="156f8-113">Member</span></span> | <span data-ttu-id="156f8-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="156f8-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="156f8-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="156f8-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="156f8-116">Membro</span><span class="sxs-lookup"><span data-stu-id="156f8-116">Member</span></span> |
| [<span data-ttu-id="156f8-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="156f8-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="156f8-118">Membro</span><span class="sxs-lookup"><span data-stu-id="156f8-118">Member</span></span> |
| [<span data-ttu-id="156f8-119">EventType</span><span class="sxs-lookup"><span data-stu-id="156f8-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="156f8-120">Membro</span><span class="sxs-lookup"><span data-stu-id="156f8-120">Member</span></span> |
| [<span data-ttu-id="156f8-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="156f8-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="156f8-122">Membro</span><span class="sxs-lookup"><span data-stu-id="156f8-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="156f8-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="156f8-123">Namespaces</span></span>

<span data-ttu-id="156f8-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="156f8-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="156f8-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): inclui um número de enumerações, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="156f8-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="156f8-126">Members</span><span class="sxs-lookup"><span data-stu-id="156f8-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="156f8-127">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="156f8-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="156f8-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="156f8-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="156f8-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="156f8-129">Type</span></span>

*   <span data-ttu-id="156f8-130">String</span><span class="sxs-lookup"><span data-stu-id="156f8-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="156f8-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="156f8-131">Properties:</span></span>

|<span data-ttu-id="156f8-132">Nome</span><span class="sxs-lookup"><span data-stu-id="156f8-132">Name</span></span>| <span data-ttu-id="156f8-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="156f8-133">Type</span></span>| <span data-ttu-id="156f8-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="156f8-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="156f8-135">String</span><span class="sxs-lookup"><span data-stu-id="156f8-135">String</span></span>|<span data-ttu-id="156f8-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="156f8-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="156f8-137">String</span><span class="sxs-lookup"><span data-stu-id="156f8-137">String</span></span>|<span data-ttu-id="156f8-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="156f8-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="156f8-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="156f8-139">Requirements</span></span>

|<span data-ttu-id="156f8-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="156f8-140">Requirement</span></span>| <span data-ttu-id="156f8-141">Valor</span><span class="sxs-lookup"><span data-stu-id="156f8-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="156f8-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="156f8-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="156f8-143">1.0</span><span class="sxs-lookup"><span data-stu-id="156f8-143">1.0</span></span>|
|[<span data-ttu-id="156f8-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="156f8-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="156f8-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="156f8-145">Compose or Read</span></span>|

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="156f8-146">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="156f8-146">CoercionType: String</span></span>

<span data-ttu-id="156f8-147">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="156f8-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="156f8-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="156f8-148">Type</span></span>

*   <span data-ttu-id="156f8-149">String</span><span class="sxs-lookup"><span data-stu-id="156f8-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="156f8-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="156f8-150">Properties:</span></span>

|<span data-ttu-id="156f8-151">Nome</span><span class="sxs-lookup"><span data-stu-id="156f8-151">Name</span></span>| <span data-ttu-id="156f8-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="156f8-152">Type</span></span>| <span data-ttu-id="156f8-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="156f8-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="156f8-154">String</span><span class="sxs-lookup"><span data-stu-id="156f8-154">String</span></span>|<span data-ttu-id="156f8-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="156f8-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="156f8-156">String</span><span class="sxs-lookup"><span data-stu-id="156f8-156">String</span></span>|<span data-ttu-id="156f8-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="156f8-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="156f8-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="156f8-158">Requirements</span></span>

|<span data-ttu-id="156f8-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="156f8-159">Requirement</span></span>| <span data-ttu-id="156f8-160">Valor</span><span class="sxs-lookup"><span data-stu-id="156f8-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="156f8-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="156f8-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="156f8-162">1.0</span><span class="sxs-lookup"><span data-stu-id="156f8-162">1.0</span></span>|
|[<span data-ttu-id="156f8-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="156f8-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="156f8-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="156f8-164">Compose or Read</span></span>|

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="156f8-165">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="156f8-165">EventType: String</span></span>

<span data-ttu-id="156f8-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="156f8-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="156f8-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="156f8-167">Type</span></span>

*   <span data-ttu-id="156f8-168">String</span><span class="sxs-lookup"><span data-stu-id="156f8-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="156f8-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="156f8-169">Properties:</span></span>

| <span data-ttu-id="156f8-170">Nome</span><span class="sxs-lookup"><span data-stu-id="156f8-170">Name</span></span> | <span data-ttu-id="156f8-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="156f8-171">Type</span></span> | <span data-ttu-id="156f8-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="156f8-172">Description</span></span> | <span data-ttu-id="156f8-173">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="156f8-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="156f8-174">String</span><span class="sxs-lookup"><span data-stu-id="156f8-174">String</span></span> | <span data-ttu-id="156f8-175">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="156f8-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="156f8-176">1.7</span><span class="sxs-lookup"><span data-stu-id="156f8-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="156f8-177">String</span><span class="sxs-lookup"><span data-stu-id="156f8-177">String</span></span> | <span data-ttu-id="156f8-178">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="156f8-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="156f8-179">1,5</span><span class="sxs-lookup"><span data-stu-id="156f8-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="156f8-180">String</span><span class="sxs-lookup"><span data-stu-id="156f8-180">String</span></span> | <span data-ttu-id="156f8-181">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="156f8-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="156f8-182">1.7</span><span class="sxs-lookup"><span data-stu-id="156f8-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="156f8-183">String</span><span class="sxs-lookup"><span data-stu-id="156f8-183">String</span></span> | <span data-ttu-id="156f8-184">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="156f8-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="156f8-185">1.7</span><span class="sxs-lookup"><span data-stu-id="156f8-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="156f8-186">Requisitos</span><span class="sxs-lookup"><span data-stu-id="156f8-186">Requirements</span></span>

|<span data-ttu-id="156f8-187">Requisito</span><span class="sxs-lookup"><span data-stu-id="156f8-187">Requirement</span></span>| <span data-ttu-id="156f8-188">Valor</span><span class="sxs-lookup"><span data-stu-id="156f8-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="156f8-189">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="156f8-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="156f8-190">1,5</span><span class="sxs-lookup"><span data-stu-id="156f8-190">1.5</span></span> |
|[<span data-ttu-id="156f8-191">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="156f8-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="156f8-192">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="156f8-192">Compose or Read</span></span> |

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="156f8-193">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="156f8-193">SourceProperty: String</span></span>

<span data-ttu-id="156f8-194">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="156f8-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="156f8-195">Tipo</span><span class="sxs-lookup"><span data-stu-id="156f8-195">Type</span></span>

*   <span data-ttu-id="156f8-196">String</span><span class="sxs-lookup"><span data-stu-id="156f8-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="156f8-197">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="156f8-197">Properties:</span></span>

|<span data-ttu-id="156f8-198">Nome</span><span class="sxs-lookup"><span data-stu-id="156f8-198">Name</span></span>| <span data-ttu-id="156f8-199">Tipo</span><span class="sxs-lookup"><span data-stu-id="156f8-199">Type</span></span>| <span data-ttu-id="156f8-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="156f8-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="156f8-201">String</span><span class="sxs-lookup"><span data-stu-id="156f8-201">String</span></span>|<span data-ttu-id="156f8-202">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="156f8-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="156f8-203">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="156f8-203">String</span></span>|<span data-ttu-id="156f8-204">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="156f8-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="156f8-205">Requisitos</span><span class="sxs-lookup"><span data-stu-id="156f8-205">Requirements</span></span>

|<span data-ttu-id="156f8-206">Requisito</span><span class="sxs-lookup"><span data-stu-id="156f8-206">Requirement</span></span>| <span data-ttu-id="156f8-207">Valor</span><span class="sxs-lookup"><span data-stu-id="156f8-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="156f8-208">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="156f8-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="156f8-209">1.0</span><span class="sxs-lookup"><span data-stu-id="156f8-209">1.0</span></span>|
|[<span data-ttu-id="156f8-210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="156f8-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="156f8-211">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="156f8-211">Compose or Read</span></span>|
