---
title: Namespace do Office – conjunto de requisitos 1,8
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 91a0bef2a8280a068763c98b17644bd9268e2fb4
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902130"
---
# <a name="office"></a><span data-ttu-id="f7c3c-102">Office</span><span class="sxs-lookup"><span data-stu-id="f7c3c-102">Office</span></span>

<span data-ttu-id="f7c3c-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f7c3c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f7c3c-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7c3c-105">Requirements</span></span>

|<span data-ttu-id="f7c3c-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7c3c-106">Requirement</span></span>| <span data-ttu-id="f7c3c-107">Valor</span><span class="sxs-lookup"><span data-stu-id="f7c3c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7c3c-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7c3c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7c3c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f7c3c-109">1.0</span></span>|
|[<span data-ttu-id="f7c3c-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7c3c-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7c3c-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7c3c-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f7c3c-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="f7c3c-112">Members and methods</span></span>

| <span data-ttu-id="f7c3c-113">Membro</span><span class="sxs-lookup"><span data-stu-id="f7c3c-113">Member</span></span> | <span data-ttu-id="f7c3c-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7c3c-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f7c3c-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f7c3c-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f7c3c-116">Member</span><span class="sxs-lookup"><span data-stu-id="f7c3c-116">Member</span></span> |
| [<span data-ttu-id="f7c3c-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f7c3c-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f7c3c-118">Member</span><span class="sxs-lookup"><span data-stu-id="f7c3c-118">Member</span></span> |
| [<span data-ttu-id="f7c3c-119">EventType</span><span class="sxs-lookup"><span data-stu-id="f7c3c-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f7c3c-120">Member</span><span class="sxs-lookup"><span data-stu-id="f7c3c-120">Member</span></span> |
| [<span data-ttu-id="f7c3c-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f7c3c-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f7c3c-122">Membro</span><span class="sxs-lookup"><span data-stu-id="f7c3c-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f7c3c-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="f7c3c-123">Namespaces</span></span>

<span data-ttu-id="f7c3c-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f7c3c-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): inclui um número de enumerações, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="f7c3c-126">Members</span><span class="sxs-lookup"><span data-stu-id="f7c3c-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f7c3c-127">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f7c3c-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="f7c3c-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f7c3c-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7c3c-129">Type</span></span>

*   <span data-ttu-id="f7c3c-130">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7c3c-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f7c3c-131">Properties:</span></span>

|<span data-ttu-id="f7c3c-132">Nome</span><span class="sxs-lookup"><span data-stu-id="f7c3c-132">Name</span></span>| <span data-ttu-id="f7c3c-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7c3c-133">Type</span></span>| <span data-ttu-id="f7c3c-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="f7c3c-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f7c3c-135">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-135">String</span></span>|<span data-ttu-id="f7c3c-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f7c3c-137">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-137">String</span></span>|<span data-ttu-id="f7c3c-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f7c3c-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7c3c-139">Requirements</span></span>

|<span data-ttu-id="f7c3c-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7c3c-140">Requirement</span></span>| <span data-ttu-id="f7c3c-141">Valor</span><span class="sxs-lookup"><span data-stu-id="f7c3c-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7c3c-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7c3c-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7c3c-143">1.0</span><span class="sxs-lookup"><span data-stu-id="f7c3c-143">1.0</span></span>|
|[<span data-ttu-id="f7c3c-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7c3c-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7c3c-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7c3c-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f7c3c-146">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f7c3c-146">CoercionType: String</span></span>

<span data-ttu-id="f7c3c-147">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f7c3c-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7c3c-148">Type</span></span>

*   <span data-ttu-id="f7c3c-149">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7c3c-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f7c3c-150">Properties:</span></span>

|<span data-ttu-id="f7c3c-151">Nome</span><span class="sxs-lookup"><span data-stu-id="f7c3c-151">Name</span></span>| <span data-ttu-id="f7c3c-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7c3c-152">Type</span></span>| <span data-ttu-id="f7c3c-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="f7c3c-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f7c3c-154">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-154">String</span></span>|<span data-ttu-id="f7c3c-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f7c3c-156">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-156">String</span></span>|<span data-ttu-id="f7c3c-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f7c3c-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7c3c-158">Requirements</span></span>

|<span data-ttu-id="f7c3c-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7c3c-159">Requirement</span></span>| <span data-ttu-id="f7c3c-160">Valor</span><span class="sxs-lookup"><span data-stu-id="f7c3c-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7c3c-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7c3c-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7c3c-162">1.0</span><span class="sxs-lookup"><span data-stu-id="f7c3c-162">1.0</span></span>|
|[<span data-ttu-id="f7c3c-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7c3c-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7c3c-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7c3c-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="f7c3c-165">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f7c3c-165">EventType: String</span></span>

<span data-ttu-id="f7c3c-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f7c3c-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7c3c-167">Type</span></span>

*   <span data-ttu-id="f7c3c-168">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7c3c-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f7c3c-169">Properties:</span></span>

| <span data-ttu-id="f7c3c-170">Nome</span><span class="sxs-lookup"><span data-stu-id="f7c3c-170">Name</span></span> | <span data-ttu-id="f7c3c-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7c3c-171">Type</span></span> | <span data-ttu-id="f7c3c-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="f7c3c-172">Description</span></span> | <span data-ttu-id="f7c3c-173">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="f7c3c-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="f7c3c-174">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-174">String</span></span> | <span data-ttu-id="f7c3c-175">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="f7c3c-176">1.7</span><span class="sxs-lookup"><span data-stu-id="f7c3c-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="f7c3c-177">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-177">String</span></span> | <span data-ttu-id="f7c3c-178">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="f7c3c-179">1,8</span><span class="sxs-lookup"><span data-stu-id="f7c3c-179">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="f7c3c-180">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-180">String</span></span> | <span data-ttu-id="f7c3c-181">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="f7c3c-182">1,8</span><span class="sxs-lookup"><span data-stu-id="f7c3c-182">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="f7c3c-183">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-183">String</span></span> | <span data-ttu-id="f7c3c-184">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="f7c3c-185">1,5</span><span class="sxs-lookup"><span data-stu-id="f7c3c-185">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="f7c3c-186">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-186">String</span></span> | <span data-ttu-id="f7c3c-187">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="f7c3c-188">1.7</span><span class="sxs-lookup"><span data-stu-id="f7c3c-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="f7c3c-189">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-189">String</span></span> | <span data-ttu-id="f7c3c-190">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="f7c3c-191">1.7</span><span class="sxs-lookup"><span data-stu-id="f7c3c-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f7c3c-192">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7c3c-192">Requirements</span></span>

|<span data-ttu-id="f7c3c-193">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7c3c-193">Requirement</span></span>| <span data-ttu-id="f7c3c-194">Valor</span><span class="sxs-lookup"><span data-stu-id="f7c3c-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7c3c-195">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7c3c-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7c3c-196">1,5</span><span class="sxs-lookup"><span data-stu-id="f7c3c-196">1.5</span></span> |
|[<span data-ttu-id="f7c3c-197">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7c3c-197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7c3c-198">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7c3c-198">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f7c3c-199">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f7c3c-199">SourceProperty: String</span></span>

<span data-ttu-id="f7c3c-200">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f7c3c-201">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7c3c-201">Type</span></span>

*   <span data-ttu-id="f7c3c-202">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7c3c-203">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="f7c3c-203">Properties:</span></span>

|<span data-ttu-id="f7c3c-204">Nome</span><span class="sxs-lookup"><span data-stu-id="f7c3c-204">Name</span></span>| <span data-ttu-id="f7c3c-205">Tipo</span><span class="sxs-lookup"><span data-stu-id="f7c3c-205">Type</span></span>| <span data-ttu-id="f7c3c-206">Descrição</span><span class="sxs-lookup"><span data-stu-id="f7c3c-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f7c3c-207">String</span><span class="sxs-lookup"><span data-stu-id="f7c3c-207">String</span></span>|<span data-ttu-id="f7c3c-208">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f7c3c-209">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="f7c3c-209">String</span></span>|<span data-ttu-id="f7c3c-210">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f7c3c-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f7c3c-211">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f7c3c-211">Requirements</span></span>

|<span data-ttu-id="f7c3c-212">Requisito</span><span class="sxs-lookup"><span data-stu-id="f7c3c-212">Requirement</span></span>| <span data-ttu-id="f7c3c-213">Valor</span><span class="sxs-lookup"><span data-stu-id="f7c3c-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7c3c-214">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="f7c3c-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7c3c-215">1.0</span><span class="sxs-lookup"><span data-stu-id="f7c3c-215">1.0</span></span>|
|[<span data-ttu-id="f7c3c-216">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="f7c3c-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7c3c-217">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="f7c3c-217">Compose or Read</span></span>|
