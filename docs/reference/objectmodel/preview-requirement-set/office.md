---
title: Namespace do Office – conjunto de requisitos de visualização
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: eae6f99d166695f24f4a94e89ea4b876bea080ef
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902099"
---
# <a name="office"></a><span data-ttu-id="5dac3-102">Office</span><span class="sxs-lookup"><span data-stu-id="5dac3-102">Office</span></span>

<span data-ttu-id="5dac3-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="5dac3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5dac3-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5dac3-105">Requirements</span></span>

|<span data-ttu-id="5dac3-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="5dac3-106">Requirement</span></span>| <span data-ttu-id="5dac3-107">Valor</span><span class="sxs-lookup"><span data-stu-id="5dac3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="5dac3-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5dac3-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5dac3-109">1.0</span><span class="sxs-lookup"><span data-stu-id="5dac3-109">1.0</span></span>|
|[<span data-ttu-id="5dac3-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5dac3-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5dac3-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5dac3-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5dac3-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="5dac3-112">Members and methods</span></span>

| <span data-ttu-id="5dac3-113">Membro</span><span class="sxs-lookup"><span data-stu-id="5dac3-113">Member</span></span> | <span data-ttu-id="5dac3-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="5dac3-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5dac3-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="5dac3-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="5dac3-116">Member</span><span class="sxs-lookup"><span data-stu-id="5dac3-116">Member</span></span> |
| [<span data-ttu-id="5dac3-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="5dac3-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="5dac3-118">Member</span><span class="sxs-lookup"><span data-stu-id="5dac3-118">Member</span></span> |
| [<span data-ttu-id="5dac3-119">EventType</span><span class="sxs-lookup"><span data-stu-id="5dac3-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="5dac3-120">Member</span><span class="sxs-lookup"><span data-stu-id="5dac3-120">Member</span></span> |
| [<span data-ttu-id="5dac3-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="5dac3-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="5dac3-122">Membro</span><span class="sxs-lookup"><span data-stu-id="5dac3-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="5dac3-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="5dac3-123">Namespaces</span></span>

<span data-ttu-id="5dac3-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="5dac3-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="5dac3-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): inclui um número de enumerações, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="5dac3-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="5dac3-126">Members</span><span class="sxs-lookup"><span data-stu-id="5dac3-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="5dac3-127">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5dac3-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="5dac3-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="5dac3-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5dac3-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="5dac3-129">Type</span></span>

*   <span data-ttu-id="5dac3-130">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5dac3-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5dac3-131">Properties:</span></span>

|<span data-ttu-id="5dac3-132">Nome</span><span class="sxs-lookup"><span data-stu-id="5dac3-132">Name</span></span>| <span data-ttu-id="5dac3-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="5dac3-133">Type</span></span>| <span data-ttu-id="5dac3-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="5dac3-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5dac3-135">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-135">String</span></span>|<span data-ttu-id="5dac3-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="5dac3-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5dac3-137">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-137">String</span></span>|<span data-ttu-id="5dac3-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="5dac3-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5dac3-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5dac3-139">Requirements</span></span>

|<span data-ttu-id="5dac3-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="5dac3-140">Requirement</span></span>| <span data-ttu-id="5dac3-141">Valor</span><span class="sxs-lookup"><span data-stu-id="5dac3-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="5dac3-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5dac3-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5dac3-143">1.0</span><span class="sxs-lookup"><span data-stu-id="5dac3-143">1.0</span></span>|
|[<span data-ttu-id="5dac3-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5dac3-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5dac3-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5dac3-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="5dac3-146">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5dac3-146">CoercionType: String</span></span>

<span data-ttu-id="5dac3-147">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="5dac3-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5dac3-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="5dac3-148">Type</span></span>

*   <span data-ttu-id="5dac3-149">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5dac3-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5dac3-150">Properties:</span></span>

|<span data-ttu-id="5dac3-151">Nome</span><span class="sxs-lookup"><span data-stu-id="5dac3-151">Name</span></span>| <span data-ttu-id="5dac3-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="5dac3-152">Type</span></span>| <span data-ttu-id="5dac3-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="5dac3-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5dac3-154">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-154">String</span></span>|<span data-ttu-id="5dac3-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="5dac3-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5dac3-156">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-156">String</span></span>|<span data-ttu-id="5dac3-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="5dac3-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5dac3-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5dac3-158">Requirements</span></span>

|<span data-ttu-id="5dac3-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="5dac3-159">Requirement</span></span>| <span data-ttu-id="5dac3-160">Valor</span><span class="sxs-lookup"><span data-stu-id="5dac3-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="5dac3-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5dac3-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5dac3-162">1.0</span><span class="sxs-lookup"><span data-stu-id="5dac3-162">1.0</span></span>|
|[<span data-ttu-id="5dac3-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5dac3-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5dac3-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5dac3-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="5dac3-165">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5dac3-165">EventType: String</span></span>

<span data-ttu-id="5dac3-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="5dac3-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="5dac3-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="5dac3-167">Type</span></span>

*   <span data-ttu-id="5dac3-168">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5dac3-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5dac3-169">Properties:</span></span>

| <span data-ttu-id="5dac3-170">Nome</span><span class="sxs-lookup"><span data-stu-id="5dac3-170">Name</span></span> | <span data-ttu-id="5dac3-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="5dac3-171">Type</span></span> | <span data-ttu-id="5dac3-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="5dac3-172">Description</span></span> | <span data-ttu-id="5dac3-173">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="5dac3-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="5dac3-174">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-174">String</span></span> | <span data-ttu-id="5dac3-175">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="5dac3-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="5dac3-176">1.7</span><span class="sxs-lookup"><span data-stu-id="5dac3-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="5dac3-177">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-177">String</span></span> | <span data-ttu-id="5dac3-178">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="5dac3-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="5dac3-179">1,8</span><span class="sxs-lookup"><span data-stu-id="5dac3-179">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="5dac3-180">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-180">String</span></span> | <span data-ttu-id="5dac3-181">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="5dac3-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="5dac3-182">1,8</span><span class="sxs-lookup"><span data-stu-id="5dac3-182">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="5dac3-183">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-183">String</span></span> | <span data-ttu-id="5dac3-184">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="5dac3-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="5dac3-185">1,5</span><span class="sxs-lookup"><span data-stu-id="5dac3-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="5dac3-186">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-186">String</span></span> | <span data-ttu-id="5dac3-187">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="5dac3-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="5dac3-188">Visualização</span><span class="sxs-lookup"><span data-stu-id="5dac3-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="5dac3-189">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-189">String</span></span> | <span data-ttu-id="5dac3-190">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="5dac3-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="5dac3-191">1.7</span><span class="sxs-lookup"><span data-stu-id="5dac3-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="5dac3-192">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-192">String</span></span> | <span data-ttu-id="5dac3-193">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="5dac3-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="5dac3-194">1.7</span><span class="sxs-lookup"><span data-stu-id="5dac3-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5dac3-195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5dac3-195">Requirements</span></span>

|<span data-ttu-id="5dac3-196">Requisito</span><span class="sxs-lookup"><span data-stu-id="5dac3-196">Requirement</span></span>| <span data-ttu-id="5dac3-197">Valor</span><span class="sxs-lookup"><span data-stu-id="5dac3-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="5dac3-198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5dac3-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5dac3-199">1,5</span><span class="sxs-lookup"><span data-stu-id="5dac3-199">1.5</span></span> |
|[<span data-ttu-id="5dac3-200">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5dac3-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5dac3-201">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5dac3-201">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="5dac3-202">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5dac3-202">SourceProperty: String</span></span>

<span data-ttu-id="5dac3-203">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="5dac3-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5dac3-204">Tipo</span><span class="sxs-lookup"><span data-stu-id="5dac3-204">Type</span></span>

*   <span data-ttu-id="5dac3-205">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5dac3-206">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5dac3-206">Properties:</span></span>

|<span data-ttu-id="5dac3-207">Nome</span><span class="sxs-lookup"><span data-stu-id="5dac3-207">Name</span></span>| <span data-ttu-id="5dac3-208">Tipo</span><span class="sxs-lookup"><span data-stu-id="5dac3-208">Type</span></span>| <span data-ttu-id="5dac3-209">Descrição</span><span class="sxs-lookup"><span data-stu-id="5dac3-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5dac3-210">String</span><span class="sxs-lookup"><span data-stu-id="5dac3-210">String</span></span>|<span data-ttu-id="5dac3-211">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5dac3-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5dac3-212">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5dac3-212">String</span></span>|<span data-ttu-id="5dac3-213">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5dac3-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5dac3-214">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5dac3-214">Requirements</span></span>

|<span data-ttu-id="5dac3-215">Requisito</span><span class="sxs-lookup"><span data-stu-id="5dac3-215">Requirement</span></span>| <span data-ttu-id="5dac3-216">Valor</span><span class="sxs-lookup"><span data-stu-id="5dac3-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="5dac3-217">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5dac3-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5dac3-218">1.0</span><span class="sxs-lookup"><span data-stu-id="5dac3-218">1.0</span></span>|
|[<span data-ttu-id="5dac3-219">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5dac3-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5dac3-220">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="5dac3-220">Compose or Read</span></span>|
