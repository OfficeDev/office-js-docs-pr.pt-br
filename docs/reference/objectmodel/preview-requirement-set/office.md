---
title: Namespace do Office – conjunto de requisitos de visualização
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: e426ea87c14c4ad21ebdbfd3df05988ba848b906
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872232"
---
# <a name="office"></a><span data-ttu-id="ae4c6-102">Office</span><span class="sxs-lookup"><span data-stu-id="ae4c6-102">Office</span></span>

<span data-ttu-id="ae4c6-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ae4c6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae4c6-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ae4c6-105">Requirements</span></span>

|<span data-ttu-id="ae4c6-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="ae4c6-106">Requirement</span></span>| <span data-ttu-id="ae4c6-107">Valor</span><span class="sxs-lookup"><span data-stu-id="ae4c6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae4c6-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ae4c6-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae4c6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ae4c6-109">1.0</span></span>|
|[<span data-ttu-id="ae4c6-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ae4c6-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae4c6-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ae4c6-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ae4c6-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="ae4c6-112">Members and methods</span></span>

| <span data-ttu-id="ae4c6-113">Membro</span><span class="sxs-lookup"><span data-stu-id="ae4c6-113">Member</span></span> | <span data-ttu-id="ae4c6-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="ae4c6-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ae4c6-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ae4c6-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ae4c6-116">Member</span><span class="sxs-lookup"><span data-stu-id="ae4c6-116">Member</span></span> |
| [<span data-ttu-id="ae4c6-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ae4c6-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ae4c6-118">Member</span><span class="sxs-lookup"><span data-stu-id="ae4c6-118">Member</span></span> |
| [<span data-ttu-id="ae4c6-119">EventType</span><span class="sxs-lookup"><span data-stu-id="ae4c6-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="ae4c6-120">Member</span><span class="sxs-lookup"><span data-stu-id="ae4c6-120">Member</span></span> |
| [<span data-ttu-id="ae4c6-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ae4c6-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ae4c6-122">Membro</span><span class="sxs-lookup"><span data-stu-id="ae4c6-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ae4c6-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="ae4c6-123">Namespaces</span></span>

<span data-ttu-id="ae4c6-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="ae4c6-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="ae4c6-126">Membros</span><span class="sxs-lookup"><span data-stu-id="ae4c6-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="ae4c6-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="ae4c6-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ae4c6-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="ae4c6-129">Type</span></span>

*   <span data-ttu-id="ae4c6-130">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ae4c6-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ae4c6-131">Properties:</span></span>

|<span data-ttu-id="ae4c6-132">Nome</span><span class="sxs-lookup"><span data-stu-id="ae4c6-132">Name</span></span>| <span data-ttu-id="ae4c6-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="ae4c6-133">Type</span></span>| <span data-ttu-id="ae4c6-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="ae4c6-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ae4c6-135">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-135">String</span></span>|<span data-ttu-id="ae4c6-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ae4c6-137">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-137">String</span></span>|<span data-ttu-id="ae4c6-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ae4c6-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ae4c6-139">Requirements</span></span>

|<span data-ttu-id="ae4c6-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="ae4c6-140">Requirement</span></span>| <span data-ttu-id="ae4c6-141">Valor</span><span class="sxs-lookup"><span data-stu-id="ae4c6-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae4c6-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ae4c6-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae4c6-143">1.0</span><span class="sxs-lookup"><span data-stu-id="ae4c6-143">1.0</span></span>|
|[<span data-ttu-id="ae4c6-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ae4c6-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae4c6-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ae4c6-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="ae4c6-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-146">CoercionType :String</span></span>

<span data-ttu-id="ae4c6-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ae4c6-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="ae4c6-148">Type</span></span>

*   <span data-ttu-id="ae4c6-149">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ae4c6-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ae4c6-150">Properties:</span></span>

|<span data-ttu-id="ae4c6-151">Nome</span><span class="sxs-lookup"><span data-stu-id="ae4c6-151">Name</span></span>| <span data-ttu-id="ae4c6-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="ae4c6-152">Type</span></span>| <span data-ttu-id="ae4c6-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="ae4c6-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ae4c6-154">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-154">String</span></span>|<span data-ttu-id="ae4c6-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ae4c6-156">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-156">String</span></span>|<span data-ttu-id="ae4c6-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ae4c6-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ae4c6-158">Requirements</span></span>

|<span data-ttu-id="ae4c6-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="ae4c6-159">Requirement</span></span>| <span data-ttu-id="ae4c6-160">Valor</span><span class="sxs-lookup"><span data-stu-id="ae4c6-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae4c6-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ae4c6-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae4c6-162">1.0</span><span class="sxs-lookup"><span data-stu-id="ae4c6-162">1.0</span></span>|
|[<span data-ttu-id="ae4c6-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ae4c6-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae4c6-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ae4c6-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="ae4c6-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-165">EventType :String</span></span>

<span data-ttu-id="ae4c6-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="ae4c6-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="ae4c6-167">Type</span></span>

*   <span data-ttu-id="ae4c6-168">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ae4c6-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ae4c6-169">Properties:</span></span>

| <span data-ttu-id="ae4c6-170">Nome</span><span class="sxs-lookup"><span data-stu-id="ae4c6-170">Name</span></span> | <span data-ttu-id="ae4c6-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="ae4c6-171">Type</span></span> | <span data-ttu-id="ae4c6-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="ae4c6-172">Description</span></span> | <span data-ttu-id="ae4c6-173">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="ae4c6-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="ae4c6-174">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-174">String</span></span> | <span data-ttu-id="ae4c6-175">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="ae4c6-176">1.7</span><span class="sxs-lookup"><span data-stu-id="ae4c6-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="ae4c6-177">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-177">String</span></span> | <span data-ttu-id="ae4c6-178">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="ae4c6-179">Visualização</span><span class="sxs-lookup"><span data-stu-id="ae4c6-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="ae4c6-180">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-180">String</span></span> | <span data-ttu-id="ae4c6-181">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="ae4c6-182">Visualização</span><span class="sxs-lookup"><span data-stu-id="ae4c6-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="ae4c6-183">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-183">String</span></span> | <span data-ttu-id="ae4c6-184">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="ae4c6-185">1,5</span><span class="sxs-lookup"><span data-stu-id="ae4c6-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="ae4c6-186">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-186">String</span></span> | <span data-ttu-id="ae4c6-187">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="ae4c6-188">Visualização</span><span class="sxs-lookup"><span data-stu-id="ae4c6-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="ae4c6-189">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-189">String</span></span> | <span data-ttu-id="ae4c6-190">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="ae4c6-191">1.7</span><span class="sxs-lookup"><span data-stu-id="ae4c6-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="ae4c6-192">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-192">String</span></span> | <span data-ttu-id="ae4c6-193">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="ae4c6-194">1.7</span><span class="sxs-lookup"><span data-stu-id="ae4c6-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ae4c6-195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ae4c6-195">Requirements</span></span>

|<span data-ttu-id="ae4c6-196">Requisito</span><span class="sxs-lookup"><span data-stu-id="ae4c6-196">Requirement</span></span>| <span data-ttu-id="ae4c6-197">Valor</span><span class="sxs-lookup"><span data-stu-id="ae4c6-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae4c6-198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ae4c6-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae4c6-199">1,5</span><span class="sxs-lookup"><span data-stu-id="ae4c6-199">1.5</span></span> |
|[<span data-ttu-id="ae4c6-200">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ae4c6-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae4c6-201">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ae4c6-201">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="ae4c6-202">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-202">SourceProperty :String</span></span>

<span data-ttu-id="ae4c6-203">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ae4c6-204">Tipo</span><span class="sxs-lookup"><span data-stu-id="ae4c6-204">Type</span></span>

*   <span data-ttu-id="ae4c6-205">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ae4c6-206">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ae4c6-206">Properties:</span></span>

|<span data-ttu-id="ae4c6-207">Nome</span><span class="sxs-lookup"><span data-stu-id="ae4c6-207">Name</span></span>| <span data-ttu-id="ae4c6-208">Tipo</span><span class="sxs-lookup"><span data-stu-id="ae4c6-208">Type</span></span>| <span data-ttu-id="ae4c6-209">Descrição</span><span class="sxs-lookup"><span data-stu-id="ae4c6-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ae4c6-210">String</span><span class="sxs-lookup"><span data-stu-id="ae4c6-210">String</span></span>|<span data-ttu-id="ae4c6-211">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ae4c6-212">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ae4c6-212">String</span></span>|<span data-ttu-id="ae4c6-213">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ae4c6-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ae4c6-214">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ae4c6-214">Requirements</span></span>

|<span data-ttu-id="ae4c6-215">Requisito</span><span class="sxs-lookup"><span data-stu-id="ae4c6-215">Requirement</span></span>| <span data-ttu-id="ae4c6-216">Valor</span><span class="sxs-lookup"><span data-stu-id="ae4c6-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae4c6-217">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ae4c6-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae4c6-218">1.0</span><span class="sxs-lookup"><span data-stu-id="ae4c6-218">1.0</span></span>|
|[<span data-ttu-id="ae4c6-219">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ae4c6-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ae4c6-220">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="ae4c6-220">Compose or Read</span></span>|
