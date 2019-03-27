---
title: Namespace do Office – conjunto de requisitos 1,7
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 533e997fc7f8be6eb6d3aefefaf023e8c7666af2
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870524"
---
# <a name="office"></a><span data-ttu-id="82cca-102">Office</span><span class="sxs-lookup"><span data-stu-id="82cca-102">Office</span></span>

<span data-ttu-id="82cca-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="82cca-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="82cca-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="82cca-105">Requirements</span></span>

|<span data-ttu-id="82cca-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="82cca-106">Requirement</span></span>| <span data-ttu-id="82cca-107">Valor</span><span class="sxs-lookup"><span data-stu-id="82cca-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="82cca-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="82cca-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="82cca-109">1.0</span><span class="sxs-lookup"><span data-stu-id="82cca-109">1.0</span></span>|
|[<span data-ttu-id="82cca-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="82cca-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="82cca-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="82cca-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="82cca-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="82cca-112">Members and methods</span></span>

| <span data-ttu-id="82cca-113">Membro</span><span class="sxs-lookup"><span data-stu-id="82cca-113">Member</span></span> | <span data-ttu-id="82cca-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="82cca-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="82cca-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="82cca-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="82cca-116">Member</span><span class="sxs-lookup"><span data-stu-id="82cca-116">Member</span></span> |
| [<span data-ttu-id="82cca-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="82cca-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="82cca-118">Member</span><span class="sxs-lookup"><span data-stu-id="82cca-118">Member</span></span> |
| [<span data-ttu-id="82cca-119">EventType</span><span class="sxs-lookup"><span data-stu-id="82cca-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="82cca-120">Member</span><span class="sxs-lookup"><span data-stu-id="82cca-120">Member</span></span> |
| [<span data-ttu-id="82cca-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="82cca-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="82cca-122">Membro</span><span class="sxs-lookup"><span data-stu-id="82cca-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="82cca-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="82cca-123">Namespaces</span></span>

<span data-ttu-id="82cca-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="82cca-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="82cca-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="82cca-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="82cca-126">Membros</span><span class="sxs-lookup"><span data-stu-id="82cca-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="82cca-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="82cca-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="82cca-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="82cca-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="82cca-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="82cca-129">Type</span></span>

*   <span data-ttu-id="82cca-130">String</span><span class="sxs-lookup"><span data-stu-id="82cca-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="82cca-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="82cca-131">Properties:</span></span>

|<span data-ttu-id="82cca-132">Nome</span><span class="sxs-lookup"><span data-stu-id="82cca-132">Name</span></span>| <span data-ttu-id="82cca-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="82cca-133">Type</span></span>| <span data-ttu-id="82cca-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="82cca-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="82cca-135">String</span><span class="sxs-lookup"><span data-stu-id="82cca-135">String</span></span>|<span data-ttu-id="82cca-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="82cca-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="82cca-137">String</span><span class="sxs-lookup"><span data-stu-id="82cca-137">String</span></span>|<span data-ttu-id="82cca-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="82cca-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="82cca-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="82cca-139">Requirements</span></span>

|<span data-ttu-id="82cca-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="82cca-140">Requirement</span></span>| <span data-ttu-id="82cca-141">Valor</span><span class="sxs-lookup"><span data-stu-id="82cca-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="82cca-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="82cca-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="82cca-143">1.0</span><span class="sxs-lookup"><span data-stu-id="82cca-143">1.0</span></span>|
|[<span data-ttu-id="82cca-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="82cca-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="82cca-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="82cca-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="82cca-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="82cca-146">CoercionType :String</span></span>

<span data-ttu-id="82cca-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="82cca-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="82cca-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="82cca-148">Type</span></span>

*   <span data-ttu-id="82cca-149">String</span><span class="sxs-lookup"><span data-stu-id="82cca-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="82cca-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="82cca-150">Properties:</span></span>

|<span data-ttu-id="82cca-151">Nome</span><span class="sxs-lookup"><span data-stu-id="82cca-151">Name</span></span>| <span data-ttu-id="82cca-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="82cca-152">Type</span></span>| <span data-ttu-id="82cca-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="82cca-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="82cca-154">String</span><span class="sxs-lookup"><span data-stu-id="82cca-154">String</span></span>|<span data-ttu-id="82cca-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="82cca-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="82cca-156">String</span><span class="sxs-lookup"><span data-stu-id="82cca-156">String</span></span>|<span data-ttu-id="82cca-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="82cca-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="82cca-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="82cca-158">Requirements</span></span>

|<span data-ttu-id="82cca-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="82cca-159">Requirement</span></span>| <span data-ttu-id="82cca-160">Valor</span><span class="sxs-lookup"><span data-stu-id="82cca-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="82cca-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="82cca-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="82cca-162">1.0</span><span class="sxs-lookup"><span data-stu-id="82cca-162">1.0</span></span>|
|[<span data-ttu-id="82cca-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="82cca-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="82cca-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="82cca-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="82cca-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="82cca-165">EventType :String</span></span>

<span data-ttu-id="82cca-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="82cca-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="82cca-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="82cca-167">Type</span></span>

*   <span data-ttu-id="82cca-168">String</span><span class="sxs-lookup"><span data-stu-id="82cca-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="82cca-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="82cca-169">Properties:</span></span>

| <span data-ttu-id="82cca-170">Nome</span><span class="sxs-lookup"><span data-stu-id="82cca-170">Name</span></span> | <span data-ttu-id="82cca-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="82cca-171">Type</span></span> | <span data-ttu-id="82cca-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="82cca-172">Description</span></span> | <span data-ttu-id="82cca-173">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="82cca-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="82cca-174">String</span><span class="sxs-lookup"><span data-stu-id="82cca-174">String</span></span> | <span data-ttu-id="82cca-175">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="82cca-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="82cca-176">1.7</span><span class="sxs-lookup"><span data-stu-id="82cca-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="82cca-177">String</span><span class="sxs-lookup"><span data-stu-id="82cca-177">String</span></span> | <span data-ttu-id="82cca-178">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="82cca-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="82cca-179">1,5</span><span class="sxs-lookup"><span data-stu-id="82cca-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="82cca-180">String</span><span class="sxs-lookup"><span data-stu-id="82cca-180">String</span></span> | <span data-ttu-id="82cca-181">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="82cca-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="82cca-182">1.7</span><span class="sxs-lookup"><span data-stu-id="82cca-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="82cca-183">String</span><span class="sxs-lookup"><span data-stu-id="82cca-183">String</span></span> | <span data-ttu-id="82cca-184">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="82cca-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="82cca-185">1.7</span><span class="sxs-lookup"><span data-stu-id="82cca-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="82cca-186">Requisitos</span><span class="sxs-lookup"><span data-stu-id="82cca-186">Requirements</span></span>

|<span data-ttu-id="82cca-187">Requisito</span><span class="sxs-lookup"><span data-stu-id="82cca-187">Requirement</span></span>| <span data-ttu-id="82cca-188">Valor</span><span class="sxs-lookup"><span data-stu-id="82cca-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="82cca-189">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="82cca-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="82cca-190">1,5</span><span class="sxs-lookup"><span data-stu-id="82cca-190">1.5</span></span> |
|[<span data-ttu-id="82cca-191">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="82cca-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="82cca-192">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="82cca-192">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="82cca-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="82cca-193">SourceProperty :String</span></span>

<span data-ttu-id="82cca-194">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="82cca-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="82cca-195">Tipo</span><span class="sxs-lookup"><span data-stu-id="82cca-195">Type</span></span>

*   <span data-ttu-id="82cca-196">String</span><span class="sxs-lookup"><span data-stu-id="82cca-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="82cca-197">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="82cca-197">Properties:</span></span>

|<span data-ttu-id="82cca-198">Nome</span><span class="sxs-lookup"><span data-stu-id="82cca-198">Name</span></span>| <span data-ttu-id="82cca-199">Tipo</span><span class="sxs-lookup"><span data-stu-id="82cca-199">Type</span></span>| <span data-ttu-id="82cca-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="82cca-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="82cca-201">String</span><span class="sxs-lookup"><span data-stu-id="82cca-201">String</span></span>|<span data-ttu-id="82cca-202">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="82cca-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="82cca-203">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="82cca-203">String</span></span>|<span data-ttu-id="82cca-204">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="82cca-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="82cca-205">Requisitos</span><span class="sxs-lookup"><span data-stu-id="82cca-205">Requirements</span></span>

|<span data-ttu-id="82cca-206">Requisito</span><span class="sxs-lookup"><span data-stu-id="82cca-206">Requirement</span></span>| <span data-ttu-id="82cca-207">Valor</span><span class="sxs-lookup"><span data-stu-id="82cca-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="82cca-208">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="82cca-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="82cca-209">1.0</span><span class="sxs-lookup"><span data-stu-id="82cca-209">1.0</span></span>|
|[<span data-ttu-id="82cca-210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="82cca-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="82cca-211">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="82cca-211">Compose or Read</span></span>|
