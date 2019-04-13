---
title: Namespace do Office – conjunto de requisitos 1,7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 523df189b28fc568ac32e8d17d4a226b52cbd23c
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838449"
---
# <a name="office"></a><span data-ttu-id="b9bb9-102">Office</span><span class="sxs-lookup"><span data-stu-id="b9bb9-102">Office</span></span>

<span data-ttu-id="b9bb9-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="b9bb9-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9bb9-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9bb9-105">Requirements</span></span>

|<span data-ttu-id="b9bb9-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9bb9-106">Requirement</span></span>| <span data-ttu-id="b9bb9-107">Valor</span><span class="sxs-lookup"><span data-stu-id="b9bb9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9bb9-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9bb9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9bb9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b9bb9-109">1.0</span></span>|
|[<span data-ttu-id="b9bb9-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9bb9-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9bb9-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9bb9-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b9bb9-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="b9bb9-112">Members and methods</span></span>

| <span data-ttu-id="b9bb9-113">Membro</span><span class="sxs-lookup"><span data-stu-id="b9bb9-113">Member</span></span> | <span data-ttu-id="b9bb9-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9bb9-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b9bb9-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b9bb9-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b9bb9-116">Membro</span><span class="sxs-lookup"><span data-stu-id="b9bb9-116">Member</span></span> |
| [<span data-ttu-id="b9bb9-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b9bb9-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b9bb9-118">Membro</span><span class="sxs-lookup"><span data-stu-id="b9bb9-118">Member</span></span> |
| [<span data-ttu-id="b9bb9-119">EventType</span><span class="sxs-lookup"><span data-stu-id="b9bb9-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="b9bb9-120">Membro</span><span class="sxs-lookup"><span data-stu-id="b9bb9-120">Member</span></span> |
| [<span data-ttu-id="b9bb9-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b9bb9-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b9bb9-122">Membro</span><span class="sxs-lookup"><span data-stu-id="b9bb9-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b9bb9-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="b9bb9-123">Namespaces</span></span>

<span data-ttu-id="b9bb9-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="b9bb9-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="b9bb9-126">Membros</span><span class="sxs-lookup"><span data-stu-id="b9bb9-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="b9bb9-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="b9bb9-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="b9bb9-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b9bb9-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9bb9-129">Type</span></span>

*   <span data-ttu-id="b9bb9-130">String</span><span class="sxs-lookup"><span data-stu-id="b9bb9-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b9bb9-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b9bb9-131">Properties:</span></span>

|<span data-ttu-id="b9bb9-132">Nome</span><span class="sxs-lookup"><span data-stu-id="b9bb9-132">Name</span></span>| <span data-ttu-id="b9bb9-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9bb9-133">Type</span></span>| <span data-ttu-id="b9bb9-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="b9bb9-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b9bb9-135">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b9bb9-135">String</span></span>|<span data-ttu-id="b9bb9-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b9bb9-137">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b9bb9-137">String</span></span>|<span data-ttu-id="b9bb9-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9bb9-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9bb9-139">Requirements</span></span>

|<span data-ttu-id="b9bb9-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9bb9-140">Requirement</span></span>| <span data-ttu-id="b9bb9-141">Valor</span><span class="sxs-lookup"><span data-stu-id="b9bb9-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9bb9-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9bb9-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9bb9-143">1.0</span><span class="sxs-lookup"><span data-stu-id="b9bb9-143">1.0</span></span>|
|[<span data-ttu-id="b9bb9-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9bb9-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9bb9-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9bb9-145">Compose or Read</span></span>|

---
---

####  <a name="coerciontype-string"></a><span data-ttu-id="b9bb9-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="b9bb9-146">CoercionType :String</span></span>

<span data-ttu-id="b9bb9-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b9bb9-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9bb9-148">Type</span></span>

*   <span data-ttu-id="b9bb9-149">String</span><span class="sxs-lookup"><span data-stu-id="b9bb9-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b9bb9-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b9bb9-150">Properties:</span></span>

|<span data-ttu-id="b9bb9-151">Nome</span><span class="sxs-lookup"><span data-stu-id="b9bb9-151">Name</span></span>| <span data-ttu-id="b9bb9-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9bb9-152">Type</span></span>| <span data-ttu-id="b9bb9-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="b9bb9-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b9bb9-154">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b9bb9-154">String</span></span>|<span data-ttu-id="b9bb9-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b9bb9-156">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b9bb9-156">String</span></span>|<span data-ttu-id="b9bb9-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9bb9-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9bb9-158">Requirements</span></span>

|<span data-ttu-id="b9bb9-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9bb9-159">Requirement</span></span>| <span data-ttu-id="b9bb9-160">Valor</span><span class="sxs-lookup"><span data-stu-id="b9bb9-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9bb9-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9bb9-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9bb9-162">1.0</span><span class="sxs-lookup"><span data-stu-id="b9bb9-162">1.0</span></span>|
|[<span data-ttu-id="b9bb9-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9bb9-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9bb9-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9bb9-164">Compose or Read</span></span>|

---
---

####  <a name="eventtype-string"></a><span data-ttu-id="b9bb9-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="b9bb9-165">EventType :String</span></span>

<span data-ttu-id="b9bb9-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="b9bb9-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9bb9-167">Type</span></span>

*   <span data-ttu-id="b9bb9-168">String</span><span class="sxs-lookup"><span data-stu-id="b9bb9-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b9bb9-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b9bb9-169">Properties:</span></span>

| <span data-ttu-id="b9bb9-170">Nome</span><span class="sxs-lookup"><span data-stu-id="b9bb9-170">Name</span></span> | <span data-ttu-id="b9bb9-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9bb9-171">Type</span></span> | <span data-ttu-id="b9bb9-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="b9bb9-172">Description</span></span> | <span data-ttu-id="b9bb9-173">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="b9bb9-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="b9bb9-174">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b9bb9-174">String</span></span> | <span data-ttu-id="b9bb9-175">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="b9bb9-176">1.7</span><span class="sxs-lookup"><span data-stu-id="b9bb9-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="b9bb9-177">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b9bb9-177">String</span></span> | <span data-ttu-id="b9bb9-178">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="b9bb9-179">1,5</span><span class="sxs-lookup"><span data-stu-id="b9bb9-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="b9bb9-180">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b9bb9-180">String</span></span> | <span data-ttu-id="b9bb9-181">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="b9bb9-182">1.7</span><span class="sxs-lookup"><span data-stu-id="b9bb9-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="b9bb9-183">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b9bb9-183">String</span></span> | <span data-ttu-id="b9bb9-184">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="b9bb9-185">1.7</span><span class="sxs-lookup"><span data-stu-id="b9bb9-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b9bb9-186">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9bb9-186">Requirements</span></span>

|<span data-ttu-id="b9bb9-187">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9bb9-187">Requirement</span></span>| <span data-ttu-id="b9bb9-188">Valor</span><span class="sxs-lookup"><span data-stu-id="b9bb9-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9bb9-189">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9bb9-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9bb9-190">1,5</span><span class="sxs-lookup"><span data-stu-id="b9bb9-190">1.5</span></span> |
|[<span data-ttu-id="b9bb9-191">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9bb9-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9bb9-192">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9bb9-192">Compose or Read</span></span> |

---
---

####  <a name="sourceproperty-string"></a><span data-ttu-id="b9bb9-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="b9bb9-193">SourceProperty :String</span></span>

<span data-ttu-id="b9bb9-194">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b9bb9-195">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9bb9-195">Type</span></span>

*   <span data-ttu-id="b9bb9-196">String</span><span class="sxs-lookup"><span data-stu-id="b9bb9-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b9bb9-197">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b9bb9-197">Properties:</span></span>

|<span data-ttu-id="b9bb9-198">Nome</span><span class="sxs-lookup"><span data-stu-id="b9bb9-198">Name</span></span>| <span data-ttu-id="b9bb9-199">Tipo</span><span class="sxs-lookup"><span data-stu-id="b9bb9-199">Type</span></span>| <span data-ttu-id="b9bb9-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="b9bb9-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b9bb9-201">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b9bb9-201">String</span></span>|<span data-ttu-id="b9bb9-202">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b9bb9-203">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b9bb9-203">String</span></span>|<span data-ttu-id="b9bb9-204">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b9bb9-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9bb9-205">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b9bb9-205">Requirements</span></span>

|<span data-ttu-id="b9bb9-206">Requisito</span><span class="sxs-lookup"><span data-stu-id="b9bb9-206">Requirement</span></span>| <span data-ttu-id="b9bb9-207">Valor</span><span class="sxs-lookup"><span data-stu-id="b9bb9-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9bb9-208">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b9bb9-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b9bb9-209">1.0</span><span class="sxs-lookup"><span data-stu-id="b9bb9-209">1.0</span></span>|
|[<span data-ttu-id="b9bb9-210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b9bb9-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b9bb9-211">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b9bb9-211">Compose or Read</span></span>|
