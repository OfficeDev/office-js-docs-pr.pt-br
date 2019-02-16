---
title: 'Namespace do Office: conjunto de requisitos da versão 1.7'
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: d6422e470864d5a02db37e1fef295e8cbb82a213
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067892"
---
# <a name="office"></a><span data-ttu-id="5df01-102">Office</span><span class="sxs-lookup"><span data-stu-id="5df01-102">Office</span></span>

<span data-ttu-id="5df01-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="5df01-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5df01-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5df01-105">Requirements</span></span>

|<span data-ttu-id="5df01-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="5df01-106">Requirement</span></span>| <span data-ttu-id="5df01-107">Valor</span><span class="sxs-lookup"><span data-stu-id="5df01-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="5df01-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5df01-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5df01-109">1.0</span><span class="sxs-lookup"><span data-stu-id="5df01-109">1.0</span></span>|
|[<span data-ttu-id="5df01-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5df01-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5df01-111">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5df01-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5df01-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="5df01-112">Members and methods</span></span>

| <span data-ttu-id="5df01-113">Membro</span><span class="sxs-lookup"><span data-stu-id="5df01-113">Member</span></span> | <span data-ttu-id="5df01-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="5df01-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5df01-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="5df01-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="5df01-116">Membro</span><span class="sxs-lookup"><span data-stu-id="5df01-116">Member</span></span> |
| [<span data-ttu-id="5df01-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="5df01-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="5df01-118">Membro</span><span class="sxs-lookup"><span data-stu-id="5df01-118">Member</span></span> |
| [<span data-ttu-id="5df01-119">EventType</span><span class="sxs-lookup"><span data-stu-id="5df01-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="5df01-120">Membro</span><span class="sxs-lookup"><span data-stu-id="5df01-120">Member</span></span> |
| [<span data-ttu-id="5df01-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="5df01-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="5df01-122">Membro</span><span class="sxs-lookup"><span data-stu-id="5df01-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="5df01-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="5df01-123">Namespaces</span></span>

<span data-ttu-id="5df01-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="5df01-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="5df01-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="5df01-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="5df01-126">Membros</span><span class="sxs-lookup"><span data-stu-id="5df01-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="5df01-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="5df01-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="5df01-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="5df01-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5df01-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="5df01-129">Type</span></span>

*   <span data-ttu-id="5df01-130">String</span><span class="sxs-lookup"><span data-stu-id="5df01-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5df01-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5df01-131">Properties:</span></span>

|<span data-ttu-id="5df01-132">Nome</span><span class="sxs-lookup"><span data-stu-id="5df01-132">Name</span></span>| <span data-ttu-id="5df01-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="5df01-133">Type</span></span>| <span data-ttu-id="5df01-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="5df01-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5df01-135">String</span><span class="sxs-lookup"><span data-stu-id="5df01-135">String</span></span>|<span data-ttu-id="5df01-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="5df01-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5df01-137">String</span><span class="sxs-lookup"><span data-stu-id="5df01-137">String</span></span>|<span data-ttu-id="5df01-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="5df01-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5df01-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5df01-139">Requirements</span></span>

|<span data-ttu-id="5df01-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="5df01-140">Requirement</span></span>| <span data-ttu-id="5df01-141">Valor</span><span class="sxs-lookup"><span data-stu-id="5df01-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="5df01-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5df01-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5df01-143">1.0</span><span class="sxs-lookup"><span data-stu-id="5df01-143">1.0</span></span>|
|[<span data-ttu-id="5df01-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5df01-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5df01-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5df01-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="5df01-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="5df01-146">CoercionType :String</span></span>

<span data-ttu-id="5df01-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="5df01-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5df01-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="5df01-148">Type</span></span>

*   <span data-ttu-id="5df01-149">String</span><span class="sxs-lookup"><span data-stu-id="5df01-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5df01-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5df01-150">Properties:</span></span>

|<span data-ttu-id="5df01-151">Nome</span><span class="sxs-lookup"><span data-stu-id="5df01-151">Name</span></span>| <span data-ttu-id="5df01-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="5df01-152">Type</span></span>| <span data-ttu-id="5df01-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="5df01-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5df01-154">String</span><span class="sxs-lookup"><span data-stu-id="5df01-154">String</span></span>|<span data-ttu-id="5df01-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="5df01-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5df01-156">String</span><span class="sxs-lookup"><span data-stu-id="5df01-156">String</span></span>|<span data-ttu-id="5df01-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="5df01-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5df01-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5df01-158">Requirements</span></span>

|<span data-ttu-id="5df01-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="5df01-159">Requirement</span></span>| <span data-ttu-id="5df01-160">Valor</span><span class="sxs-lookup"><span data-stu-id="5df01-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="5df01-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5df01-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5df01-162">1.0</span><span class="sxs-lookup"><span data-stu-id="5df01-162">1.0</span></span>|
|[<span data-ttu-id="5df01-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5df01-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5df01-164">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5df01-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="5df01-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="5df01-165">EventType :String</span></span>

<span data-ttu-id="5df01-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="5df01-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="5df01-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="5df01-167">Type</span></span>

*   <span data-ttu-id="5df01-168">String</span><span class="sxs-lookup"><span data-stu-id="5df01-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5df01-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5df01-169">Properties:</span></span>

| <span data-ttu-id="5df01-170">Nome</span><span class="sxs-lookup"><span data-stu-id="5df01-170">Name</span></span> | <span data-ttu-id="5df01-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="5df01-171">Type</span></span> | <span data-ttu-id="5df01-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="5df01-172">Description</span></span> | <span data-ttu-id="5df01-173">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="5df01-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="5df01-174">String</span><span class="sxs-lookup"><span data-stu-id="5df01-174">String</span></span> | <span data-ttu-id="5df01-175">A data ou hora da série ou do compromisso selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="5df01-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="5df01-176">1.7</span><span class="sxs-lookup"><span data-stu-id="5df01-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="5df01-177">String</span><span class="sxs-lookup"><span data-stu-id="5df01-177">String</span></span> | <span data-ttu-id="5df01-178">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="5df01-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="5df01-179">1,5</span><span class="sxs-lookup"><span data-stu-id="5df01-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="5df01-180">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5df01-180">String</span></span> | <span data-ttu-id="5df01-181">A lista de destinatários do item selecionado ou o local do compromisso foi alterado.</span><span class="sxs-lookup"><span data-stu-id="5df01-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="5df01-182">1.7</span><span class="sxs-lookup"><span data-stu-id="5df01-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="5df01-183">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5df01-183">String</span></span> | <span data-ttu-id="5df01-184">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="5df01-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="5df01-185">1.7</span><span class="sxs-lookup"><span data-stu-id="5df01-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5df01-186">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5df01-186">Requirements</span></span>

|<span data-ttu-id="5df01-187">Requisito</span><span class="sxs-lookup"><span data-stu-id="5df01-187">Requirement</span></span>| <span data-ttu-id="5df01-188">Valor</span><span class="sxs-lookup"><span data-stu-id="5df01-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="5df01-189">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5df01-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5df01-190">1.5</span><span class="sxs-lookup"><span data-stu-id="5df01-190">1.5</span></span> |
|[<span data-ttu-id="5df01-191">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5df01-191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5df01-192">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5df01-192">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="5df01-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="5df01-193">SourceProperty :String</span></span>

<span data-ttu-id="5df01-194">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="5df01-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5df01-195">Tipo</span><span class="sxs-lookup"><span data-stu-id="5df01-195">Type</span></span>

*   <span data-ttu-id="5df01-196">String</span><span class="sxs-lookup"><span data-stu-id="5df01-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5df01-197">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5df01-197">Properties:</span></span>

|<span data-ttu-id="5df01-198">Nome</span><span class="sxs-lookup"><span data-stu-id="5df01-198">Name</span></span>| <span data-ttu-id="5df01-199">Tipo</span><span class="sxs-lookup"><span data-stu-id="5df01-199">Type</span></span>| <span data-ttu-id="5df01-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="5df01-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5df01-201">String</span><span class="sxs-lookup"><span data-stu-id="5df01-201">String</span></span>|<span data-ttu-id="5df01-202">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5df01-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5df01-203">String</span><span class="sxs-lookup"><span data-stu-id="5df01-203">String</span></span>|<span data-ttu-id="5df01-204">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5df01-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5df01-205">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5df01-205">Requirements</span></span>

|<span data-ttu-id="5df01-206">Requisito</span><span class="sxs-lookup"><span data-stu-id="5df01-206">Requirement</span></span>| <span data-ttu-id="5df01-207">Valor</span><span class="sxs-lookup"><span data-stu-id="5df01-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="5df01-208">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5df01-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5df01-209">1.0</span><span class="sxs-lookup"><span data-stu-id="5df01-209">1.0</span></span>|
|[<span data-ttu-id="5df01-210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5df01-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5df01-211">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5df01-211">Compose or Read</span></span>|
