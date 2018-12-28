---
title: 'Namespace do Office: conjunto de requisitos da versão 1.7'
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 2bf1c31f4dc4156cb4f1d0eb3508193305c860e9
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432799"
---
# <a name="office"></a><span data-ttu-id="afebf-102">Office</span><span class="sxs-lookup"><span data-stu-id="afebf-102">Office</span></span>

<span data-ttu-id="afebf-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="afebf-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="afebf-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afebf-105">Requirements</span></span>

|<span data-ttu-id="afebf-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="afebf-106">Requirement</span></span>| <span data-ttu-id="afebf-107">Valor</span><span class="sxs-lookup"><span data-stu-id="afebf-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="afebf-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afebf-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afebf-109">1.0</span><span class="sxs-lookup"><span data-stu-id="afebf-109">1.0</span></span>|
|[<span data-ttu-id="afebf-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afebf-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="afebf-111">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="afebf-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="afebf-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="afebf-112">Members and methods</span></span>

| <span data-ttu-id="afebf-113">Membro</span><span class="sxs-lookup"><span data-stu-id="afebf-113">Member</span></span> | <span data-ttu-id="afebf-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="afebf-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="afebf-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="afebf-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="afebf-116">Membro</span><span class="sxs-lookup"><span data-stu-id="afebf-116">Member</span></span> |
| [<span data-ttu-id="afebf-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="afebf-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="afebf-118">Membro</span><span class="sxs-lookup"><span data-stu-id="afebf-118">Member</span></span> |
| [<span data-ttu-id="afebf-119">EventType</span><span class="sxs-lookup"><span data-stu-id="afebf-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="afebf-120">Membro</span><span class="sxs-lookup"><span data-stu-id="afebf-120">Member</span></span> |
| [<span data-ttu-id="afebf-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="afebf-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="afebf-122">Membro</span><span class="sxs-lookup"><span data-stu-id="afebf-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="afebf-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="afebf-123">Namespaces</span></span>

<span data-ttu-id="afebf-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="afebf-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="afebf-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="afebf-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="afebf-126">Membros</span><span class="sxs-lookup"><span data-stu-id="afebf-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="afebf-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="afebf-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="afebf-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="afebf-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="afebf-129">Tipo:</span><span class="sxs-lookup"><span data-stu-id="afebf-129">Type:</span></span>

*   <span data-ttu-id="afebf-130">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="afebf-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="afebf-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="afebf-131">Properties:</span></span>

|<span data-ttu-id="afebf-132">Nome</span><span class="sxs-lookup"><span data-stu-id="afebf-132">Name</span></span>| <span data-ttu-id="afebf-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="afebf-133">Type</span></span>| <span data-ttu-id="afebf-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="afebf-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="afebf-135">String</span><span class="sxs-lookup"><span data-stu-id="afebf-135">String</span></span>|<span data-ttu-id="afebf-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="afebf-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="afebf-137">String</span><span class="sxs-lookup"><span data-stu-id="afebf-137">String</span></span>|<span data-ttu-id="afebf-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="afebf-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="afebf-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afebf-139">Requirements</span></span>

|<span data-ttu-id="afebf-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="afebf-140">Requirement</span></span>| <span data-ttu-id="afebf-141">Valor</span><span class="sxs-lookup"><span data-stu-id="afebf-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="afebf-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afebf-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afebf-143">1.0</span><span class="sxs-lookup"><span data-stu-id="afebf-143">1.0</span></span>|
|[<span data-ttu-id="afebf-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afebf-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="afebf-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="afebf-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="afebf-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="afebf-146">CoercionType :String</span></span>

<span data-ttu-id="afebf-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="afebf-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="afebf-148">Tipo:</span><span class="sxs-lookup"><span data-stu-id="afebf-148">Type:</span></span>

*   <span data-ttu-id="afebf-149">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="afebf-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="afebf-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="afebf-150">Properties:</span></span>

|<span data-ttu-id="afebf-151">Nome</span><span class="sxs-lookup"><span data-stu-id="afebf-151">Name</span></span>| <span data-ttu-id="afebf-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="afebf-152">Type</span></span>| <span data-ttu-id="afebf-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="afebf-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="afebf-154">String</span><span class="sxs-lookup"><span data-stu-id="afebf-154">String</span></span>|<span data-ttu-id="afebf-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="afebf-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="afebf-156">String</span><span class="sxs-lookup"><span data-stu-id="afebf-156">String</span></span>|<span data-ttu-id="afebf-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="afebf-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="afebf-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afebf-158">Requirements</span></span>

|<span data-ttu-id="afebf-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="afebf-159">Requirement</span></span>| <span data-ttu-id="afebf-160">Valor</span><span class="sxs-lookup"><span data-stu-id="afebf-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="afebf-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afebf-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afebf-162">1.0</span><span class="sxs-lookup"><span data-stu-id="afebf-162">1.0</span></span>|
|[<span data-ttu-id="afebf-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afebf-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="afebf-164">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="afebf-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="afebf-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="afebf-165">EventType :String</span></span>

<span data-ttu-id="afebf-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="afebf-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="afebf-167">Tipo:</span><span class="sxs-lookup"><span data-stu-id="afebf-167">Type:</span></span>

*   <span data-ttu-id="afebf-168">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="afebf-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="afebf-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="afebf-169">Properties:</span></span>

| <span data-ttu-id="afebf-170">Nome</span><span class="sxs-lookup"><span data-stu-id="afebf-170">Name</span></span> | <span data-ttu-id="afebf-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="afebf-171">Type</span></span> | <span data-ttu-id="afebf-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="afebf-172">Description</span></span> | <span data-ttu-id="afebf-173">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="afebf-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="afebf-174">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="afebf-174">String</span></span> | <span data-ttu-id="afebf-175">A data ou hora da série ou do compromisso selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="afebf-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="afebf-176">1.7</span><span class="sxs-lookup"><span data-stu-id="afebf-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="afebf-177">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="afebf-177">String</span></span> | <span data-ttu-id="afebf-178">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="afebf-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="afebf-179">1.5</span><span class="sxs-lookup"><span data-stu-id="afebf-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="afebf-180">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="afebf-180">String</span></span> | <span data-ttu-id="afebf-181">A lista de destinatários do item selecionado ou o local do compromisso foi alterado.</span><span class="sxs-lookup"><span data-stu-id="afebf-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="afebf-182">1.7</span><span class="sxs-lookup"><span data-stu-id="afebf-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="afebf-183">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="afebf-183">String</span></span> | <span data-ttu-id="afebf-184">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="afebf-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="afebf-185">1.7</span><span class="sxs-lookup"><span data-stu-id="afebf-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="afebf-186">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afebf-186">Requirements</span></span>

|<span data-ttu-id="afebf-187">Requisito</span><span class="sxs-lookup"><span data-stu-id="afebf-187">Requirement</span></span>| <span data-ttu-id="afebf-188">Valor</span><span class="sxs-lookup"><span data-stu-id="afebf-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="afebf-189">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afebf-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afebf-190">1.5</span><span class="sxs-lookup"><span data-stu-id="afebf-190">1.5</span></span> |
|[<span data-ttu-id="afebf-191">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afebf-191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="afebf-192">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="afebf-192">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="afebf-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="afebf-193">SourceProperty :String</span></span>

<span data-ttu-id="afebf-194">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="afebf-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="afebf-195">Tipo:</span><span class="sxs-lookup"><span data-stu-id="afebf-195">Type:</span></span>

*   <span data-ttu-id="afebf-196">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="afebf-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="afebf-197">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="afebf-197">Properties:</span></span>

|<span data-ttu-id="afebf-198">Nome</span><span class="sxs-lookup"><span data-stu-id="afebf-198">Name</span></span>| <span data-ttu-id="afebf-199">Tipo</span><span class="sxs-lookup"><span data-stu-id="afebf-199">Type</span></span>| <span data-ttu-id="afebf-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="afebf-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="afebf-201">String</span><span class="sxs-lookup"><span data-stu-id="afebf-201">String</span></span>|<span data-ttu-id="afebf-202">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="afebf-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="afebf-203">String</span><span class="sxs-lookup"><span data-stu-id="afebf-203">String</span></span>|<span data-ttu-id="afebf-204">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="afebf-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="afebf-205">Requisitos</span><span class="sxs-lookup"><span data-stu-id="afebf-205">Requirements</span></span>

|<span data-ttu-id="afebf-206">Requisito</span><span class="sxs-lookup"><span data-stu-id="afebf-206">Requirement</span></span>| <span data-ttu-id="afebf-207">Valor</span><span class="sxs-lookup"><span data-stu-id="afebf-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="afebf-208">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="afebf-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="afebf-209">1.0</span><span class="sxs-lookup"><span data-stu-id="afebf-209">1.0</span></span>|
|[<span data-ttu-id="afebf-210">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="afebf-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="afebf-211">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="afebf-211">Compose or read</span></span>|