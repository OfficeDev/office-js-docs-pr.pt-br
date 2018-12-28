---
title: Namespace do Office – conjunto de requisitos de pré-visualização
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: a276af19ebd1816ad6bd59af5a75c39f13aa0b3c
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432891"
---
# <a name="office"></a><span data-ttu-id="ba0d7-102">Office</span><span class="sxs-lookup"><span data-stu-id="ba0d7-102">Office</span></span>

<span data-ttu-id="ba0d7-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ba0d7-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ba0d7-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ba0d7-105">Requirements</span></span>

|<span data-ttu-id="ba0d7-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="ba0d7-106">Requirement</span></span>| <span data-ttu-id="ba0d7-107">Valor</span><span class="sxs-lookup"><span data-stu-id="ba0d7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba0d7-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ba0d7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba0d7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ba0d7-109">1.0</span></span>|
|[<span data-ttu-id="ba0d7-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ba0d7-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba0d7-111">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="ba0d7-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ba0d7-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="ba0d7-112">Members and methods</span></span>

| <span data-ttu-id="ba0d7-113">Membro</span><span class="sxs-lookup"><span data-stu-id="ba0d7-113">Member</span></span> | <span data-ttu-id="ba0d7-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="ba0d7-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ba0d7-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ba0d7-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ba0d7-116">Membro</span><span class="sxs-lookup"><span data-stu-id="ba0d7-116">Member</span></span> |
| [<span data-ttu-id="ba0d7-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ba0d7-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ba0d7-118">Membro</span><span class="sxs-lookup"><span data-stu-id="ba0d7-118">Member</span></span> |
| [<span data-ttu-id="ba0d7-119">EventType</span><span class="sxs-lookup"><span data-stu-id="ba0d7-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="ba0d7-120">Membro</span><span class="sxs-lookup"><span data-stu-id="ba0d7-120">Member</span></span> |
| [<span data-ttu-id="ba0d7-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ba0d7-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ba0d7-122">Membro</span><span class="sxs-lookup"><span data-stu-id="ba0d7-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ba0d7-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="ba0d7-123">Namespaces</span></span>

<span data-ttu-id="ba0d7-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="ba0d7-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="ba0d7-126">Membros</span><span class="sxs-lookup"><span data-stu-id="ba0d7-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="ba0d7-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="ba0d7-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="ba0d7-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ba0d7-129">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ba0d7-129">Type:</span></span>

*   <span data-ttu-id="ba0d7-130">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ba0d7-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ba0d7-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ba0d7-131">Properties:</span></span>

|<span data-ttu-id="ba0d7-132">Nome</span><span class="sxs-lookup"><span data-stu-id="ba0d7-132">Name</span></span>| <span data-ttu-id="ba0d7-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="ba0d7-133">Type</span></span>| <span data-ttu-id="ba0d7-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="ba0d7-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ba0d7-135">String</span><span class="sxs-lookup"><span data-stu-id="ba0d7-135">String</span></span>|<span data-ttu-id="ba0d7-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ba0d7-137">String</span><span class="sxs-lookup"><span data-stu-id="ba0d7-137">String</span></span>|<span data-ttu-id="ba0d7-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ba0d7-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ba0d7-139">Requirements</span></span>

|<span data-ttu-id="ba0d7-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="ba0d7-140">Requirement</span></span>| <span data-ttu-id="ba0d7-141">Valor</span><span class="sxs-lookup"><span data-stu-id="ba0d7-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba0d7-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ba0d7-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba0d7-143">1.0</span><span class="sxs-lookup"><span data-stu-id="ba0d7-143">1.0</span></span>|
|[<span data-ttu-id="ba0d7-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ba0d7-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba0d7-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ba0d7-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="ba0d7-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="ba0d7-146">CoercionType :String</span></span>

<span data-ttu-id="ba0d7-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ba0d7-148">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ba0d7-148">Type:</span></span>

*   <span data-ttu-id="ba0d7-149">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ba0d7-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ba0d7-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ba0d7-150">Properties:</span></span>

|<span data-ttu-id="ba0d7-151">Nome</span><span class="sxs-lookup"><span data-stu-id="ba0d7-151">Name</span></span>| <span data-ttu-id="ba0d7-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="ba0d7-152">Type</span></span>| <span data-ttu-id="ba0d7-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="ba0d7-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ba0d7-154">String</span><span class="sxs-lookup"><span data-stu-id="ba0d7-154">String</span></span>|<span data-ttu-id="ba0d7-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ba0d7-156">String</span><span class="sxs-lookup"><span data-stu-id="ba0d7-156">String</span></span>|<span data-ttu-id="ba0d7-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ba0d7-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ba0d7-158">Requirements</span></span>

|<span data-ttu-id="ba0d7-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="ba0d7-159">Requirement</span></span>| <span data-ttu-id="ba0d7-160">Valor</span><span class="sxs-lookup"><span data-stu-id="ba0d7-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba0d7-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ba0d7-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba0d7-162">1.0</span><span class="sxs-lookup"><span data-stu-id="ba0d7-162">1.0</span></span>|
|[<span data-ttu-id="ba0d7-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ba0d7-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba0d7-164">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ba0d7-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="ba0d7-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="ba0d7-165">EventType :String</span></span>

<span data-ttu-id="ba0d7-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="ba0d7-167">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ba0d7-167">Type:</span></span>

*   <span data-ttu-id="ba0d7-168">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ba0d7-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ba0d7-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ba0d7-169">Properties:</span></span>

| <span data-ttu-id="ba0d7-170">Nome</span><span class="sxs-lookup"><span data-stu-id="ba0d7-170">Name</span></span> | <span data-ttu-id="ba0d7-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="ba0d7-171">Type</span></span> | <span data-ttu-id="ba0d7-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="ba0d7-172">Description</span></span> | <span data-ttu-id="ba0d7-173">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="ba0d7-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="ba0d7-174">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ba0d7-174">String</span></span> | <span data-ttu-id="ba0d7-175">A data ou hora da série ou do compromisso selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="ba0d7-176">1.7</span><span class="sxs-lookup"><span data-stu-id="ba0d7-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="ba0d7-177">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ba0d7-177">String</span></span> | <span data-ttu-id="ba0d7-178">Um anexo foi adicionado a ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="ba0d7-179">Visualização</span><span class="sxs-lookup"><span data-stu-id="ba0d7-179">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="ba0d7-180">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ba0d7-180">String</span></span> | <span data-ttu-id="ba0d7-181">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-181">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="ba0d7-182">1.5</span><span class="sxs-lookup"><span data-stu-id="ba0d7-182">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="ba0d7-183">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ba0d7-183">String</span></span> | <span data-ttu-id="ba0d7-184">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-184">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="ba0d7-185">Visualização</span><span class="sxs-lookup"><span data-stu-id="ba0d7-185">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="ba0d7-186">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ba0d7-186">String</span></span> | <span data-ttu-id="ba0d7-187">A lista de destinatários do item selecionado ou o local do compromisso foi alterado.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="ba0d7-188">1.7</span><span class="sxs-lookup"><span data-stu-id="ba0d7-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="ba0d7-189">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ba0d7-189">String</span></span> | <span data-ttu-id="ba0d7-190">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="ba0d7-191">1.7</span><span class="sxs-lookup"><span data-stu-id="ba0d7-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ba0d7-192">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ba0d7-192">Requirements</span></span>

|<span data-ttu-id="ba0d7-193">Requisito</span><span class="sxs-lookup"><span data-stu-id="ba0d7-193">Requirement</span></span>| <span data-ttu-id="ba0d7-194">Valor</span><span class="sxs-lookup"><span data-stu-id="ba0d7-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba0d7-195">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ba0d7-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba0d7-196">1.5</span><span class="sxs-lookup"><span data-stu-id="ba0d7-196">1.5</span></span> |
|[<span data-ttu-id="ba0d7-197">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ba0d7-197">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba0d7-198">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ba0d7-198">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="ba0d7-199">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="ba0d7-199">SourceProperty :String</span></span>

<span data-ttu-id="ba0d7-200">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ba0d7-201">Tipo:</span><span class="sxs-lookup"><span data-stu-id="ba0d7-201">Type:</span></span>

*   <span data-ttu-id="ba0d7-202">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="ba0d7-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ba0d7-203">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="ba0d7-203">Properties:</span></span>

|<span data-ttu-id="ba0d7-204">Nome</span><span class="sxs-lookup"><span data-stu-id="ba0d7-204">Name</span></span>| <span data-ttu-id="ba0d7-205">Tipo</span><span class="sxs-lookup"><span data-stu-id="ba0d7-205">Type</span></span>| <span data-ttu-id="ba0d7-206">Descrição</span><span class="sxs-lookup"><span data-stu-id="ba0d7-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ba0d7-207">String</span><span class="sxs-lookup"><span data-stu-id="ba0d7-207">String</span></span>|<span data-ttu-id="ba0d7-208">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ba0d7-209">String</span><span class="sxs-lookup"><span data-stu-id="ba0d7-209">String</span></span>|<span data-ttu-id="ba0d7-210">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="ba0d7-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ba0d7-211">Requisitos</span><span class="sxs-lookup"><span data-stu-id="ba0d7-211">Requirements</span></span>

|<span data-ttu-id="ba0d7-212">Requisito</span><span class="sxs-lookup"><span data-stu-id="ba0d7-212">Requirement</span></span>| <span data-ttu-id="ba0d7-213">Valor</span><span class="sxs-lookup"><span data-stu-id="ba0d7-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="ba0d7-214">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="ba0d7-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ba0d7-215">1.0</span><span class="sxs-lookup"><span data-stu-id="ba0d7-215">1.0</span></span>|
|[<span data-ttu-id="ba0d7-216">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="ba0d7-216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ba0d7-217">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="ba0d7-217">Compose or read</span></span>|