---
title: Namespace do Office – conjunto de requisitos de pré-visualização
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: f4a4f0d7a4ce0de433d4e70b6a4675b5f63f26f0
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457926"
---
# <a name="office"></a><span data-ttu-id="7c14f-102">Office</span><span class="sxs-lookup"><span data-stu-id="7c14f-102">Office</span></span>

<span data-ttu-id="7c14f-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="7c14f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c14f-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c14f-105">Requirements</span></span>

|<span data-ttu-id="7c14f-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c14f-106">Requirement</span></span>| <span data-ttu-id="7c14f-107">Valor</span><span class="sxs-lookup"><span data-stu-id="7c14f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c14f-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c14f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c14f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="7c14f-109">1.0</span></span>|
|[<span data-ttu-id="7c14f-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c14f-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7c14f-111">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="7c14f-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7c14f-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="7c14f-112">Members and methods</span></span>

| <span data-ttu-id="7c14f-113">Membro</span><span class="sxs-lookup"><span data-stu-id="7c14f-113">Member</span></span> | <span data-ttu-id="7c14f-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c14f-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7c14f-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="7c14f-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="7c14f-116">Membro</span><span class="sxs-lookup"><span data-stu-id="7c14f-116">Member</span></span> |
| [<span data-ttu-id="7c14f-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="7c14f-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="7c14f-118">Membro</span><span class="sxs-lookup"><span data-stu-id="7c14f-118">Member</span></span> |
| [<span data-ttu-id="7c14f-119">EventType</span><span class="sxs-lookup"><span data-stu-id="7c14f-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="7c14f-120">Membro</span><span class="sxs-lookup"><span data-stu-id="7c14f-120">Member</span></span> |
| [<span data-ttu-id="7c14f-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="7c14f-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="7c14f-122">Membro</span><span class="sxs-lookup"><span data-stu-id="7c14f-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="7c14f-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="7c14f-123">Namespaces</span></span>

<span data-ttu-id="7c14f-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="7c14f-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="7c14f-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="7c14f-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="7c14f-126">Membros</span><span class="sxs-lookup"><span data-stu-id="7c14f-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="7c14f-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="7c14f-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="7c14f-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="7c14f-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="7c14f-129">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7c14f-129">Type:</span></span>

*   <span data-ttu-id="7c14f-130">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c14f-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7c14f-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7c14f-131">Properties:</span></span>

|<span data-ttu-id="7c14f-132">Nome</span><span class="sxs-lookup"><span data-stu-id="7c14f-132">Name</span></span>| <span data-ttu-id="7c14f-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c14f-133">Type</span></span>| <span data-ttu-id="7c14f-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c14f-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="7c14f-135">String</span><span class="sxs-lookup"><span data-stu-id="7c14f-135">String</span></span>|<span data-ttu-id="7c14f-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="7c14f-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="7c14f-137">String</span><span class="sxs-lookup"><span data-stu-id="7c14f-137">String</span></span>|<span data-ttu-id="7c14f-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="7c14f-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7c14f-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c14f-139">Requirements</span></span>

|<span data-ttu-id="7c14f-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c14f-140">Requirement</span></span>| <span data-ttu-id="7c14f-141">Valor</span><span class="sxs-lookup"><span data-stu-id="7c14f-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c14f-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c14f-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c14f-143">1.0</span><span class="sxs-lookup"><span data-stu-id="7c14f-143">1.0</span></span>|
|[<span data-ttu-id="7c14f-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c14f-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7c14f-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="7c14f-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="7c14f-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="7c14f-146">CoercionType :String</span></span>

<span data-ttu-id="7c14f-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="7c14f-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7c14f-148">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7c14f-148">Type:</span></span>

*   <span data-ttu-id="7c14f-149">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c14f-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7c14f-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7c14f-150">Properties:</span></span>

|<span data-ttu-id="7c14f-151">Nome</span><span class="sxs-lookup"><span data-stu-id="7c14f-151">Name</span></span>| <span data-ttu-id="7c14f-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c14f-152">Type</span></span>| <span data-ttu-id="7c14f-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c14f-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="7c14f-154">String</span><span class="sxs-lookup"><span data-stu-id="7c14f-154">String</span></span>|<span data-ttu-id="7c14f-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="7c14f-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="7c14f-156">String</span><span class="sxs-lookup"><span data-stu-id="7c14f-156">String</span></span>|<span data-ttu-id="7c14f-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="7c14f-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7c14f-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c14f-158">Requirements</span></span>

|<span data-ttu-id="7c14f-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c14f-159">Requirement</span></span>| <span data-ttu-id="7c14f-160">Valor</span><span class="sxs-lookup"><span data-stu-id="7c14f-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c14f-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c14f-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c14f-162">1.0</span><span class="sxs-lookup"><span data-stu-id="7c14f-162">1.0</span></span>|
|[<span data-ttu-id="7c14f-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c14f-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7c14f-164">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="7c14f-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="7c14f-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="7c14f-165">EventType :String</span></span>

<span data-ttu-id="7c14f-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="7c14f-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="7c14f-167">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7c14f-167">Type:</span></span>

*   <span data-ttu-id="7c14f-168">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c14f-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7c14f-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7c14f-169">Properties:</span></span>

| <span data-ttu-id="7c14f-170">Nome</span><span class="sxs-lookup"><span data-stu-id="7c14f-170">Name</span></span> | <span data-ttu-id="7c14f-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c14f-171">Type</span></span> | <span data-ttu-id="7c14f-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c14f-172">Description</span></span> | <span data-ttu-id="7c14f-173">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="7c14f-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="7c14f-174">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c14f-174">String</span></span> | <span data-ttu-id="7c14f-175">A data ou hora da série ou do compromisso selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="7c14f-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="7c14f-176">1.7</span><span class="sxs-lookup"><span data-stu-id="7c14f-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="7c14f-177">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c14f-177">String</span></span> | <span data-ttu-id="7c14f-178">Um anexo foi adicionado a ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="7c14f-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="7c14f-179">Visualização</span><span class="sxs-lookup"><span data-stu-id="7c14f-179">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="7c14f-180">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c14f-180">String</span></span> | <span data-ttu-id="7c14f-181">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="7c14f-181">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="7c14f-182">1.5</span><span class="sxs-lookup"><span data-stu-id="7c14f-182">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="7c14f-183">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c14f-183">String</span></span> | <span data-ttu-id="7c14f-184">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="7c14f-184">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="7c14f-185">Visualização</span><span class="sxs-lookup"><span data-stu-id="7c14f-185">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="7c14f-186">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c14f-186">String</span></span> | <span data-ttu-id="7c14f-187">A lista de destinatários do item selecionado ou o local do compromisso foi alterado.</span><span class="sxs-lookup"><span data-stu-id="7c14f-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="7c14f-188">1.7</span><span class="sxs-lookup"><span data-stu-id="7c14f-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="7c14f-189">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c14f-189">String</span></span> | <span data-ttu-id="7c14f-190">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="7c14f-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="7c14f-191">1.7</span><span class="sxs-lookup"><span data-stu-id="7c14f-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7c14f-192">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c14f-192">Requirements</span></span>

|<span data-ttu-id="7c14f-193">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c14f-193">Requirement</span></span>| <span data-ttu-id="7c14f-194">Valor</span><span class="sxs-lookup"><span data-stu-id="7c14f-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c14f-195">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c14f-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c14f-196">1.5</span><span class="sxs-lookup"><span data-stu-id="7c14f-196">1.5</span></span> |
|[<span data-ttu-id="7c14f-197">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c14f-197">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7c14f-198">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="7c14f-198">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="7c14f-199">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="7c14f-199">SourceProperty :String</span></span>

<span data-ttu-id="7c14f-200">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="7c14f-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7c14f-201">Tipo:</span><span class="sxs-lookup"><span data-stu-id="7c14f-201">Type:</span></span>

*   <span data-ttu-id="7c14f-202">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="7c14f-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7c14f-203">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="7c14f-203">Properties:</span></span>

|<span data-ttu-id="7c14f-204">Nome</span><span class="sxs-lookup"><span data-stu-id="7c14f-204">Name</span></span>| <span data-ttu-id="7c14f-205">Tipo</span><span class="sxs-lookup"><span data-stu-id="7c14f-205">Type</span></span>| <span data-ttu-id="7c14f-206">Descrição</span><span class="sxs-lookup"><span data-stu-id="7c14f-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="7c14f-207">String</span><span class="sxs-lookup"><span data-stu-id="7c14f-207">String</span></span>|<span data-ttu-id="7c14f-208">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7c14f-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="7c14f-209">String</span><span class="sxs-lookup"><span data-stu-id="7c14f-209">String</span></span>|<span data-ttu-id="7c14f-210">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="7c14f-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7c14f-211">Requisitos</span><span class="sxs-lookup"><span data-stu-id="7c14f-211">Requirements</span></span>

|<span data-ttu-id="7c14f-212">Requisito</span><span class="sxs-lookup"><span data-stu-id="7c14f-212">Requirement</span></span>| <span data-ttu-id="7c14f-213">Valor</span><span class="sxs-lookup"><span data-stu-id="7c14f-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c14f-214">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="7c14f-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c14f-215">1.0</span><span class="sxs-lookup"><span data-stu-id="7c14f-215">1.0</span></span>|
|[<span data-ttu-id="7c14f-216">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="7c14f-216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7c14f-217">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="7c14f-217">Compose or read</span></span>|