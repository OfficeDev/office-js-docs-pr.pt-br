---
title: Namespace do Office – conjunto de requisitos de pré-visualização
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: bbec602680da7914666daf33ed36c45751ae69c6
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068319"
---
# <a name="office"></a><span data-ttu-id="5325d-102">Office</span><span class="sxs-lookup"><span data-stu-id="5325d-102">Office</span></span>

<span data-ttu-id="5325d-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="5325d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5325d-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5325d-105">Requirements</span></span>

|<span data-ttu-id="5325d-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="5325d-106">Requirement</span></span>| <span data-ttu-id="5325d-107">Valor</span><span class="sxs-lookup"><span data-stu-id="5325d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="5325d-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5325d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5325d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="5325d-109">1.0</span></span>|
|[<span data-ttu-id="5325d-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5325d-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5325d-111">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5325d-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5325d-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="5325d-112">Members and methods</span></span>

| <span data-ttu-id="5325d-113">Membro</span><span class="sxs-lookup"><span data-stu-id="5325d-113">Member</span></span> | <span data-ttu-id="5325d-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="5325d-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5325d-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="5325d-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="5325d-116">Membro</span><span class="sxs-lookup"><span data-stu-id="5325d-116">Member</span></span> |
| [<span data-ttu-id="5325d-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="5325d-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="5325d-118">Membro</span><span class="sxs-lookup"><span data-stu-id="5325d-118">Member</span></span> |
| [<span data-ttu-id="5325d-119">EventType</span><span class="sxs-lookup"><span data-stu-id="5325d-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="5325d-120">Membro</span><span class="sxs-lookup"><span data-stu-id="5325d-120">Member</span></span> |
| [<span data-ttu-id="5325d-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="5325d-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="5325d-122">Membro</span><span class="sxs-lookup"><span data-stu-id="5325d-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="5325d-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="5325d-123">Namespaces</span></span>

<span data-ttu-id="5325d-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="5325d-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="5325d-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="5325d-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="5325d-126">Membros</span><span class="sxs-lookup"><span data-stu-id="5325d-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="5325d-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="5325d-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="5325d-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="5325d-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5325d-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="5325d-129">Type</span></span>

*   <span data-ttu-id="5325d-130">String</span><span class="sxs-lookup"><span data-stu-id="5325d-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5325d-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5325d-131">Properties:</span></span>

|<span data-ttu-id="5325d-132">Nome</span><span class="sxs-lookup"><span data-stu-id="5325d-132">Name</span></span>| <span data-ttu-id="5325d-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="5325d-133">Type</span></span>| <span data-ttu-id="5325d-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="5325d-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5325d-135">String</span><span class="sxs-lookup"><span data-stu-id="5325d-135">String</span></span>|<span data-ttu-id="5325d-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="5325d-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5325d-137">String</span><span class="sxs-lookup"><span data-stu-id="5325d-137">String</span></span>|<span data-ttu-id="5325d-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="5325d-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5325d-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5325d-139">Requirements</span></span>

|<span data-ttu-id="5325d-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="5325d-140">Requirement</span></span>| <span data-ttu-id="5325d-141">Valor</span><span class="sxs-lookup"><span data-stu-id="5325d-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="5325d-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5325d-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5325d-143">1.0</span><span class="sxs-lookup"><span data-stu-id="5325d-143">1.0</span></span>|
|[<span data-ttu-id="5325d-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5325d-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5325d-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5325d-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="5325d-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="5325d-146">CoercionType :String</span></span>

<span data-ttu-id="5325d-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="5325d-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5325d-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="5325d-148">Type</span></span>

*   <span data-ttu-id="5325d-149">String</span><span class="sxs-lookup"><span data-stu-id="5325d-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5325d-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5325d-150">Properties:</span></span>

|<span data-ttu-id="5325d-151">Nome</span><span class="sxs-lookup"><span data-stu-id="5325d-151">Name</span></span>| <span data-ttu-id="5325d-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="5325d-152">Type</span></span>| <span data-ttu-id="5325d-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="5325d-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5325d-154">String</span><span class="sxs-lookup"><span data-stu-id="5325d-154">String</span></span>|<span data-ttu-id="5325d-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="5325d-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5325d-156">String</span><span class="sxs-lookup"><span data-stu-id="5325d-156">String</span></span>|<span data-ttu-id="5325d-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="5325d-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5325d-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5325d-158">Requirements</span></span>

|<span data-ttu-id="5325d-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="5325d-159">Requirement</span></span>| <span data-ttu-id="5325d-160">Valor</span><span class="sxs-lookup"><span data-stu-id="5325d-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="5325d-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5325d-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5325d-162">1.0</span><span class="sxs-lookup"><span data-stu-id="5325d-162">1.0</span></span>|
|[<span data-ttu-id="5325d-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5325d-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5325d-164">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5325d-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="5325d-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="5325d-165">EventType :String</span></span>

<span data-ttu-id="5325d-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="5325d-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="5325d-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="5325d-167">Type</span></span>

*   <span data-ttu-id="5325d-168">String</span><span class="sxs-lookup"><span data-stu-id="5325d-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5325d-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5325d-169">Properties:</span></span>

| <span data-ttu-id="5325d-170">Nome</span><span class="sxs-lookup"><span data-stu-id="5325d-170">Name</span></span> | <span data-ttu-id="5325d-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="5325d-171">Type</span></span> | <span data-ttu-id="5325d-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="5325d-172">Description</span></span> | <span data-ttu-id="5325d-173">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="5325d-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="5325d-174">String</span><span class="sxs-lookup"><span data-stu-id="5325d-174">String</span></span> | <span data-ttu-id="5325d-175">A data ou hora da série ou do compromisso selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="5325d-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="5325d-176">1.7</span><span class="sxs-lookup"><span data-stu-id="5325d-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="5325d-177">String</span><span class="sxs-lookup"><span data-stu-id="5325d-177">String</span></span> | <span data-ttu-id="5325d-178">Um anexo foi adicionado a ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="5325d-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="5325d-179">Visualização</span><span class="sxs-lookup"><span data-stu-id="5325d-179">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="5325d-180">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5325d-180">String</span></span> | <span data-ttu-id="5325d-181">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="5325d-181">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="5325d-182">1,5</span><span class="sxs-lookup"><span data-stu-id="5325d-182">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="5325d-183">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5325d-183">String</span></span> | <span data-ttu-id="5325d-184">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="5325d-184">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="5325d-185">Visualização</span><span class="sxs-lookup"><span data-stu-id="5325d-185">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="5325d-186">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="5325d-186">String</span></span> | <span data-ttu-id="5325d-187">A lista de destinatários do item selecionado ou o local do compromisso foi alterado.</span><span class="sxs-lookup"><span data-stu-id="5325d-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="5325d-188">1.7</span><span class="sxs-lookup"><span data-stu-id="5325d-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="5325d-189">String</span><span class="sxs-lookup"><span data-stu-id="5325d-189">String</span></span> | <span data-ttu-id="5325d-190">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="5325d-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="5325d-191">1.7</span><span class="sxs-lookup"><span data-stu-id="5325d-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5325d-192">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5325d-192">Requirements</span></span>

|<span data-ttu-id="5325d-193">Requisito</span><span class="sxs-lookup"><span data-stu-id="5325d-193">Requirement</span></span>| <span data-ttu-id="5325d-194">Valor</span><span class="sxs-lookup"><span data-stu-id="5325d-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="5325d-195">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5325d-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5325d-196">1.5</span><span class="sxs-lookup"><span data-stu-id="5325d-196">1.5</span></span> |
|[<span data-ttu-id="5325d-197">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5325d-197">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5325d-198">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5325d-198">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="5325d-199">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="5325d-199">SourceProperty :String</span></span>

<span data-ttu-id="5325d-200">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="5325d-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5325d-201">Tipo</span><span class="sxs-lookup"><span data-stu-id="5325d-201">Type</span></span>

*   <span data-ttu-id="5325d-202">String</span><span class="sxs-lookup"><span data-stu-id="5325d-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5325d-203">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="5325d-203">Properties:</span></span>

|<span data-ttu-id="5325d-204">Nome</span><span class="sxs-lookup"><span data-stu-id="5325d-204">Name</span></span>| <span data-ttu-id="5325d-205">Tipo</span><span class="sxs-lookup"><span data-stu-id="5325d-205">Type</span></span>| <span data-ttu-id="5325d-206">Descrição</span><span class="sxs-lookup"><span data-stu-id="5325d-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5325d-207">String</span><span class="sxs-lookup"><span data-stu-id="5325d-207">String</span></span>|<span data-ttu-id="5325d-208">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5325d-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5325d-209">String</span><span class="sxs-lookup"><span data-stu-id="5325d-209">String</span></span>|<span data-ttu-id="5325d-210">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="5325d-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5325d-211">Requisitos</span><span class="sxs-lookup"><span data-stu-id="5325d-211">Requirements</span></span>

|<span data-ttu-id="5325d-212">Requisito</span><span class="sxs-lookup"><span data-stu-id="5325d-212">Requirement</span></span>| <span data-ttu-id="5325d-213">Valor</span><span class="sxs-lookup"><span data-stu-id="5325d-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="5325d-214">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="5325d-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5325d-215">1.0</span><span class="sxs-lookup"><span data-stu-id="5325d-215">1.0</span></span>|
|[<span data-ttu-id="5325d-216">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="5325d-216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="5325d-217">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="5325d-217">Compose or Read</span></span>|
