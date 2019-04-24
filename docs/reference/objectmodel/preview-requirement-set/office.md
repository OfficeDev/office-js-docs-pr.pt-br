---
title: Namespace do Office – conjunto de requisitos de visualização
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 7effc930d196aa009c3c779b702e082ae388fada
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451952"
---
# <a name="office"></a><span data-ttu-id="b45cc-102">Office</span><span class="sxs-lookup"><span data-stu-id="b45cc-102">Office</span></span>

<span data-ttu-id="b45cc-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="b45cc-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b45cc-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b45cc-105">Requirements</span></span>

|<span data-ttu-id="b45cc-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="b45cc-106">Requirement</span></span>| <span data-ttu-id="b45cc-107">Valor</span><span class="sxs-lookup"><span data-stu-id="b45cc-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b45cc-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b45cc-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b45cc-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b45cc-109">1.0</span></span>|
|[<span data-ttu-id="b45cc-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b45cc-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b45cc-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b45cc-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b45cc-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="b45cc-112">Members and methods</span></span>

| <span data-ttu-id="b45cc-113">Membro</span><span class="sxs-lookup"><span data-stu-id="b45cc-113">Member</span></span> | <span data-ttu-id="b45cc-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="b45cc-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b45cc-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b45cc-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b45cc-116">Member</span><span class="sxs-lookup"><span data-stu-id="b45cc-116">Member</span></span> |
| [<span data-ttu-id="b45cc-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b45cc-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b45cc-118">Member</span><span class="sxs-lookup"><span data-stu-id="b45cc-118">Member</span></span> |
| [<span data-ttu-id="b45cc-119">EventType</span><span class="sxs-lookup"><span data-stu-id="b45cc-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="b45cc-120">Member</span><span class="sxs-lookup"><span data-stu-id="b45cc-120">Member</span></span> |
| [<span data-ttu-id="b45cc-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b45cc-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b45cc-122">Membro</span><span class="sxs-lookup"><span data-stu-id="b45cc-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b45cc-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="b45cc-123">Namespaces</span></span>

<span data-ttu-id="b45cc-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="b45cc-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="b45cc-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="b45cc-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="b45cc-126">Membros</span><span class="sxs-lookup"><span data-stu-id="b45cc-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="b45cc-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="b45cc-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="b45cc-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="b45cc-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b45cc-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="b45cc-129">Type</span></span>

*   <span data-ttu-id="b45cc-130">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b45cc-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b45cc-131">Properties:</span></span>

|<span data-ttu-id="b45cc-132">Name</span><span class="sxs-lookup"><span data-stu-id="b45cc-132">Name</span></span>| <span data-ttu-id="b45cc-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="b45cc-133">Type</span></span>| <span data-ttu-id="b45cc-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="b45cc-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b45cc-135">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-135">String</span></span>|<span data-ttu-id="b45cc-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="b45cc-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b45cc-137">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-137">String</span></span>|<span data-ttu-id="b45cc-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="b45cc-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b45cc-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b45cc-139">Requirements</span></span>

|<span data-ttu-id="b45cc-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="b45cc-140">Requirement</span></span>| <span data-ttu-id="b45cc-141">Valor</span><span class="sxs-lookup"><span data-stu-id="b45cc-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="b45cc-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b45cc-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b45cc-143">1.0</span><span class="sxs-lookup"><span data-stu-id="b45cc-143">1.0</span></span>|
|[<span data-ttu-id="b45cc-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b45cc-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b45cc-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b45cc-145">Compose or Read</span></span>|

---
---

####  <a name="coerciontype-string"></a><span data-ttu-id="b45cc-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="b45cc-146">CoercionType :String</span></span>

<span data-ttu-id="b45cc-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="b45cc-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b45cc-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="b45cc-148">Type</span></span>

*   <span data-ttu-id="b45cc-149">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b45cc-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b45cc-150">Properties:</span></span>

|<span data-ttu-id="b45cc-151">Name</span><span class="sxs-lookup"><span data-stu-id="b45cc-151">Name</span></span>| <span data-ttu-id="b45cc-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="b45cc-152">Type</span></span>| <span data-ttu-id="b45cc-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="b45cc-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b45cc-154">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-154">String</span></span>|<span data-ttu-id="b45cc-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="b45cc-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b45cc-156">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-156">String</span></span>|<span data-ttu-id="b45cc-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="b45cc-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b45cc-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b45cc-158">Requirements</span></span>

|<span data-ttu-id="b45cc-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="b45cc-159">Requirement</span></span>| <span data-ttu-id="b45cc-160">Valor</span><span class="sxs-lookup"><span data-stu-id="b45cc-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="b45cc-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b45cc-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b45cc-162">1.0</span><span class="sxs-lookup"><span data-stu-id="b45cc-162">1.0</span></span>|
|[<span data-ttu-id="b45cc-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b45cc-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b45cc-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b45cc-164">Compose or Read</span></span>|

---
---

####  <a name="eventtype-string"></a><span data-ttu-id="b45cc-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="b45cc-165">EventType :String</span></span>

<span data-ttu-id="b45cc-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="b45cc-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="b45cc-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="b45cc-167">Type</span></span>

*   <span data-ttu-id="b45cc-168">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b45cc-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b45cc-169">Properties:</span></span>

| <span data-ttu-id="b45cc-170">Name</span><span class="sxs-lookup"><span data-stu-id="b45cc-170">Name</span></span> | <span data-ttu-id="b45cc-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="b45cc-171">Type</span></span> | <span data-ttu-id="b45cc-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="b45cc-172">Description</span></span> | <span data-ttu-id="b45cc-173">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="b45cc-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="b45cc-174">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-174">String</span></span> | <span data-ttu-id="b45cc-175">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="b45cc-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="b45cc-176">1.7</span><span class="sxs-lookup"><span data-stu-id="b45cc-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="b45cc-177">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-177">String</span></span> | <span data-ttu-id="b45cc-178">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="b45cc-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="b45cc-179">Visualização</span><span class="sxs-lookup"><span data-stu-id="b45cc-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="b45cc-180">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-180">String</span></span> | <span data-ttu-id="b45cc-181">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="b45cc-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="b45cc-182">Visualização</span><span class="sxs-lookup"><span data-stu-id="b45cc-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="b45cc-183">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-183">String</span></span> | <span data-ttu-id="b45cc-184">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="b45cc-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="b45cc-185">1,5</span><span class="sxs-lookup"><span data-stu-id="b45cc-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="b45cc-186">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-186">String</span></span> | <span data-ttu-id="b45cc-187">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="b45cc-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="b45cc-188">Visualização</span><span class="sxs-lookup"><span data-stu-id="b45cc-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="b45cc-189">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-189">String</span></span> | <span data-ttu-id="b45cc-190">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="b45cc-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="b45cc-191">1.7</span><span class="sxs-lookup"><span data-stu-id="b45cc-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="b45cc-192">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-192">String</span></span> | <span data-ttu-id="b45cc-193">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="b45cc-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="b45cc-194">1.7</span><span class="sxs-lookup"><span data-stu-id="b45cc-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b45cc-195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b45cc-195">Requirements</span></span>

|<span data-ttu-id="b45cc-196">Requisito</span><span class="sxs-lookup"><span data-stu-id="b45cc-196">Requirement</span></span>| <span data-ttu-id="b45cc-197">Valor</span><span class="sxs-lookup"><span data-stu-id="b45cc-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="b45cc-198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b45cc-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b45cc-199">1,5</span><span class="sxs-lookup"><span data-stu-id="b45cc-199">1.5</span></span> |
|[<span data-ttu-id="b45cc-200">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b45cc-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b45cc-201">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b45cc-201">Compose or Read</span></span> |

---
---

####  <a name="sourceproperty-string"></a><span data-ttu-id="b45cc-202">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="b45cc-202">SourceProperty :String</span></span>

<span data-ttu-id="b45cc-203">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="b45cc-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b45cc-204">Tipo</span><span class="sxs-lookup"><span data-stu-id="b45cc-204">Type</span></span>

*   <span data-ttu-id="b45cc-205">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b45cc-206">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="b45cc-206">Properties:</span></span>

|<span data-ttu-id="b45cc-207">Name</span><span class="sxs-lookup"><span data-stu-id="b45cc-207">Name</span></span>| <span data-ttu-id="b45cc-208">Tipo</span><span class="sxs-lookup"><span data-stu-id="b45cc-208">Type</span></span>| <span data-ttu-id="b45cc-209">Descrição</span><span class="sxs-lookup"><span data-stu-id="b45cc-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b45cc-210">String</span><span class="sxs-lookup"><span data-stu-id="b45cc-210">String</span></span>|<span data-ttu-id="b45cc-211">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b45cc-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b45cc-212">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="b45cc-212">String</span></span>|<span data-ttu-id="b45cc-213">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="b45cc-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b45cc-214">Requisitos</span><span class="sxs-lookup"><span data-stu-id="b45cc-214">Requirements</span></span>

|<span data-ttu-id="b45cc-215">Requisito</span><span class="sxs-lookup"><span data-stu-id="b45cc-215">Requirement</span></span>| <span data-ttu-id="b45cc-216">Valor</span><span class="sxs-lookup"><span data-stu-id="b45cc-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="b45cc-217">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="b45cc-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b45cc-218">1.0</span><span class="sxs-lookup"><span data-stu-id="b45cc-218">1.0</span></span>|
|[<span data-ttu-id="b45cc-219">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="b45cc-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b45cc-220">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="b45cc-220">Compose or Read</span></span>|
