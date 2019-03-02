---
title: Namespace do Office – conjunto de requisitos de pré-visualização
description: ''
ms.date: 02/26/2019
localization_priority: Normal
ms.openlocfilehash: 7b27963a85f1dcdaa6f269fce242c45bf1bdd146
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359230"
---
# <a name="office"></a><span data-ttu-id="88e2c-102">Office</span><span class="sxs-lookup"><span data-stu-id="88e2c-102">Office</span></span>

<span data-ttu-id="88e2c-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="88e2c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="88e2c-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="88e2c-105">Requirements</span></span>

|<span data-ttu-id="88e2c-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="88e2c-106">Requirement</span></span>| <span data-ttu-id="88e2c-107">Valor</span><span class="sxs-lookup"><span data-stu-id="88e2c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e2c-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="88e2c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e2c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="88e2c-109">1.0</span></span>|
|[<span data-ttu-id="88e2c-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="88e2c-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e2c-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="88e2c-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="88e2c-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="88e2c-112">Members and methods</span></span>

| <span data-ttu-id="88e2c-113">Membro</span><span class="sxs-lookup"><span data-stu-id="88e2c-113">Member</span></span> | <span data-ttu-id="88e2c-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="88e2c-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="88e2c-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="88e2c-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="88e2c-116">Membro</span><span class="sxs-lookup"><span data-stu-id="88e2c-116">Member</span></span> |
| [<span data-ttu-id="88e2c-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="88e2c-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="88e2c-118">Membro</span><span class="sxs-lookup"><span data-stu-id="88e2c-118">Member</span></span> |
| [<span data-ttu-id="88e2c-119">EventType</span><span class="sxs-lookup"><span data-stu-id="88e2c-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="88e2c-120">Membro</span><span class="sxs-lookup"><span data-stu-id="88e2c-120">Member</span></span> |
| [<span data-ttu-id="88e2c-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="88e2c-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="88e2c-122">Membro</span><span class="sxs-lookup"><span data-stu-id="88e2c-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="88e2c-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="88e2c-123">Namespaces</span></span>

<span data-ttu-id="88e2c-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="88e2c-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="88e2c-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="88e2c-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="88e2c-126">Membros</span><span class="sxs-lookup"><span data-stu-id="88e2c-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="88e2c-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="88e2c-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="88e2c-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="88e2c-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="88e2c-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="88e2c-129">Type</span></span>

*   <span data-ttu-id="88e2c-130">String</span><span class="sxs-lookup"><span data-stu-id="88e2c-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="88e2c-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="88e2c-131">Properties:</span></span>

|<span data-ttu-id="88e2c-132">Nome</span><span class="sxs-lookup"><span data-stu-id="88e2c-132">Name</span></span>| <span data-ttu-id="88e2c-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="88e2c-133">Type</span></span>| <span data-ttu-id="88e2c-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="88e2c-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="88e2c-135">String</span><span class="sxs-lookup"><span data-stu-id="88e2c-135">String</span></span>|<span data-ttu-id="88e2c-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="88e2c-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="88e2c-137">String</span><span class="sxs-lookup"><span data-stu-id="88e2c-137">String</span></span>|<span data-ttu-id="88e2c-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="88e2c-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88e2c-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="88e2c-139">Requirements</span></span>

|<span data-ttu-id="88e2c-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="88e2c-140">Requirement</span></span>| <span data-ttu-id="88e2c-141">Valor</span><span class="sxs-lookup"><span data-stu-id="88e2c-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e2c-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="88e2c-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e2c-143">1.0</span><span class="sxs-lookup"><span data-stu-id="88e2c-143">1.0</span></span>|
|[<span data-ttu-id="88e2c-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="88e2c-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e2c-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="88e2c-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="88e2c-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="88e2c-146">CoercionType :String</span></span>

<span data-ttu-id="88e2c-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="88e2c-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="88e2c-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="88e2c-148">Type</span></span>

*   <span data-ttu-id="88e2c-149">String</span><span class="sxs-lookup"><span data-stu-id="88e2c-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="88e2c-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="88e2c-150">Properties:</span></span>

|<span data-ttu-id="88e2c-151">Nome</span><span class="sxs-lookup"><span data-stu-id="88e2c-151">Name</span></span>| <span data-ttu-id="88e2c-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="88e2c-152">Type</span></span>| <span data-ttu-id="88e2c-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="88e2c-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="88e2c-154">String</span><span class="sxs-lookup"><span data-stu-id="88e2c-154">String</span></span>|<span data-ttu-id="88e2c-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="88e2c-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="88e2c-156">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="88e2c-156">String</span></span>|<span data-ttu-id="88e2c-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="88e2c-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88e2c-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="88e2c-158">Requirements</span></span>

|<span data-ttu-id="88e2c-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="88e2c-159">Requirement</span></span>| <span data-ttu-id="88e2c-160">Valor</span><span class="sxs-lookup"><span data-stu-id="88e2c-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e2c-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="88e2c-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e2c-162">1.0</span><span class="sxs-lookup"><span data-stu-id="88e2c-162">1.0</span></span>|
|[<span data-ttu-id="88e2c-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="88e2c-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e2c-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="88e2c-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="88e2c-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="88e2c-165">EventType :String</span></span>

<span data-ttu-id="88e2c-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="88e2c-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="88e2c-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="88e2c-167">Type</span></span>

*   <span data-ttu-id="88e2c-168">String</span><span class="sxs-lookup"><span data-stu-id="88e2c-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="88e2c-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="88e2c-169">Properties:</span></span>

| <span data-ttu-id="88e2c-170">Nome</span><span class="sxs-lookup"><span data-stu-id="88e2c-170">Name</span></span> | <span data-ttu-id="88e2c-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="88e2c-171">Type</span></span> | <span data-ttu-id="88e2c-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="88e2c-172">Description</span></span> | <span data-ttu-id="88e2c-173">Conjunto de requisitos mínimos</span><span class="sxs-lookup"><span data-stu-id="88e2c-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="88e2c-174">String</span><span class="sxs-lookup"><span data-stu-id="88e2c-174">String</span></span> | <span data-ttu-id="88e2c-175">A data ou hora da série ou do compromisso selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="88e2c-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="88e2c-176">1.7</span><span class="sxs-lookup"><span data-stu-id="88e2c-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="88e2c-177">String</span><span class="sxs-lookup"><span data-stu-id="88e2c-177">String</span></span> | <span data-ttu-id="88e2c-178">Um anexo foi adicionado a ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="88e2c-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="88e2c-179">Visualização</span><span class="sxs-lookup"><span data-stu-id="88e2c-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="88e2c-180">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="88e2c-180">String</span></span> | <span data-ttu-id="88e2c-181">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="88e2c-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="88e2c-182">Visualização</span><span class="sxs-lookup"><span data-stu-id="88e2c-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="88e2c-183">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="88e2c-183">String</span></span> | <span data-ttu-id="88e2c-184">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="88e2c-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="88e2c-185">1,5</span><span class="sxs-lookup"><span data-stu-id="88e2c-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="88e2c-186">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="88e2c-186">String</span></span> | <span data-ttu-id="88e2c-187">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="88e2c-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="88e2c-188">Visualização</span><span class="sxs-lookup"><span data-stu-id="88e2c-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="88e2c-189">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="88e2c-189">String</span></span> | <span data-ttu-id="88e2c-190">A lista de destinatários do item selecionado ou o local do compromisso foi alterado.</span><span class="sxs-lookup"><span data-stu-id="88e2c-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="88e2c-191">1.7</span><span class="sxs-lookup"><span data-stu-id="88e2c-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="88e2c-192">String</span><span class="sxs-lookup"><span data-stu-id="88e2c-192">String</span></span> | <span data-ttu-id="88e2c-193">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="88e2c-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="88e2c-194">1.7</span><span class="sxs-lookup"><span data-stu-id="88e2c-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="88e2c-195">Requisitos</span><span class="sxs-lookup"><span data-stu-id="88e2c-195">Requirements</span></span>

|<span data-ttu-id="88e2c-196">Requisito</span><span class="sxs-lookup"><span data-stu-id="88e2c-196">Requirement</span></span>| <span data-ttu-id="88e2c-197">Valor</span><span class="sxs-lookup"><span data-stu-id="88e2c-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e2c-198">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="88e2c-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e2c-199">1.5</span><span class="sxs-lookup"><span data-stu-id="88e2c-199">1.5</span></span> |
|[<span data-ttu-id="88e2c-200">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="88e2c-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e2c-201">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="88e2c-201">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="88e2c-202">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="88e2c-202">SourceProperty :String</span></span>

<span data-ttu-id="88e2c-203">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="88e2c-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="88e2c-204">Tipo</span><span class="sxs-lookup"><span data-stu-id="88e2c-204">Type</span></span>

*   <span data-ttu-id="88e2c-205">String</span><span class="sxs-lookup"><span data-stu-id="88e2c-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="88e2c-206">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="88e2c-206">Properties:</span></span>

|<span data-ttu-id="88e2c-207">Nome</span><span class="sxs-lookup"><span data-stu-id="88e2c-207">Name</span></span>| <span data-ttu-id="88e2c-208">Tipo</span><span class="sxs-lookup"><span data-stu-id="88e2c-208">Type</span></span>| <span data-ttu-id="88e2c-209">Descrição</span><span class="sxs-lookup"><span data-stu-id="88e2c-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="88e2c-210">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="88e2c-210">String</span></span>|<span data-ttu-id="88e2c-211">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="88e2c-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="88e2c-212">String</span><span class="sxs-lookup"><span data-stu-id="88e2c-212">String</span></span>|<span data-ttu-id="88e2c-213">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="88e2c-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88e2c-214">Requisitos</span><span class="sxs-lookup"><span data-stu-id="88e2c-214">Requirements</span></span>

|<span data-ttu-id="88e2c-215">Requisito</span><span class="sxs-lookup"><span data-stu-id="88e2c-215">Requirement</span></span>| <span data-ttu-id="88e2c-216">Valor</span><span class="sxs-lookup"><span data-stu-id="88e2c-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="88e2c-217">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="88e2c-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88e2c-218">1.0</span><span class="sxs-lookup"><span data-stu-id="88e2c-218">1.0</span></span>|
|[<span data-ttu-id="88e2c-219">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="88e2c-219">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88e2c-220">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="88e2c-220">Compose or Read</span></span>|
