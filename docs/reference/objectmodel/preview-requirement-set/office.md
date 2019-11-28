---
title: Namespace do Office – conjunto de requisitos de visualização
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: bd37b1be4d77d73cb56b0b2593ccc57dea6cab27
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629227"
---
# <a name="office"></a><span data-ttu-id="98f35-102">Office</span><span class="sxs-lookup"><span data-stu-id="98f35-102">Office</span></span>

<span data-ttu-id="98f35-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="98f35-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="98f35-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98f35-105">Requirements</span></span>

|<span data-ttu-id="98f35-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="98f35-106">Requirement</span></span>| <span data-ttu-id="98f35-107">Valor</span><span class="sxs-lookup"><span data-stu-id="98f35-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="98f35-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98f35-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98f35-109">1.0</span><span class="sxs-lookup"><span data-stu-id="98f35-109">1.0</span></span>|
|[<span data-ttu-id="98f35-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98f35-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="98f35-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="98f35-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="98f35-112">Propriedades</span><span class="sxs-lookup"><span data-stu-id="98f35-112">Properties</span></span>

| <span data-ttu-id="98f35-113">Propriedade</span><span class="sxs-lookup"><span data-stu-id="98f35-113">Property</span></span> | <span data-ttu-id="98f35-114">Modelos</span><span class="sxs-lookup"><span data-stu-id="98f35-114">Modes</span></span> | <span data-ttu-id="98f35-115">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="98f35-115">Return type</span></span> | <span data-ttu-id="98f35-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="98f35-116">Minimum</span></span><br><span data-ttu-id="98f35-117">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="98f35-117">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="98f35-118">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="98f35-118">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="98f35-119">Escrever</span><span class="sxs-lookup"><span data-stu-id="98f35-119">Compose</span></span><br><span data-ttu-id="98f35-120">Ler</span><span class="sxs-lookup"><span data-stu-id="98f35-120">Read</span></span> | <span data-ttu-id="98f35-121">String</span><span class="sxs-lookup"><span data-stu-id="98f35-121">String</span></span> | <span data-ttu-id="98f35-122">1.0</span><span class="sxs-lookup"><span data-stu-id="98f35-122">1.0</span></span> |
| [<span data-ttu-id="98f35-123">CoercionType</span><span class="sxs-lookup"><span data-stu-id="98f35-123">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="98f35-124">Escrever</span><span class="sxs-lookup"><span data-stu-id="98f35-124">Compose</span></span><br><span data-ttu-id="98f35-125">Ler</span><span class="sxs-lookup"><span data-stu-id="98f35-125">Read</span></span> | <span data-ttu-id="98f35-126">String</span><span class="sxs-lookup"><span data-stu-id="98f35-126">String</span></span> | <span data-ttu-id="98f35-127">1.0</span><span class="sxs-lookup"><span data-stu-id="98f35-127">1.0</span></span> |
| [<span data-ttu-id="98f35-128">EventType</span><span class="sxs-lookup"><span data-stu-id="98f35-128">EventType</span></span>](#eventtype-string) | <span data-ttu-id="98f35-129">Escrever</span><span class="sxs-lookup"><span data-stu-id="98f35-129">Compose</span></span><br><span data-ttu-id="98f35-130">Ler</span><span class="sxs-lookup"><span data-stu-id="98f35-130">Read</span></span> | <span data-ttu-id="98f35-131">String</span><span class="sxs-lookup"><span data-stu-id="98f35-131">String</span></span> | <span data-ttu-id="98f35-132">1,5</span><span class="sxs-lookup"><span data-stu-id="98f35-132">1.5</span></span> |
| [<span data-ttu-id="98f35-133">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="98f35-133">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="98f35-134">Escrever</span><span class="sxs-lookup"><span data-stu-id="98f35-134">Compose</span></span><br><span data-ttu-id="98f35-135">Ler</span><span class="sxs-lookup"><span data-stu-id="98f35-135">Read</span></span> | <span data-ttu-id="98f35-136">String</span><span class="sxs-lookup"><span data-stu-id="98f35-136">String</span></span> | <span data-ttu-id="98f35-137">1.0</span><span class="sxs-lookup"><span data-stu-id="98f35-137">1.0</span></span> |

### <a name="namespaces"></a><span data-ttu-id="98f35-138">Namespaces</span><span class="sxs-lookup"><span data-stu-id="98f35-138">Namespaces</span></span>

<span data-ttu-id="98f35-139">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="98f35-139">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="98f35-140">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): inclui um número de enumerações, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="98f35-140">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="property-details"></a><span data-ttu-id="98f35-141">Detalhes da propriedade</span><span class="sxs-lookup"><span data-stu-id="98f35-141">Property details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="98f35-142">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="98f35-142">AsyncResultStatus: String</span></span>

<span data-ttu-id="98f35-143">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="98f35-143">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="98f35-144">Tipo</span><span class="sxs-lookup"><span data-stu-id="98f35-144">Type</span></span>

*   <span data-ttu-id="98f35-145">String</span><span class="sxs-lookup"><span data-stu-id="98f35-145">String</span></span>

##### <a name="properties"></a><span data-ttu-id="98f35-146">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="98f35-146">Properties:</span></span>

|<span data-ttu-id="98f35-147">Nome</span><span class="sxs-lookup"><span data-stu-id="98f35-147">Name</span></span>| <span data-ttu-id="98f35-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="98f35-148">Type</span></span>| <span data-ttu-id="98f35-149">Descrição</span><span class="sxs-lookup"><span data-stu-id="98f35-149">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="98f35-150">String</span><span class="sxs-lookup"><span data-stu-id="98f35-150">String</span></span>|<span data-ttu-id="98f35-151">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="98f35-151">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="98f35-152">String</span><span class="sxs-lookup"><span data-stu-id="98f35-152">String</span></span>|<span data-ttu-id="98f35-153">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="98f35-153">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="98f35-154">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98f35-154">Requirements</span></span>

|<span data-ttu-id="98f35-155">Requisito</span><span class="sxs-lookup"><span data-stu-id="98f35-155">Requirement</span></span>| <span data-ttu-id="98f35-156">Valor</span><span class="sxs-lookup"><span data-stu-id="98f35-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="98f35-157">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98f35-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98f35-158">1.0</span><span class="sxs-lookup"><span data-stu-id="98f35-158">1.0</span></span>|
|[<span data-ttu-id="98f35-159">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98f35-159">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="98f35-160">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="98f35-160">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="98f35-161">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="98f35-161">CoercionType: String</span></span>

<span data-ttu-id="98f35-162">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="98f35-162">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="98f35-163">Tipo</span><span class="sxs-lookup"><span data-stu-id="98f35-163">Type</span></span>

*   <span data-ttu-id="98f35-164">String</span><span class="sxs-lookup"><span data-stu-id="98f35-164">String</span></span>

##### <a name="properties"></a><span data-ttu-id="98f35-165">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="98f35-165">Properties:</span></span>

|<span data-ttu-id="98f35-166">Nome</span><span class="sxs-lookup"><span data-stu-id="98f35-166">Name</span></span>| <span data-ttu-id="98f35-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="98f35-167">Type</span></span>| <span data-ttu-id="98f35-168">Descrição</span><span class="sxs-lookup"><span data-stu-id="98f35-168">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="98f35-169">String</span><span class="sxs-lookup"><span data-stu-id="98f35-169">String</span></span>|<span data-ttu-id="98f35-170">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="98f35-170">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="98f35-171">String</span><span class="sxs-lookup"><span data-stu-id="98f35-171">String</span></span>|<span data-ttu-id="98f35-172">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="98f35-172">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="98f35-173">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98f35-173">Requirements</span></span>

|<span data-ttu-id="98f35-174">Requisito</span><span class="sxs-lookup"><span data-stu-id="98f35-174">Requirement</span></span>| <span data-ttu-id="98f35-175">Valor</span><span class="sxs-lookup"><span data-stu-id="98f35-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="98f35-176">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98f35-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98f35-177">1.0</span><span class="sxs-lookup"><span data-stu-id="98f35-177">1.0</span></span>|
|[<span data-ttu-id="98f35-178">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98f35-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="98f35-179">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="98f35-179">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="98f35-180">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="98f35-180">EventType: String</span></span>

<span data-ttu-id="98f35-181">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="98f35-181">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="98f35-182">Tipo</span><span class="sxs-lookup"><span data-stu-id="98f35-182">Type</span></span>

*   <span data-ttu-id="98f35-183">String</span><span class="sxs-lookup"><span data-stu-id="98f35-183">String</span></span>

##### <a name="properties"></a><span data-ttu-id="98f35-184">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="98f35-184">Properties:</span></span>

| <span data-ttu-id="98f35-185">Nome</span><span class="sxs-lookup"><span data-stu-id="98f35-185">Name</span></span> | <span data-ttu-id="98f35-186">Tipo</span><span class="sxs-lookup"><span data-stu-id="98f35-186">Type</span></span> | <span data-ttu-id="98f35-187">Descrição</span><span class="sxs-lookup"><span data-stu-id="98f35-187">Description</span></span> | <span data-ttu-id="98f35-188">Conjunto de requisitos mínimo</span><span class="sxs-lookup"><span data-stu-id="98f35-188">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="98f35-189">String</span><span class="sxs-lookup"><span data-stu-id="98f35-189">String</span></span> | <span data-ttu-id="98f35-190">A data ou hora do compromisso ou série selecionado foi alterada.</span><span class="sxs-lookup"><span data-stu-id="98f35-190">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="98f35-191">1.7</span><span class="sxs-lookup"><span data-stu-id="98f35-191">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="98f35-192">String</span><span class="sxs-lookup"><span data-stu-id="98f35-192">String</span></span> | <span data-ttu-id="98f35-193">Um anexo foi adicionado ou removido do item.</span><span class="sxs-lookup"><span data-stu-id="98f35-193">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="98f35-194">1,8</span><span class="sxs-lookup"><span data-stu-id="98f35-194">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="98f35-195">String</span><span class="sxs-lookup"><span data-stu-id="98f35-195">String</span></span> | <span data-ttu-id="98f35-196">O local do compromisso selecionado foi alterado.</span><span class="sxs-lookup"><span data-stu-id="98f35-196">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="98f35-197">1,8</span><span class="sxs-lookup"><span data-stu-id="98f35-197">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="98f35-198">String</span><span class="sxs-lookup"><span data-stu-id="98f35-198">String</span></span> | <span data-ttu-id="98f35-199">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="98f35-199">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="98f35-200">1,5</span><span class="sxs-lookup"><span data-stu-id="98f35-200">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="98f35-201">String</span><span class="sxs-lookup"><span data-stu-id="98f35-201">String</span></span> | <span data-ttu-id="98f35-202">O tema do Office na caixa de correio foi alterado.</span><span class="sxs-lookup"><span data-stu-id="98f35-202">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="98f35-203">Visualização</span><span class="sxs-lookup"><span data-stu-id="98f35-203">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="98f35-204">String</span><span class="sxs-lookup"><span data-stu-id="98f35-204">String</span></span> | <span data-ttu-id="98f35-205">A lista de destinatários do item selecionado ou local do compromisso foi alterada.</span><span class="sxs-lookup"><span data-stu-id="98f35-205">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="98f35-206">1.7</span><span class="sxs-lookup"><span data-stu-id="98f35-206">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="98f35-207">String</span><span class="sxs-lookup"><span data-stu-id="98f35-207">String</span></span> | <span data-ttu-id="98f35-208">O padrão de recorrência da série selecionada foi alterado.</span><span class="sxs-lookup"><span data-stu-id="98f35-208">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="98f35-209">1.7</span><span class="sxs-lookup"><span data-stu-id="98f35-209">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="98f35-210">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98f35-210">Requirements</span></span>

|<span data-ttu-id="98f35-211">Requisito</span><span class="sxs-lookup"><span data-stu-id="98f35-211">Requirement</span></span>| <span data-ttu-id="98f35-212">Valor</span><span class="sxs-lookup"><span data-stu-id="98f35-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="98f35-213">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98f35-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98f35-214">1,5</span><span class="sxs-lookup"><span data-stu-id="98f35-214">1.5</span></span> |
|[<span data-ttu-id="98f35-215">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98f35-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="98f35-216">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="98f35-216">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="98f35-217">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="98f35-217">SourceProperty: String</span></span>

<span data-ttu-id="98f35-218">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="98f35-218">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="98f35-219">Tipo</span><span class="sxs-lookup"><span data-stu-id="98f35-219">Type</span></span>

*   <span data-ttu-id="98f35-220">String</span><span class="sxs-lookup"><span data-stu-id="98f35-220">String</span></span>

##### <a name="properties"></a><span data-ttu-id="98f35-221">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="98f35-221">Properties:</span></span>

|<span data-ttu-id="98f35-222">Nome</span><span class="sxs-lookup"><span data-stu-id="98f35-222">Name</span></span>| <span data-ttu-id="98f35-223">Tipo</span><span class="sxs-lookup"><span data-stu-id="98f35-223">Type</span></span>| <span data-ttu-id="98f35-224">Descrição</span><span class="sxs-lookup"><span data-stu-id="98f35-224">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="98f35-225">String</span><span class="sxs-lookup"><span data-stu-id="98f35-225">String</span></span>|<span data-ttu-id="98f35-226">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="98f35-226">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="98f35-227">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="98f35-227">String</span></span>|<span data-ttu-id="98f35-228">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="98f35-228">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="98f35-229">Requisitos</span><span class="sxs-lookup"><span data-stu-id="98f35-229">Requirements</span></span>

|<span data-ttu-id="98f35-230">Requisito</span><span class="sxs-lookup"><span data-stu-id="98f35-230">Requirement</span></span>| <span data-ttu-id="98f35-231">Valor</span><span class="sxs-lookup"><span data-stu-id="98f35-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="98f35-232">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="98f35-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="98f35-233">1.0</span><span class="sxs-lookup"><span data-stu-id="98f35-233">1.0</span></span>|
|[<span data-ttu-id="98f35-234">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="98f35-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="98f35-235">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="98f35-235">Compose or Read</span></span>|
