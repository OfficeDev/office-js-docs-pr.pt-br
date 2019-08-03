---
title: Namespace do Office – conjunto de requisitos 1,6
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: e211a3a2983567b79b73a791914f8d4ed1501ab1
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064660"
---
# <a name="office"></a><span data-ttu-id="d0e3f-102">Office</span><span class="sxs-lookup"><span data-stu-id="d0e3f-102">Office</span></span>

<span data-ttu-id="d0e3f-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d0e3f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d0e3f-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d0e3f-105">Requirements</span></span>

|<span data-ttu-id="d0e3f-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="d0e3f-106">Requirement</span></span>| <span data-ttu-id="d0e3f-107">Valor</span><span class="sxs-lookup"><span data-stu-id="d0e3f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0e3f-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d0e3f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0e3f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d0e3f-109">1.0</span></span>|
|[<span data-ttu-id="d0e3f-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d0e3f-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d0e3f-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d0e3f-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d0e3f-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="d0e3f-112">Members and methods</span></span>

| <span data-ttu-id="d0e3f-113">Membro</span><span class="sxs-lookup"><span data-stu-id="d0e3f-113">Member</span></span> | <span data-ttu-id="d0e3f-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="d0e3f-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d0e3f-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d0e3f-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d0e3f-116">Membro</span><span class="sxs-lookup"><span data-stu-id="d0e3f-116">Member</span></span> |
| [<span data-ttu-id="d0e3f-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d0e3f-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d0e3f-118">Membro</span><span class="sxs-lookup"><span data-stu-id="d0e3f-118">Member</span></span> |
| [<span data-ttu-id="d0e3f-119">EventType</span><span class="sxs-lookup"><span data-stu-id="d0e3f-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="d0e3f-120">Membro</span><span class="sxs-lookup"><span data-stu-id="d0e3f-120">Member</span></span> |
| [<span data-ttu-id="d0e3f-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d0e3f-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d0e3f-122">Membro</span><span class="sxs-lookup"><span data-stu-id="d0e3f-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d0e3f-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="d0e3f-123">Namespaces</span></span>

<span data-ttu-id="d0e3f-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d0e3f-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="d0e3f-126">Membros</span><span class="sxs-lookup"><span data-stu-id="d0e3f-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="d0e3f-127">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d0e3f-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="d0e3f-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d0e3f-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="d0e3f-129">Type</span></span>

*   <span data-ttu-id="d0e3f-130">String</span><span class="sxs-lookup"><span data-stu-id="d0e3f-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d0e3f-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="d0e3f-131">Properties:</span></span>

|<span data-ttu-id="d0e3f-132">Nome</span><span class="sxs-lookup"><span data-stu-id="d0e3f-132">Name</span></span>| <span data-ttu-id="d0e3f-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="d0e3f-133">Type</span></span>| <span data-ttu-id="d0e3f-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="d0e3f-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d0e3f-135">String</span><span class="sxs-lookup"><span data-stu-id="d0e3f-135">String</span></span>|<span data-ttu-id="d0e3f-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d0e3f-137">String</span><span class="sxs-lookup"><span data-stu-id="d0e3f-137">String</span></span>|<span data-ttu-id="d0e3f-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0e3f-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d0e3f-139">Requirements</span></span>

|<span data-ttu-id="d0e3f-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="d0e3f-140">Requirement</span></span>| <span data-ttu-id="d0e3f-141">Valor</span><span class="sxs-lookup"><span data-stu-id="d0e3f-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0e3f-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d0e3f-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0e3f-143">1.0</span><span class="sxs-lookup"><span data-stu-id="d0e3f-143">1.0</span></span>|
|[<span data-ttu-id="d0e3f-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d0e3f-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d0e3f-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d0e3f-145">Compose or Read</span></span>|

---

#### <a name="coerciontype-string"></a><span data-ttu-id="d0e3f-146">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d0e3f-146">CoercionType: String</span></span>

<span data-ttu-id="d0e3f-147">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d0e3f-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="d0e3f-148">Type</span></span>

*   <span data-ttu-id="d0e3f-149">String</span><span class="sxs-lookup"><span data-stu-id="d0e3f-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d0e3f-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="d0e3f-150">Properties:</span></span>

|<span data-ttu-id="d0e3f-151">Nome</span><span class="sxs-lookup"><span data-stu-id="d0e3f-151">Name</span></span>| <span data-ttu-id="d0e3f-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="d0e3f-152">Type</span></span>| <span data-ttu-id="d0e3f-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="d0e3f-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d0e3f-154">String</span><span class="sxs-lookup"><span data-stu-id="d0e3f-154">String</span></span>|<span data-ttu-id="d0e3f-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d0e3f-156">String</span><span class="sxs-lookup"><span data-stu-id="d0e3f-156">String</span></span>|<span data-ttu-id="d0e3f-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0e3f-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d0e3f-158">Requirements</span></span>

|<span data-ttu-id="d0e3f-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="d0e3f-159">Requirement</span></span>| <span data-ttu-id="d0e3f-160">Valor</span><span class="sxs-lookup"><span data-stu-id="d0e3f-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0e3f-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d0e3f-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0e3f-162">1.0</span><span class="sxs-lookup"><span data-stu-id="d0e3f-162">1.0</span></span>|
|[<span data-ttu-id="d0e3f-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d0e3f-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d0e3f-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d0e3f-164">Compose or Read</span></span>|

---

#### <a name="eventtype-string"></a><span data-ttu-id="d0e3f-165">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d0e3f-165">EventType: String</span></span>

<span data-ttu-id="d0e3f-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="d0e3f-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="d0e3f-167">Type</span></span>

*   <span data-ttu-id="d0e3f-168">String</span><span class="sxs-lookup"><span data-stu-id="d0e3f-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d0e3f-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="d0e3f-169">Properties:</span></span>

| <span data-ttu-id="d0e3f-170">Nome</span><span class="sxs-lookup"><span data-stu-id="d0e3f-170">Name</span></span> | <span data-ttu-id="d0e3f-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="d0e3f-171">Type</span></span> | <span data-ttu-id="d0e3f-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="d0e3f-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="d0e3f-173">String</span><span class="sxs-lookup"><span data-stu-id="d0e3f-173">String</span></span> | <span data-ttu-id="d0e3f-174">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d0e3f-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d0e3f-175">Requirements</span></span>

|<span data-ttu-id="d0e3f-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="d0e3f-176">Requirement</span></span>| <span data-ttu-id="d0e3f-177">Valor</span><span class="sxs-lookup"><span data-stu-id="d0e3f-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0e3f-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d0e3f-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0e3f-179">1,5</span><span class="sxs-lookup"><span data-stu-id="d0e3f-179">1.5</span></span> |
|[<span data-ttu-id="d0e3f-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d0e3f-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d0e3f-181">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d0e3f-181">Compose or Read</span></span> |

---

#### <a name="sourceproperty-string"></a><span data-ttu-id="d0e3f-182">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d0e3f-182">SourceProperty: String</span></span>

<span data-ttu-id="d0e3f-183">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d0e3f-184">Tipo</span><span class="sxs-lookup"><span data-stu-id="d0e3f-184">Type</span></span>

*   <span data-ttu-id="d0e3f-185">String</span><span class="sxs-lookup"><span data-stu-id="d0e3f-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d0e3f-186">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="d0e3f-186">Properties:</span></span>

|<span data-ttu-id="d0e3f-187">Nome</span><span class="sxs-lookup"><span data-stu-id="d0e3f-187">Name</span></span>| <span data-ttu-id="d0e3f-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="d0e3f-188">Type</span></span>| <span data-ttu-id="d0e3f-189">Descrição</span><span class="sxs-lookup"><span data-stu-id="d0e3f-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d0e3f-190">String</span><span class="sxs-lookup"><span data-stu-id="d0e3f-190">String</span></span>|<span data-ttu-id="d0e3f-191">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d0e3f-192">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="d0e3f-192">String</span></span>|<span data-ttu-id="d0e3f-193">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="d0e3f-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d0e3f-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="d0e3f-194">Requirements</span></span>

|<span data-ttu-id="d0e3f-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="d0e3f-195">Requirement</span></span>| <span data-ttu-id="d0e3f-196">Valor</span><span class="sxs-lookup"><span data-stu-id="d0e3f-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0e3f-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="d0e3f-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0e3f-198">1.0</span><span class="sxs-lookup"><span data-stu-id="d0e3f-198">1.0</span></span>|
|[<span data-ttu-id="d0e3f-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="d0e3f-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d0e3f-200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="d0e3f-200">Compose or Read</span></span>|
