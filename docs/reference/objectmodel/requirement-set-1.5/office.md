---
title: Namespace do Office – conjunto de requisitos 1,5
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 402737f0f6648e42f569906df59be0fa26991146
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395684"
---
# <a name="office"></a><span data-ttu-id="a8867-102">Office</span><span class="sxs-lookup"><span data-stu-id="a8867-102">Office</span></span>

<span data-ttu-id="a8867-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a8867-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a8867-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a8867-105">Requirements</span></span>

|<span data-ttu-id="a8867-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="a8867-106">Requirement</span></span>| <span data-ttu-id="a8867-107">Valor</span><span class="sxs-lookup"><span data-stu-id="a8867-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8867-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a8867-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8867-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a8867-109">1.0</span></span>|
|[<span data-ttu-id="a8867-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a8867-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a8867-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a8867-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a8867-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="a8867-112">Members and methods</span></span>

| <span data-ttu-id="a8867-113">Membro</span><span class="sxs-lookup"><span data-stu-id="a8867-113">Member</span></span> | <span data-ttu-id="a8867-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="a8867-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a8867-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a8867-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a8867-116">Membro</span><span class="sxs-lookup"><span data-stu-id="a8867-116">Member</span></span> |
| [<span data-ttu-id="a8867-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a8867-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a8867-118">Membro</span><span class="sxs-lookup"><span data-stu-id="a8867-118">Member</span></span> |
| [<span data-ttu-id="a8867-119">EventType</span><span class="sxs-lookup"><span data-stu-id="a8867-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="a8867-120">Membro</span><span class="sxs-lookup"><span data-stu-id="a8867-120">Member</span></span> |
| [<span data-ttu-id="a8867-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a8867-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a8867-122">Membro</span><span class="sxs-lookup"><span data-stu-id="a8867-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a8867-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="a8867-123">Namespaces</span></span>

<span data-ttu-id="a8867-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a8867-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a8867-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5): inclui um número de enumerações, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="a8867-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="a8867-126">Members</span><span class="sxs-lookup"><span data-stu-id="a8867-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="a8867-127">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a8867-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="a8867-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="a8867-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a8867-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="a8867-129">Type</span></span>

*   <span data-ttu-id="a8867-130">String</span><span class="sxs-lookup"><span data-stu-id="a8867-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a8867-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a8867-131">Properties:</span></span>

|<span data-ttu-id="a8867-132">Nome</span><span class="sxs-lookup"><span data-stu-id="a8867-132">Name</span></span>| <span data-ttu-id="a8867-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="a8867-133">Type</span></span>| <span data-ttu-id="a8867-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="a8867-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a8867-135">String</span><span class="sxs-lookup"><span data-stu-id="a8867-135">String</span></span>|<span data-ttu-id="a8867-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="a8867-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a8867-137">String</span><span class="sxs-lookup"><span data-stu-id="a8867-137">String</span></span>|<span data-ttu-id="a8867-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="a8867-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a8867-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a8867-139">Requirements</span></span>

|<span data-ttu-id="a8867-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="a8867-140">Requirement</span></span>| <span data-ttu-id="a8867-141">Valor</span><span class="sxs-lookup"><span data-stu-id="a8867-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8867-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a8867-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8867-143">1.0</span><span class="sxs-lookup"><span data-stu-id="a8867-143">1.0</span></span>|
|[<span data-ttu-id="a8867-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a8867-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a8867-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a8867-145">Compose or Read</span></span>|

---

#### <a name="coerciontype-string"></a><span data-ttu-id="a8867-146">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a8867-146">CoercionType: String</span></span>

<span data-ttu-id="a8867-147">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="a8867-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a8867-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="a8867-148">Type</span></span>

*   <span data-ttu-id="a8867-149">String</span><span class="sxs-lookup"><span data-stu-id="a8867-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a8867-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a8867-150">Properties:</span></span>

|<span data-ttu-id="a8867-151">Nome</span><span class="sxs-lookup"><span data-stu-id="a8867-151">Name</span></span>| <span data-ttu-id="a8867-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="a8867-152">Type</span></span>| <span data-ttu-id="a8867-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="a8867-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a8867-154">String</span><span class="sxs-lookup"><span data-stu-id="a8867-154">String</span></span>|<span data-ttu-id="a8867-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="a8867-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a8867-156">String</span><span class="sxs-lookup"><span data-stu-id="a8867-156">String</span></span>|<span data-ttu-id="a8867-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="a8867-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a8867-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a8867-158">Requirements</span></span>

|<span data-ttu-id="a8867-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="a8867-159">Requirement</span></span>| <span data-ttu-id="a8867-160">Valor</span><span class="sxs-lookup"><span data-stu-id="a8867-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8867-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a8867-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8867-162">1.0</span><span class="sxs-lookup"><span data-stu-id="a8867-162">1.0</span></span>|
|[<span data-ttu-id="a8867-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a8867-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a8867-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a8867-164">Compose or Read</span></span>|

---

#### <a name="eventtype-string"></a><span data-ttu-id="a8867-165">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a8867-165">EventType: String</span></span>

<span data-ttu-id="a8867-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="a8867-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="a8867-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="a8867-167">Type</span></span>

*   <span data-ttu-id="a8867-168">String</span><span class="sxs-lookup"><span data-stu-id="a8867-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a8867-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a8867-169">Properties:</span></span>

| <span data-ttu-id="a8867-170">Nome</span><span class="sxs-lookup"><span data-stu-id="a8867-170">Name</span></span> | <span data-ttu-id="a8867-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="a8867-171">Type</span></span> | <span data-ttu-id="a8867-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="a8867-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="a8867-173">String</span><span class="sxs-lookup"><span data-stu-id="a8867-173">String</span></span> | <span data-ttu-id="a8867-174">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="a8867-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a8867-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a8867-175">Requirements</span></span>

|<span data-ttu-id="a8867-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="a8867-176">Requirement</span></span>| <span data-ttu-id="a8867-177">Valor</span><span class="sxs-lookup"><span data-stu-id="a8867-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8867-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a8867-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8867-179">1,5</span><span class="sxs-lookup"><span data-stu-id="a8867-179">1.5</span></span> |
|[<span data-ttu-id="a8867-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a8867-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a8867-181">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a8867-181">Compose or Read</span></span> |

---

#### <a name="sourceproperty-string"></a><span data-ttu-id="a8867-182">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a8867-182">SourceProperty: String</span></span>

<span data-ttu-id="a8867-183">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="a8867-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a8867-184">Tipo</span><span class="sxs-lookup"><span data-stu-id="a8867-184">Type</span></span>

*   <span data-ttu-id="a8867-185">String</span><span class="sxs-lookup"><span data-stu-id="a8867-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a8867-186">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a8867-186">Properties:</span></span>

|<span data-ttu-id="a8867-187">Nome</span><span class="sxs-lookup"><span data-stu-id="a8867-187">Name</span></span>| <span data-ttu-id="a8867-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="a8867-188">Type</span></span>| <span data-ttu-id="a8867-189">Descrição</span><span class="sxs-lookup"><span data-stu-id="a8867-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a8867-190">String</span><span class="sxs-lookup"><span data-stu-id="a8867-190">String</span></span>|<span data-ttu-id="a8867-191">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a8867-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a8867-192">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a8867-192">String</span></span>|<span data-ttu-id="a8867-193">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a8867-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a8867-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a8867-194">Requirements</span></span>

|<span data-ttu-id="a8867-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="a8867-195">Requirement</span></span>| <span data-ttu-id="a8867-196">Valor</span><span class="sxs-lookup"><span data-stu-id="a8867-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8867-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a8867-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8867-198">1.0</span><span class="sxs-lookup"><span data-stu-id="a8867-198">1.0</span></span>|
|[<span data-ttu-id="a8867-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a8867-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a8867-200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a8867-200">Compose or Read</span></span>|
