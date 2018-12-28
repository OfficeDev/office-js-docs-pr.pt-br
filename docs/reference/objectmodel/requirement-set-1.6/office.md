---
title: 'Namespace do Office: conjunto de requisitos da versão 1.6'
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 73411efee9dcfffa5f9f0fa9de85dafc31a4173a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432281"
---
# <a name="office"></a><span data-ttu-id="fba01-102">Office</span><span class="sxs-lookup"><span data-stu-id="fba01-102">Office</span></span>

<span data-ttu-id="fba01-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="fba01-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="fba01-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fba01-105">Requirements</span></span>

|<span data-ttu-id="fba01-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="fba01-106">Requirement</span></span>| <span data-ttu-id="fba01-107">Valor</span><span class="sxs-lookup"><span data-stu-id="fba01-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="fba01-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fba01-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fba01-109">1.0</span><span class="sxs-lookup"><span data-stu-id="fba01-109">1.0</span></span>|
|[<span data-ttu-id="fba01-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fba01-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fba01-111">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="fba01-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fba01-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="fba01-112">Members and methods</span></span>

| <span data-ttu-id="fba01-113">Membro</span><span class="sxs-lookup"><span data-stu-id="fba01-113">Member</span></span> | <span data-ttu-id="fba01-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="fba01-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fba01-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="fba01-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="fba01-116">Membro</span><span class="sxs-lookup"><span data-stu-id="fba01-116">Member</span></span> |
| [<span data-ttu-id="fba01-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="fba01-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="fba01-118">Membro</span><span class="sxs-lookup"><span data-stu-id="fba01-118">Member</span></span> |
| [<span data-ttu-id="fba01-119">EventType</span><span class="sxs-lookup"><span data-stu-id="fba01-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="fba01-120">Membro</span><span class="sxs-lookup"><span data-stu-id="fba01-120">Member</span></span> |
| [<span data-ttu-id="fba01-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="fba01-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="fba01-122">Membro</span><span class="sxs-lookup"><span data-stu-id="fba01-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="fba01-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="fba01-123">Namespaces</span></span>

<span data-ttu-id="fba01-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="fba01-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="fba01-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="fba01-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="fba01-126">Membros</span><span class="sxs-lookup"><span data-stu-id="fba01-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="fba01-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="fba01-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="fba01-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="fba01-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="fba01-129">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fba01-129">Type:</span></span>

*   <span data-ttu-id="fba01-130">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fba01-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fba01-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="fba01-131">Properties:</span></span>

|<span data-ttu-id="fba01-132">Nome</span><span class="sxs-lookup"><span data-stu-id="fba01-132">Name</span></span>| <span data-ttu-id="fba01-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="fba01-133">Type</span></span>| <span data-ttu-id="fba01-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="fba01-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="fba01-135">String</span><span class="sxs-lookup"><span data-stu-id="fba01-135">String</span></span>|<span data-ttu-id="fba01-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="fba01-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="fba01-137">String</span><span class="sxs-lookup"><span data-stu-id="fba01-137">String</span></span>|<span data-ttu-id="fba01-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="fba01-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fba01-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fba01-139">Requirements</span></span>

|<span data-ttu-id="fba01-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="fba01-140">Requirement</span></span>| <span data-ttu-id="fba01-141">Valor</span><span class="sxs-lookup"><span data-stu-id="fba01-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="fba01-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fba01-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fba01-143">1.0</span><span class="sxs-lookup"><span data-stu-id="fba01-143">1.0</span></span>|
|[<span data-ttu-id="fba01-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fba01-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fba01-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fba01-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="fba01-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="fba01-146">CoercionType :String</span></span>

<span data-ttu-id="fba01-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="fba01-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="fba01-148">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fba01-148">Type:</span></span>

*   <span data-ttu-id="fba01-149">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fba01-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fba01-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="fba01-150">Properties:</span></span>

|<span data-ttu-id="fba01-151">Nome</span><span class="sxs-lookup"><span data-stu-id="fba01-151">Name</span></span>| <span data-ttu-id="fba01-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="fba01-152">Type</span></span>| <span data-ttu-id="fba01-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="fba01-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="fba01-154">String</span><span class="sxs-lookup"><span data-stu-id="fba01-154">String</span></span>|<span data-ttu-id="fba01-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="fba01-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="fba01-156">String</span><span class="sxs-lookup"><span data-stu-id="fba01-156">String</span></span>|<span data-ttu-id="fba01-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="fba01-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fba01-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fba01-158">Requirements</span></span>

|<span data-ttu-id="fba01-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="fba01-159">Requirement</span></span>| <span data-ttu-id="fba01-160">Valor</span><span class="sxs-lookup"><span data-stu-id="fba01-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="fba01-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fba01-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fba01-162">1.0</span><span class="sxs-lookup"><span data-stu-id="fba01-162">1.0</span></span>|
|[<span data-ttu-id="fba01-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fba01-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fba01-164">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fba01-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="fba01-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="fba01-165">EventType :String</span></span>

<span data-ttu-id="fba01-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="fba01-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="fba01-167">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fba01-167">Type:</span></span>

*   <span data-ttu-id="fba01-168">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fba01-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fba01-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="fba01-169">Properties:</span></span>

| <span data-ttu-id="fba01-170">Nome</span><span class="sxs-lookup"><span data-stu-id="fba01-170">Name</span></span> | <span data-ttu-id="fba01-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="fba01-171">Type</span></span> | <span data-ttu-id="fba01-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="fba01-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="fba01-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fba01-173">String</span></span> | <span data-ttu-id="fba01-174">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="fba01-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fba01-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fba01-175">Requirements</span></span>

|<span data-ttu-id="fba01-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="fba01-176">Requirement</span></span>| <span data-ttu-id="fba01-177">Valor</span><span class="sxs-lookup"><span data-stu-id="fba01-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="fba01-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fba01-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fba01-179">1.5</span><span class="sxs-lookup"><span data-stu-id="fba01-179">1.5</span></span> |
|[<span data-ttu-id="fba01-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fba01-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fba01-181">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fba01-181">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="fba01-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="fba01-182">SourceProperty :String</span></span>

<span data-ttu-id="fba01-183">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="fba01-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="fba01-184">Tipo:</span><span class="sxs-lookup"><span data-stu-id="fba01-184">Type:</span></span>

*   <span data-ttu-id="fba01-185">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="fba01-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fba01-186">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="fba01-186">Properties:</span></span>

|<span data-ttu-id="fba01-187">Nome</span><span class="sxs-lookup"><span data-stu-id="fba01-187">Name</span></span>| <span data-ttu-id="fba01-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="fba01-188">Type</span></span>| <span data-ttu-id="fba01-189">Descrição</span><span class="sxs-lookup"><span data-stu-id="fba01-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="fba01-190">String</span><span class="sxs-lookup"><span data-stu-id="fba01-190">String</span></span>|<span data-ttu-id="fba01-191">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fba01-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="fba01-192">String</span><span class="sxs-lookup"><span data-stu-id="fba01-192">String</span></span>|<span data-ttu-id="fba01-193">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="fba01-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fba01-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="fba01-194">Requirements</span></span>

|<span data-ttu-id="fba01-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="fba01-195">Requirement</span></span>| <span data-ttu-id="fba01-196">Valor</span><span class="sxs-lookup"><span data-stu-id="fba01-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="fba01-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="fba01-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fba01-198">1.0</span><span class="sxs-lookup"><span data-stu-id="fba01-198">1.0</span></span>|
|[<span data-ttu-id="fba01-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="fba01-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fba01-200">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="fba01-200">Compose or read</span></span>|