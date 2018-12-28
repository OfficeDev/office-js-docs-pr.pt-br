---
title: Namespace do Office – conjunto de requisitos versão 1.5
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 46b21df77456d2392fbc543e45513246a4ad9a10
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433646"
---
# <a name="office"></a><span data-ttu-id="2b628-102">Office</span><span class="sxs-lookup"><span data-stu-id="2b628-102">Office</span></span>

<span data-ttu-id="2b628-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="2b628-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b628-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2b628-105">Requirements</span></span>

|<span data-ttu-id="2b628-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="2b628-106">Requirement</span></span>| <span data-ttu-id="2b628-107">Valor</span><span class="sxs-lookup"><span data-stu-id="2b628-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b628-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2b628-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b628-109">1.0</span><span class="sxs-lookup"><span data-stu-id="2b628-109">1.0</span></span>|
|[<span data-ttu-id="2b628-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2b628-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2b628-111">Composição ou leitura</span><span class="sxs-lookup"><span data-stu-id="2b628-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2b628-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="2b628-112">Members and methods</span></span>

| <span data-ttu-id="2b628-113">Membro</span><span class="sxs-lookup"><span data-stu-id="2b628-113">Member</span></span> | <span data-ttu-id="2b628-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="2b628-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2b628-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2b628-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2b628-116">Membro</span><span class="sxs-lookup"><span data-stu-id="2b628-116">Member</span></span> |
| [<span data-ttu-id="2b628-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2b628-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2b628-118">Membro</span><span class="sxs-lookup"><span data-stu-id="2b628-118">Member</span></span> |
| [<span data-ttu-id="2b628-119">EventType</span><span class="sxs-lookup"><span data-stu-id="2b628-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2b628-120">Membro</span><span class="sxs-lookup"><span data-stu-id="2b628-120">Member</span></span> |
| [<span data-ttu-id="2b628-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2b628-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2b628-122">Membro</span><span class="sxs-lookup"><span data-stu-id="2b628-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="2b628-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="2b628-123">Namespaces</span></span>

<span data-ttu-id="2b628-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="2b628-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="2b628-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="2b628-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="2b628-126">Membros</span><span class="sxs-lookup"><span data-stu-id="2b628-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="2b628-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="2b628-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="2b628-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="2b628-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2b628-129">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2b628-129">Type:</span></span>

*   <span data-ttu-id="2b628-130">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2b628-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2b628-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="2b628-131">Properties:</span></span>

|<span data-ttu-id="2b628-132">Nome</span><span class="sxs-lookup"><span data-stu-id="2b628-132">Name</span></span>| <span data-ttu-id="2b628-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="2b628-133">Type</span></span>| <span data-ttu-id="2b628-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="2b628-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2b628-135">String</span><span class="sxs-lookup"><span data-stu-id="2b628-135">String</span></span>|<span data-ttu-id="2b628-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="2b628-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2b628-137">String</span><span class="sxs-lookup"><span data-stu-id="2b628-137">String</span></span>|<span data-ttu-id="2b628-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="2b628-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2b628-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2b628-139">Requirements</span></span>

|<span data-ttu-id="2b628-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="2b628-140">Requirement</span></span>| <span data-ttu-id="2b628-141">Valor</span><span class="sxs-lookup"><span data-stu-id="2b628-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b628-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2b628-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b628-143">1.0</span><span class="sxs-lookup"><span data-stu-id="2b628-143">1.0</span></span>|
|[<span data-ttu-id="2b628-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2b628-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2b628-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="2b628-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="2b628-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="2b628-146">CoercionType :String</span></span>

<span data-ttu-id="2b628-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="2b628-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2b628-148">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2b628-148">Type:</span></span>

*   <span data-ttu-id="2b628-149">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2b628-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2b628-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="2b628-150">Properties:</span></span>

|<span data-ttu-id="2b628-151">Nome</span><span class="sxs-lookup"><span data-stu-id="2b628-151">Name</span></span>| <span data-ttu-id="2b628-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="2b628-152">Type</span></span>| <span data-ttu-id="2b628-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="2b628-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2b628-154">String</span><span class="sxs-lookup"><span data-stu-id="2b628-154">String</span></span>|<span data-ttu-id="2b628-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="2b628-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2b628-156">String</span><span class="sxs-lookup"><span data-stu-id="2b628-156">String</span></span>|<span data-ttu-id="2b628-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="2b628-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2b628-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2b628-158">Requirements</span></span>

|<span data-ttu-id="2b628-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="2b628-159">Requirement</span></span>| <span data-ttu-id="2b628-160">Valor</span><span class="sxs-lookup"><span data-stu-id="2b628-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b628-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2b628-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b628-162">1.0</span><span class="sxs-lookup"><span data-stu-id="2b628-162">1.0</span></span>|
|[<span data-ttu-id="2b628-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2b628-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2b628-164">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="2b628-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="2b628-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="2b628-165">EventType :String</span></span>

<span data-ttu-id="2b628-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="2b628-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2b628-167">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2b628-167">Type:</span></span>

*   <span data-ttu-id="2b628-168">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2b628-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2b628-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="2b628-169">Properties:</span></span>

| <span data-ttu-id="2b628-170">Nome</span><span class="sxs-lookup"><span data-stu-id="2b628-170">Name</span></span> | <span data-ttu-id="2b628-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="2b628-171">Type</span></span> | <span data-ttu-id="2b628-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="2b628-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="2b628-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2b628-173">String</span></span> | <span data-ttu-id="2b628-174">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="2b628-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2b628-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2b628-175">Requirements</span></span>

|<span data-ttu-id="2b628-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="2b628-176">Requirement</span></span>| <span data-ttu-id="2b628-177">Valor</span><span class="sxs-lookup"><span data-stu-id="2b628-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b628-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2b628-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b628-179">1.5</span><span class="sxs-lookup"><span data-stu-id="2b628-179">1.5</span></span> |
|[<span data-ttu-id="2b628-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2b628-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2b628-181">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="2b628-181">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="2b628-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="2b628-182">SourceProperty :String</span></span>

<span data-ttu-id="2b628-183">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="2b628-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2b628-184">Tipo:</span><span class="sxs-lookup"><span data-stu-id="2b628-184">Type:</span></span>

*   <span data-ttu-id="2b628-185">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="2b628-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2b628-186">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="2b628-186">Properties:</span></span>

|<span data-ttu-id="2b628-187">Nome</span><span class="sxs-lookup"><span data-stu-id="2b628-187">Name</span></span>| <span data-ttu-id="2b628-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="2b628-188">Type</span></span>| <span data-ttu-id="2b628-189">Descrição</span><span class="sxs-lookup"><span data-stu-id="2b628-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2b628-190">String</span><span class="sxs-lookup"><span data-stu-id="2b628-190">String</span></span>|<span data-ttu-id="2b628-191">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="2b628-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2b628-192">String</span><span class="sxs-lookup"><span data-stu-id="2b628-192">String</span></span>|<span data-ttu-id="2b628-193">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="2b628-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2b628-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="2b628-194">Requirements</span></span>

|<span data-ttu-id="2b628-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="2b628-195">Requirement</span></span>| <span data-ttu-id="2b628-196">Valor</span><span class="sxs-lookup"><span data-stu-id="2b628-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b628-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="2b628-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b628-198">1.0</span><span class="sxs-lookup"><span data-stu-id="2b628-198">1.0</span></span>|
|[<span data-ttu-id="2b628-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="2b628-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2b628-200">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="2b628-200">Compose or read</span></span>|