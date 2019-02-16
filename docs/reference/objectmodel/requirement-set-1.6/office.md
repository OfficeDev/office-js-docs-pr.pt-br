---
title: 'Namespace do Office: conjunto de requisitos da versão 1.6'
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 29b3d58a4cd9dad631c2b23cabc84ade45260451
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068151"
---
# <a name="office"></a><span data-ttu-id="269e6-102">Office</span><span class="sxs-lookup"><span data-stu-id="269e6-102">Office</span></span>

<span data-ttu-id="269e6-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="269e6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="269e6-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="269e6-105">Requirements</span></span>

|<span data-ttu-id="269e6-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="269e6-106">Requirement</span></span>| <span data-ttu-id="269e6-107">Valor</span><span class="sxs-lookup"><span data-stu-id="269e6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="269e6-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="269e6-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="269e6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="269e6-109">1.0</span></span>|
|[<span data-ttu-id="269e6-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="269e6-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="269e6-111">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="269e6-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="269e6-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="269e6-112">Members and methods</span></span>

| <span data-ttu-id="269e6-113">Membro</span><span class="sxs-lookup"><span data-stu-id="269e6-113">Member</span></span> | <span data-ttu-id="269e6-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="269e6-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="269e6-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="269e6-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="269e6-116">Membro</span><span class="sxs-lookup"><span data-stu-id="269e6-116">Member</span></span> |
| [<span data-ttu-id="269e6-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="269e6-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="269e6-118">Membro</span><span class="sxs-lookup"><span data-stu-id="269e6-118">Member</span></span> |
| [<span data-ttu-id="269e6-119">EventType</span><span class="sxs-lookup"><span data-stu-id="269e6-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="269e6-120">Membro</span><span class="sxs-lookup"><span data-stu-id="269e6-120">Member</span></span> |
| [<span data-ttu-id="269e6-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="269e6-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="269e6-122">Membro</span><span class="sxs-lookup"><span data-stu-id="269e6-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="269e6-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="269e6-123">Namespaces</span></span>

<span data-ttu-id="269e6-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="269e6-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="269e6-125">[MailboxEnums](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="269e6-125">[MailboxEnums](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="269e6-126">Membros</span><span class="sxs-lookup"><span data-stu-id="269e6-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="269e6-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="269e6-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="269e6-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="269e6-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="269e6-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="269e6-129">Type</span></span>

*   <span data-ttu-id="269e6-130">String</span><span class="sxs-lookup"><span data-stu-id="269e6-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="269e6-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="269e6-131">Properties:</span></span>

|<span data-ttu-id="269e6-132">Nome</span><span class="sxs-lookup"><span data-stu-id="269e6-132">Name</span></span>| <span data-ttu-id="269e6-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="269e6-133">Type</span></span>| <span data-ttu-id="269e6-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="269e6-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="269e6-135">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="269e6-135">String</span></span>|<span data-ttu-id="269e6-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="269e6-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="269e6-137">String</span><span class="sxs-lookup"><span data-stu-id="269e6-137">String</span></span>|<span data-ttu-id="269e6-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="269e6-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="269e6-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="269e6-139">Requirements</span></span>

|<span data-ttu-id="269e6-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="269e6-140">Requirement</span></span>| <span data-ttu-id="269e6-141">Valor</span><span class="sxs-lookup"><span data-stu-id="269e6-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="269e6-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="269e6-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="269e6-143">1.0</span><span class="sxs-lookup"><span data-stu-id="269e6-143">1.0</span></span>|
|[<span data-ttu-id="269e6-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="269e6-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="269e6-145">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="269e6-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="269e6-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="269e6-146">CoercionType :String</span></span>

<span data-ttu-id="269e6-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="269e6-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="269e6-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="269e6-148">Type</span></span>

*   <span data-ttu-id="269e6-149">String</span><span class="sxs-lookup"><span data-stu-id="269e6-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="269e6-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="269e6-150">Properties:</span></span>

|<span data-ttu-id="269e6-151">Nome</span><span class="sxs-lookup"><span data-stu-id="269e6-151">Name</span></span>| <span data-ttu-id="269e6-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="269e6-152">Type</span></span>| <span data-ttu-id="269e6-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="269e6-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="269e6-154">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="269e6-154">String</span></span>|<span data-ttu-id="269e6-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="269e6-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="269e6-156">String</span><span class="sxs-lookup"><span data-stu-id="269e6-156">String</span></span>|<span data-ttu-id="269e6-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="269e6-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="269e6-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="269e6-158">Requirements</span></span>

|<span data-ttu-id="269e6-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="269e6-159">Requirement</span></span>| <span data-ttu-id="269e6-160">Valor</span><span class="sxs-lookup"><span data-stu-id="269e6-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="269e6-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="269e6-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="269e6-162">1.0</span><span class="sxs-lookup"><span data-stu-id="269e6-162">1.0</span></span>|
|[<span data-ttu-id="269e6-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="269e6-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="269e6-164">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="269e6-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="269e6-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="269e6-165">EventType :String</span></span>

<span data-ttu-id="269e6-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="269e6-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="269e6-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="269e6-167">Type</span></span>

*   <span data-ttu-id="269e6-168">String</span><span class="sxs-lookup"><span data-stu-id="269e6-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="269e6-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="269e6-169">Properties:</span></span>

| <span data-ttu-id="269e6-170">Nome</span><span class="sxs-lookup"><span data-stu-id="269e6-170">Name</span></span> | <span data-ttu-id="269e6-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="269e6-171">Type</span></span> | <span data-ttu-id="269e6-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="269e6-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="269e6-173">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="269e6-173">String</span></span> | <span data-ttu-id="269e6-174">Um item diferente do Outlook está marcado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="269e6-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="269e6-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="269e6-175">Requirements</span></span>

|<span data-ttu-id="269e6-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="269e6-176">Requirement</span></span>| <span data-ttu-id="269e6-177">Valor</span><span class="sxs-lookup"><span data-stu-id="269e6-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="269e6-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="269e6-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="269e6-179">1.5</span><span class="sxs-lookup"><span data-stu-id="269e6-179">1.5</span></span> |
|[<span data-ttu-id="269e6-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="269e6-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="269e6-181">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="269e6-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="269e6-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="269e6-182">SourceProperty :String</span></span>

<span data-ttu-id="269e6-183">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="269e6-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="269e6-184">Tipo</span><span class="sxs-lookup"><span data-stu-id="269e6-184">Type</span></span>

*   <span data-ttu-id="269e6-185">String</span><span class="sxs-lookup"><span data-stu-id="269e6-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="269e6-186">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="269e6-186">Properties:</span></span>

|<span data-ttu-id="269e6-187">Nome</span><span class="sxs-lookup"><span data-stu-id="269e6-187">Name</span></span>| <span data-ttu-id="269e6-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="269e6-188">Type</span></span>| <span data-ttu-id="269e6-189">Descrição</span><span class="sxs-lookup"><span data-stu-id="269e6-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="269e6-190">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="269e6-190">String</span></span>|<span data-ttu-id="269e6-191">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="269e6-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="269e6-192">String</span><span class="sxs-lookup"><span data-stu-id="269e6-192">String</span></span>|<span data-ttu-id="269e6-193">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="269e6-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="269e6-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="269e6-194">Requirements</span></span>

|<span data-ttu-id="269e6-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="269e6-195">Requirement</span></span>| <span data-ttu-id="269e6-196">Valor</span><span class="sxs-lookup"><span data-stu-id="269e6-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="269e6-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="269e6-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="269e6-198">1.0</span><span class="sxs-lookup"><span data-stu-id="269e6-198">1.0</span></span>|
|[<span data-ttu-id="269e6-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="269e6-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="269e6-200">Redação ou leitura</span><span class="sxs-lookup"><span data-stu-id="269e6-200">Compose or Read</span></span>|
