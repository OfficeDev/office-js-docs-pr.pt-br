---
title: Namespace do Office – conjunto de requisitos 1,6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: dde96f48863459da5072d6b4864169f198264133
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870804"
---
# <a name="office"></a><span data-ttu-id="a6ac8-102">Office</span><span class="sxs-lookup"><span data-stu-id="a6ac8-102">Office</span></span>

<span data-ttu-id="a6ac8-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a6ac8-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6ac8-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6ac8-105">Requirements</span></span>

|<span data-ttu-id="a6ac8-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6ac8-106">Requirement</span></span>| <span data-ttu-id="a6ac8-107">Valor</span><span class="sxs-lookup"><span data-stu-id="a6ac8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6ac8-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6ac8-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6ac8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a6ac8-109">1.0</span></span>|
|[<span data-ttu-id="a6ac8-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6ac8-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6ac8-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a6ac8-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a6ac8-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="a6ac8-112">Members and methods</span></span>

| <span data-ttu-id="a6ac8-113">Membro</span><span class="sxs-lookup"><span data-stu-id="a6ac8-113">Member</span></span> | <span data-ttu-id="a6ac8-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6ac8-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a6ac8-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a6ac8-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a6ac8-116">Member</span><span class="sxs-lookup"><span data-stu-id="a6ac8-116">Member</span></span> |
| [<span data-ttu-id="a6ac8-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a6ac8-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a6ac8-118">Member</span><span class="sxs-lookup"><span data-stu-id="a6ac8-118">Member</span></span> |
| [<span data-ttu-id="a6ac8-119">EventType</span><span class="sxs-lookup"><span data-stu-id="a6ac8-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="a6ac8-120">Member</span><span class="sxs-lookup"><span data-stu-id="a6ac8-120">Member</span></span> |
| [<span data-ttu-id="a6ac8-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a6ac8-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a6ac8-122">Membro</span><span class="sxs-lookup"><span data-stu-id="a6ac8-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a6ac8-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="a6ac8-123">Namespaces</span></span>

<span data-ttu-id="a6ac8-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a6ac8-125">[MailboxEnums](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-125">[MailboxEnums](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="a6ac8-126">Membros</span><span class="sxs-lookup"><span data-stu-id="a6ac8-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="a6ac8-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="a6ac8-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a6ac8-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6ac8-129">Type</span></span>

*   <span data-ttu-id="a6ac8-130">String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a6ac8-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a6ac8-131">Properties:</span></span>

|<span data-ttu-id="a6ac8-132">Nome</span><span class="sxs-lookup"><span data-stu-id="a6ac8-132">Name</span></span>| <span data-ttu-id="a6ac8-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6ac8-133">Type</span></span>| <span data-ttu-id="a6ac8-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6ac8-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a6ac8-135">String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-135">String</span></span>|<span data-ttu-id="a6ac8-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a6ac8-137">String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-137">String</span></span>|<span data-ttu-id="a6ac8-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6ac8-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6ac8-139">Requirements</span></span>

|<span data-ttu-id="a6ac8-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6ac8-140">Requirement</span></span>| <span data-ttu-id="a6ac8-141">Valor</span><span class="sxs-lookup"><span data-stu-id="a6ac8-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6ac8-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6ac8-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6ac8-143">1.0</span><span class="sxs-lookup"><span data-stu-id="a6ac8-143">1.0</span></span>|
|[<span data-ttu-id="a6ac8-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6ac8-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6ac8-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a6ac8-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="a6ac8-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-146">CoercionType :String</span></span>

<span data-ttu-id="a6ac8-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a6ac8-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6ac8-148">Type</span></span>

*   <span data-ttu-id="a6ac8-149">String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a6ac8-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a6ac8-150">Properties:</span></span>

|<span data-ttu-id="a6ac8-151">Nome</span><span class="sxs-lookup"><span data-stu-id="a6ac8-151">Name</span></span>| <span data-ttu-id="a6ac8-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6ac8-152">Type</span></span>| <span data-ttu-id="a6ac8-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6ac8-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a6ac8-154">String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-154">String</span></span>|<span data-ttu-id="a6ac8-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a6ac8-156">String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-156">String</span></span>|<span data-ttu-id="a6ac8-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6ac8-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6ac8-158">Requirements</span></span>

|<span data-ttu-id="a6ac8-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6ac8-159">Requirement</span></span>| <span data-ttu-id="a6ac8-160">Valor</span><span class="sxs-lookup"><span data-stu-id="a6ac8-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6ac8-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6ac8-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6ac8-162">1.0</span><span class="sxs-lookup"><span data-stu-id="a6ac8-162">1.0</span></span>|
|[<span data-ttu-id="a6ac8-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6ac8-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6ac8-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a6ac8-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="a6ac8-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-165">EventType :String</span></span>

<span data-ttu-id="a6ac8-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="a6ac8-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6ac8-167">Type</span></span>

*   <span data-ttu-id="a6ac8-168">String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a6ac8-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a6ac8-169">Properties:</span></span>

| <span data-ttu-id="a6ac8-170">Nome</span><span class="sxs-lookup"><span data-stu-id="a6ac8-170">Name</span></span> | <span data-ttu-id="a6ac8-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6ac8-171">Type</span></span> | <span data-ttu-id="a6ac8-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6ac8-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="a6ac8-173">String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-173">String</span></span> | <span data-ttu-id="a6ac8-174">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a6ac8-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6ac8-175">Requirements</span></span>

|<span data-ttu-id="a6ac8-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6ac8-176">Requirement</span></span>| <span data-ttu-id="a6ac8-177">Valor</span><span class="sxs-lookup"><span data-stu-id="a6ac8-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6ac8-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6ac8-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6ac8-179">1,5</span><span class="sxs-lookup"><span data-stu-id="a6ac8-179">1.5</span></span> |
|[<span data-ttu-id="a6ac8-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6ac8-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6ac8-181">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a6ac8-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="a6ac8-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-182">SourceProperty :String</span></span>

<span data-ttu-id="a6ac8-183">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a6ac8-184">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6ac8-184">Type</span></span>

*   <span data-ttu-id="a6ac8-185">String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a6ac8-186">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="a6ac8-186">Properties:</span></span>

|<span data-ttu-id="a6ac8-187">Nome</span><span class="sxs-lookup"><span data-stu-id="a6ac8-187">Name</span></span>| <span data-ttu-id="a6ac8-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="a6ac8-188">Type</span></span>| <span data-ttu-id="a6ac8-189">Descrição</span><span class="sxs-lookup"><span data-stu-id="a6ac8-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a6ac8-190">String</span><span class="sxs-lookup"><span data-stu-id="a6ac8-190">String</span></span>|<span data-ttu-id="a6ac8-191">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a6ac8-192">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="a6ac8-192">String</span></span>|<span data-ttu-id="a6ac8-193">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="a6ac8-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6ac8-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="a6ac8-194">Requirements</span></span>

|<span data-ttu-id="a6ac8-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="a6ac8-195">Requirement</span></span>| <span data-ttu-id="a6ac8-196">Valor</span><span class="sxs-lookup"><span data-stu-id="a6ac8-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6ac8-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="a6ac8-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6ac8-198">1.0</span><span class="sxs-lookup"><span data-stu-id="a6ac8-198">1.0</span></span>|
|[<span data-ttu-id="a6ac8-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="a6ac8-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6ac8-200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="a6ac8-200">Compose or Read</span></span>|
