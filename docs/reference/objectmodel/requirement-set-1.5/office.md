---
title: Namespace do Office – conjunto de requisitos 1,5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d8c51646818681629fa0c184962776beffe22a55
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871413"
---
# <a name="office"></a><span data-ttu-id="977b1-102">Office</span><span class="sxs-lookup"><span data-stu-id="977b1-102">Office</span></span>

<span data-ttu-id="977b1-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="977b1-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="977b1-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="977b1-105">Requirements</span></span>

|<span data-ttu-id="977b1-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="977b1-106">Requirement</span></span>| <span data-ttu-id="977b1-107">Valor</span><span class="sxs-lookup"><span data-stu-id="977b1-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="977b1-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="977b1-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="977b1-109">1.0</span><span class="sxs-lookup"><span data-stu-id="977b1-109">1.0</span></span>|
|[<span data-ttu-id="977b1-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="977b1-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="977b1-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="977b1-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="977b1-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="977b1-112">Members and methods</span></span>

| <span data-ttu-id="977b1-113">Membro</span><span class="sxs-lookup"><span data-stu-id="977b1-113">Member</span></span> | <span data-ttu-id="977b1-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="977b1-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="977b1-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="977b1-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="977b1-116">Member</span><span class="sxs-lookup"><span data-stu-id="977b1-116">Member</span></span> |
| [<span data-ttu-id="977b1-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="977b1-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="977b1-118">Member</span><span class="sxs-lookup"><span data-stu-id="977b1-118">Member</span></span> |
| [<span data-ttu-id="977b1-119">EventType</span><span class="sxs-lookup"><span data-stu-id="977b1-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="977b1-120">Member</span><span class="sxs-lookup"><span data-stu-id="977b1-120">Member</span></span> |
| [<span data-ttu-id="977b1-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="977b1-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="977b1-122">Membro</span><span class="sxs-lookup"><span data-stu-id="977b1-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="977b1-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="977b1-123">Namespaces</span></span>

<span data-ttu-id="977b1-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="977b1-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="977b1-125">[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype): inclui as enumerações ItemType, EntityType, AttachmentType, RecipientType, ResponseType e ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="977b1-125">[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="977b1-126">Membros</span><span class="sxs-lookup"><span data-stu-id="977b1-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="977b1-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="977b1-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="977b1-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="977b1-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="977b1-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="977b1-129">Type</span></span>

*   <span data-ttu-id="977b1-130">String</span><span class="sxs-lookup"><span data-stu-id="977b1-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="977b1-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="977b1-131">Properties:</span></span>

|<span data-ttu-id="977b1-132">Nome</span><span class="sxs-lookup"><span data-stu-id="977b1-132">Name</span></span>| <span data-ttu-id="977b1-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="977b1-133">Type</span></span>| <span data-ttu-id="977b1-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="977b1-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="977b1-135">String</span><span class="sxs-lookup"><span data-stu-id="977b1-135">String</span></span>|<span data-ttu-id="977b1-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="977b1-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="977b1-137">String</span><span class="sxs-lookup"><span data-stu-id="977b1-137">String</span></span>|<span data-ttu-id="977b1-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="977b1-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="977b1-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="977b1-139">Requirements</span></span>

|<span data-ttu-id="977b1-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="977b1-140">Requirement</span></span>| <span data-ttu-id="977b1-141">Valor</span><span class="sxs-lookup"><span data-stu-id="977b1-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="977b1-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="977b1-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="977b1-143">1.0</span><span class="sxs-lookup"><span data-stu-id="977b1-143">1.0</span></span>|
|[<span data-ttu-id="977b1-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="977b1-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="977b1-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="977b1-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="977b1-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="977b1-146">CoercionType :String</span></span>

<span data-ttu-id="977b1-147">Especifica como forçar os dados retornados ou definir de acordo com o método chamado.</span><span class="sxs-lookup"><span data-stu-id="977b1-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="977b1-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="977b1-148">Type</span></span>

*   <span data-ttu-id="977b1-149">String</span><span class="sxs-lookup"><span data-stu-id="977b1-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="977b1-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="977b1-150">Properties:</span></span>

|<span data-ttu-id="977b1-151">Nome</span><span class="sxs-lookup"><span data-stu-id="977b1-151">Name</span></span>| <span data-ttu-id="977b1-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="977b1-152">Type</span></span>| <span data-ttu-id="977b1-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="977b1-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="977b1-154">String</span><span class="sxs-lookup"><span data-stu-id="977b1-154">String</span></span>|<span data-ttu-id="977b1-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="977b1-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="977b1-156">String</span><span class="sxs-lookup"><span data-stu-id="977b1-156">String</span></span>|<span data-ttu-id="977b1-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="977b1-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="977b1-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="977b1-158">Requirements</span></span>

|<span data-ttu-id="977b1-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="977b1-159">Requirement</span></span>| <span data-ttu-id="977b1-160">Valor</span><span class="sxs-lookup"><span data-stu-id="977b1-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="977b1-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="977b1-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="977b1-162">1.0</span><span class="sxs-lookup"><span data-stu-id="977b1-162">1.0</span></span>|
|[<span data-ttu-id="977b1-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="977b1-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="977b1-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="977b1-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="977b1-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="977b1-165">EventType :String</span></span>

<span data-ttu-id="977b1-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="977b1-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="977b1-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="977b1-167">Type</span></span>

*   <span data-ttu-id="977b1-168">String</span><span class="sxs-lookup"><span data-stu-id="977b1-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="977b1-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="977b1-169">Properties:</span></span>

| <span data-ttu-id="977b1-170">Nome</span><span class="sxs-lookup"><span data-stu-id="977b1-170">Name</span></span> | <span data-ttu-id="977b1-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="977b1-171">Type</span></span> | <span data-ttu-id="977b1-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="977b1-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="977b1-173">String</span><span class="sxs-lookup"><span data-stu-id="977b1-173">String</span></span> | <span data-ttu-id="977b1-174">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="977b1-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="977b1-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="977b1-175">Requirements</span></span>

|<span data-ttu-id="977b1-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="977b1-176">Requirement</span></span>| <span data-ttu-id="977b1-177">Valor</span><span class="sxs-lookup"><span data-stu-id="977b1-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="977b1-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="977b1-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="977b1-179">1,5</span><span class="sxs-lookup"><span data-stu-id="977b1-179">1.5</span></span> |
|[<span data-ttu-id="977b1-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="977b1-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="977b1-181">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="977b1-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="977b1-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="977b1-182">SourceProperty :String</span></span>

<span data-ttu-id="977b1-183">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="977b1-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="977b1-184">Tipo</span><span class="sxs-lookup"><span data-stu-id="977b1-184">Type</span></span>

*   <span data-ttu-id="977b1-185">String</span><span class="sxs-lookup"><span data-stu-id="977b1-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="977b1-186">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="977b1-186">Properties:</span></span>

|<span data-ttu-id="977b1-187">Nome</span><span class="sxs-lookup"><span data-stu-id="977b1-187">Name</span></span>| <span data-ttu-id="977b1-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="977b1-188">Type</span></span>| <span data-ttu-id="977b1-189">Descrição</span><span class="sxs-lookup"><span data-stu-id="977b1-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="977b1-190">String</span><span class="sxs-lookup"><span data-stu-id="977b1-190">String</span></span>|<span data-ttu-id="977b1-191">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="977b1-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="977b1-192">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="977b1-192">String</span></span>|<span data-ttu-id="977b1-193">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="977b1-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="977b1-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="977b1-194">Requirements</span></span>

|<span data-ttu-id="977b1-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="977b1-195">Requirement</span></span>| <span data-ttu-id="977b1-196">Valor</span><span class="sxs-lookup"><span data-stu-id="977b1-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="977b1-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="977b1-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="977b1-198">1.0</span><span class="sxs-lookup"><span data-stu-id="977b1-198">1.0</span></span>|
|[<span data-ttu-id="977b1-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="977b1-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="977b1-200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="977b1-200">Compose or Read</span></span>|
