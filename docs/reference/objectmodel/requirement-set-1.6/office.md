---
title: Namespace do Office – conjunto de requisitos 1,6
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 84e8fa49e1d4dce4239525badafaa051325bb3ec
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395635"
---
# <a name="office"></a><span data-ttu-id="1e948-102">Office</span><span class="sxs-lookup"><span data-stu-id="1e948-102">Office</span></span>

<span data-ttu-id="1e948-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="1e948-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e948-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1e948-105">Requirements</span></span>

|<span data-ttu-id="1e948-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="1e948-106">Requirement</span></span>| <span data-ttu-id="1e948-107">Valor</span><span class="sxs-lookup"><span data-stu-id="1e948-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e948-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1e948-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e948-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1e948-109">1.0</span></span>|
|[<span data-ttu-id="1e948-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1e948-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1e948-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1e948-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1e948-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="1e948-112">Members and methods</span></span>

| <span data-ttu-id="1e948-113">Membro</span><span class="sxs-lookup"><span data-stu-id="1e948-113">Member</span></span> | <span data-ttu-id="1e948-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="1e948-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1e948-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="1e948-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="1e948-116">Membro</span><span class="sxs-lookup"><span data-stu-id="1e948-116">Member</span></span> |
| [<span data-ttu-id="1e948-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="1e948-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="1e948-118">Membro</span><span class="sxs-lookup"><span data-stu-id="1e948-118">Member</span></span> |
| [<span data-ttu-id="1e948-119">EventType</span><span class="sxs-lookup"><span data-stu-id="1e948-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="1e948-120">Membro</span><span class="sxs-lookup"><span data-stu-id="1e948-120">Member</span></span> |
| [<span data-ttu-id="1e948-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="1e948-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="1e948-122">Membro</span><span class="sxs-lookup"><span data-stu-id="1e948-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="1e948-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="1e948-123">Namespaces</span></span>

<span data-ttu-id="1e948-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="1e948-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="1e948-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): inclui um número de enumerações, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="1e948-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="1e948-126">Members</span><span class="sxs-lookup"><span data-stu-id="1e948-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="1e948-127">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1e948-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="1e948-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="1e948-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="1e948-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="1e948-129">Type</span></span>

*   <span data-ttu-id="1e948-130">String</span><span class="sxs-lookup"><span data-stu-id="1e948-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1e948-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1e948-131">Properties:</span></span>

|<span data-ttu-id="1e948-132">Nome</span><span class="sxs-lookup"><span data-stu-id="1e948-132">Name</span></span>| <span data-ttu-id="1e948-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="1e948-133">Type</span></span>| <span data-ttu-id="1e948-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="1e948-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="1e948-135">String</span><span class="sxs-lookup"><span data-stu-id="1e948-135">String</span></span>|<span data-ttu-id="1e948-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="1e948-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="1e948-137">String</span><span class="sxs-lookup"><span data-stu-id="1e948-137">String</span></span>|<span data-ttu-id="1e948-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="1e948-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e948-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1e948-139">Requirements</span></span>

|<span data-ttu-id="1e948-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="1e948-140">Requirement</span></span>| <span data-ttu-id="1e948-141">Valor</span><span class="sxs-lookup"><span data-stu-id="1e948-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e948-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1e948-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e948-143">1.0</span><span class="sxs-lookup"><span data-stu-id="1e948-143">1.0</span></span>|
|[<span data-ttu-id="1e948-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1e948-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1e948-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1e948-145">Compose or Read</span></span>|

---

#### <a name="coerciontype-string"></a><span data-ttu-id="1e948-146">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1e948-146">CoercionType: String</span></span>

<span data-ttu-id="1e948-147">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="1e948-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1e948-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="1e948-148">Type</span></span>

*   <span data-ttu-id="1e948-149">String</span><span class="sxs-lookup"><span data-stu-id="1e948-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1e948-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1e948-150">Properties:</span></span>

|<span data-ttu-id="1e948-151">Nome</span><span class="sxs-lookup"><span data-stu-id="1e948-151">Name</span></span>| <span data-ttu-id="1e948-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="1e948-152">Type</span></span>| <span data-ttu-id="1e948-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="1e948-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="1e948-154">String</span><span class="sxs-lookup"><span data-stu-id="1e948-154">String</span></span>|<span data-ttu-id="1e948-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="1e948-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="1e948-156">String</span><span class="sxs-lookup"><span data-stu-id="1e948-156">String</span></span>|<span data-ttu-id="1e948-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="1e948-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e948-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1e948-158">Requirements</span></span>

|<span data-ttu-id="1e948-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="1e948-159">Requirement</span></span>| <span data-ttu-id="1e948-160">Valor</span><span class="sxs-lookup"><span data-stu-id="1e948-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e948-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1e948-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e948-162">1.0</span><span class="sxs-lookup"><span data-stu-id="1e948-162">1.0</span></span>|
|[<span data-ttu-id="1e948-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1e948-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1e948-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1e948-164">Compose or Read</span></span>|

---

#### <a name="eventtype-string"></a><span data-ttu-id="1e948-165">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1e948-165">EventType: String</span></span>

<span data-ttu-id="1e948-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="1e948-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="1e948-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="1e948-167">Type</span></span>

*   <span data-ttu-id="1e948-168">String</span><span class="sxs-lookup"><span data-stu-id="1e948-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1e948-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1e948-169">Properties:</span></span>

| <span data-ttu-id="1e948-170">Nome</span><span class="sxs-lookup"><span data-stu-id="1e948-170">Name</span></span> | <span data-ttu-id="1e948-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="1e948-171">Type</span></span> | <span data-ttu-id="1e948-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="1e948-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="1e948-173">String</span><span class="sxs-lookup"><span data-stu-id="1e948-173">String</span></span> | <span data-ttu-id="1e948-174">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="1e948-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1e948-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1e948-175">Requirements</span></span>

|<span data-ttu-id="1e948-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="1e948-176">Requirement</span></span>| <span data-ttu-id="1e948-177">Valor</span><span class="sxs-lookup"><span data-stu-id="1e948-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e948-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1e948-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e948-179">1,5</span><span class="sxs-lookup"><span data-stu-id="1e948-179">1.5</span></span> |
|[<span data-ttu-id="1e948-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1e948-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1e948-181">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1e948-181">Compose or Read</span></span> |

---

#### <a name="sourceproperty-string"></a><span data-ttu-id="1e948-182">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1e948-182">SourceProperty: String</span></span>

<span data-ttu-id="1e948-183">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="1e948-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1e948-184">Tipo</span><span class="sxs-lookup"><span data-stu-id="1e948-184">Type</span></span>

*   <span data-ttu-id="1e948-185">String</span><span class="sxs-lookup"><span data-stu-id="1e948-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1e948-186">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="1e948-186">Properties:</span></span>

|<span data-ttu-id="1e948-187">Nome</span><span class="sxs-lookup"><span data-stu-id="1e948-187">Name</span></span>| <span data-ttu-id="1e948-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="1e948-188">Type</span></span>| <span data-ttu-id="1e948-189">Descrição</span><span class="sxs-lookup"><span data-stu-id="1e948-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="1e948-190">String</span><span class="sxs-lookup"><span data-stu-id="1e948-190">String</span></span>|<span data-ttu-id="1e948-191">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1e948-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="1e948-192">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="1e948-192">String</span></span>|<span data-ttu-id="1e948-193">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="1e948-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1e948-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="1e948-194">Requirements</span></span>

|<span data-ttu-id="1e948-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="1e948-195">Requirement</span></span>| <span data-ttu-id="1e948-196">Valor</span><span class="sxs-lookup"><span data-stu-id="1e948-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e948-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="1e948-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e948-198">1.0</span><span class="sxs-lookup"><span data-stu-id="1e948-198">1.0</span></span>|
|[<span data-ttu-id="1e948-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="1e948-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1e948-200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="1e948-200">Compose or Read</span></span>|
