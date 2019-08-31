---
title: Namespace do Office – conjunto de requisitos 1,6
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: ae764e8cda2b3f14e33b883d054379db7b37a687
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695998"
---
# <a name="office"></a><span data-ttu-id="09bb5-102">Office</span><span class="sxs-lookup"><span data-stu-id="09bb5-102">Office</span></span>

<span data-ttu-id="09bb5-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="09bb5-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="09bb5-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb5-105">Requirements</span></span>

|<span data-ttu-id="09bb5-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb5-106">Requirement</span></span>| <span data-ttu-id="09bb5-107">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb5-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb5-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb5-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09bb5-109">1.0</span><span class="sxs-lookup"><span data-stu-id="09bb5-109">1.0</span></span>|
|[<span data-ttu-id="09bb5-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb5-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb5-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb5-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="09bb5-112">Membros e métodos</span><span class="sxs-lookup"><span data-stu-id="09bb5-112">Members and methods</span></span>

| <span data-ttu-id="09bb5-113">Membro</span><span class="sxs-lookup"><span data-stu-id="09bb5-113">Member</span></span> | <span data-ttu-id="09bb5-114">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb5-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="09bb5-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="09bb5-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="09bb5-116">Membro</span><span class="sxs-lookup"><span data-stu-id="09bb5-116">Member</span></span> |
| [<span data-ttu-id="09bb5-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="09bb5-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="09bb5-118">Membro</span><span class="sxs-lookup"><span data-stu-id="09bb5-118">Member</span></span> |
| [<span data-ttu-id="09bb5-119">EventType</span><span class="sxs-lookup"><span data-stu-id="09bb5-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="09bb5-120">Membro</span><span class="sxs-lookup"><span data-stu-id="09bb5-120">Member</span></span> |
| [<span data-ttu-id="09bb5-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="09bb5-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="09bb5-122">Membro</span><span class="sxs-lookup"><span data-stu-id="09bb5-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="09bb5-123">Namespaces</span><span class="sxs-lookup"><span data-stu-id="09bb5-123">Namespaces</span></span>

<span data-ttu-id="09bb5-124">[context](office.context.md): fornece interfaces compartilhadas do namespace de contexto da API de Suplementos do Office para uso na API de suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="09bb5-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="09bb5-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): inclui um número de enumerações, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="09bb5-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="09bb5-126">Members</span><span class="sxs-lookup"><span data-stu-id="09bb5-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="09bb5-127">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="09bb5-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="09bb5-128">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="09bb5-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb5-129">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb5-129">Type</span></span>

*   <span data-ttu-id="09bb5-130">String</span><span class="sxs-lookup"><span data-stu-id="09bb5-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="09bb5-131">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="09bb5-131">Properties:</span></span>

|<span data-ttu-id="09bb5-132">Nome</span><span class="sxs-lookup"><span data-stu-id="09bb5-132">Name</span></span>| <span data-ttu-id="09bb5-133">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb5-133">Type</span></span>| <span data-ttu-id="09bb5-134">Descrição</span><span class="sxs-lookup"><span data-stu-id="09bb5-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="09bb5-135">String</span><span class="sxs-lookup"><span data-stu-id="09bb5-135">String</span></span>|<span data-ttu-id="09bb5-136">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="09bb5-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="09bb5-137">String</span><span class="sxs-lookup"><span data-stu-id="09bb5-137">String</span></span>|<span data-ttu-id="09bb5-138">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="09bb5-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09bb5-139">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb5-139">Requirements</span></span>

|<span data-ttu-id="09bb5-140">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb5-140">Requirement</span></span>| <span data-ttu-id="09bb5-141">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb5-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb5-142">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb5-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09bb5-143">1.0</span><span class="sxs-lookup"><span data-stu-id="09bb5-143">1.0</span></span>|
|[<span data-ttu-id="09bb5-144">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb5-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb5-145">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb5-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="09bb5-146">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="09bb5-146">CoercionType: String</span></span>

<span data-ttu-id="09bb5-147">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="09bb5-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb5-148">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb5-148">Type</span></span>

*   <span data-ttu-id="09bb5-149">String</span><span class="sxs-lookup"><span data-stu-id="09bb5-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="09bb5-150">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="09bb5-150">Properties:</span></span>

|<span data-ttu-id="09bb5-151">Nome</span><span class="sxs-lookup"><span data-stu-id="09bb5-151">Name</span></span>| <span data-ttu-id="09bb5-152">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb5-152">Type</span></span>| <span data-ttu-id="09bb5-153">Descrição</span><span class="sxs-lookup"><span data-stu-id="09bb5-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="09bb5-154">String</span><span class="sxs-lookup"><span data-stu-id="09bb5-154">String</span></span>|<span data-ttu-id="09bb5-155">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="09bb5-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="09bb5-156">String</span><span class="sxs-lookup"><span data-stu-id="09bb5-156">String</span></span>|<span data-ttu-id="09bb5-157">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="09bb5-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09bb5-158">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb5-158">Requirements</span></span>

|<span data-ttu-id="09bb5-159">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb5-159">Requirement</span></span>| <span data-ttu-id="09bb5-160">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb5-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb5-161">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb5-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09bb5-162">1.0</span><span class="sxs-lookup"><span data-stu-id="09bb5-162">1.0</span></span>|
|[<span data-ttu-id="09bb5-163">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb5-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb5-164">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb5-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="09bb5-165">EventType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="09bb5-165">EventType: String</span></span>

<span data-ttu-id="09bb5-166">Especifica o evento associado a um manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="09bb5-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb5-167">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb5-167">Type</span></span>

*   <span data-ttu-id="09bb5-168">String</span><span class="sxs-lookup"><span data-stu-id="09bb5-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="09bb5-169">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="09bb5-169">Properties:</span></span>

| <span data-ttu-id="09bb5-170">Nome</span><span class="sxs-lookup"><span data-stu-id="09bb5-170">Name</span></span> | <span data-ttu-id="09bb5-171">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb5-171">Type</span></span> | <span data-ttu-id="09bb5-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="09bb5-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="09bb5-173">String</span><span class="sxs-lookup"><span data-stu-id="09bb5-173">String</span></span> | <span data-ttu-id="09bb5-174">Um item diferente do Outlook é selecionado para exibição enquanto o painel de tarefas está fixado.</span><span class="sxs-lookup"><span data-stu-id="09bb5-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="09bb5-175">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb5-175">Requirements</span></span>

|<span data-ttu-id="09bb5-176">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb5-176">Requirement</span></span>| <span data-ttu-id="09bb5-177">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb5-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb5-178">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb5-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09bb5-179">1,5</span><span class="sxs-lookup"><span data-stu-id="09bb5-179">1.5</span></span> |
|[<span data-ttu-id="09bb5-180">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb5-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb5-181">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb5-181">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="09bb5-182">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="09bb5-182">SourceProperty: String</span></span>

<span data-ttu-id="09bb5-183">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="09bb5-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="09bb5-184">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb5-184">Type</span></span>

*   <span data-ttu-id="09bb5-185">String</span><span class="sxs-lookup"><span data-stu-id="09bb5-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="09bb5-186">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="09bb5-186">Properties:</span></span>

|<span data-ttu-id="09bb5-187">Nome</span><span class="sxs-lookup"><span data-stu-id="09bb5-187">Name</span></span>| <span data-ttu-id="09bb5-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="09bb5-188">Type</span></span>| <span data-ttu-id="09bb5-189">Descrição</span><span class="sxs-lookup"><span data-stu-id="09bb5-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="09bb5-190">String</span><span class="sxs-lookup"><span data-stu-id="09bb5-190">String</span></span>|<span data-ttu-id="09bb5-191">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="09bb5-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="09bb5-192">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="09bb5-192">String</span></span>|<span data-ttu-id="09bb5-193">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="09bb5-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09bb5-194">Requisitos</span><span class="sxs-lookup"><span data-stu-id="09bb5-194">Requirements</span></span>

|<span data-ttu-id="09bb5-195">Requisito</span><span class="sxs-lookup"><span data-stu-id="09bb5-195">Requirement</span></span>| <span data-ttu-id="09bb5-196">Valor</span><span class="sxs-lookup"><span data-stu-id="09bb5-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="09bb5-197">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="09bb5-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09bb5-198">1.0</span><span class="sxs-lookup"><span data-stu-id="09bb5-198">1.0</span></span>|
|[<span data-ttu-id="09bb5-199">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="09bb5-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="09bb5-200">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="09bb5-200">Compose or Read</span></span>|
