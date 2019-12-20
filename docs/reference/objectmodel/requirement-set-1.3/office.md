---
title: Namespace do Office – conjunto de requisitos 1,3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 3c6ddc34001f4d1622bc76d9bca1fbde9425be8b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814896"
---
# <a name="office"></a><span data-ttu-id="77f35-102">Office</span><span class="sxs-lookup"><span data-stu-id="77f35-102">Office</span></span>

<span data-ttu-id="77f35-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="77f35-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="77f35-105">Requisitos</span><span class="sxs-lookup"><span data-stu-id="77f35-105">Requirements</span></span>

|<span data-ttu-id="77f35-106">Requisito</span><span class="sxs-lookup"><span data-stu-id="77f35-106">Requirement</span></span>| <span data-ttu-id="77f35-107">Valor</span><span class="sxs-lookup"><span data-stu-id="77f35-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="77f35-108">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="77f35-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="77f35-109">1.1</span><span class="sxs-lookup"><span data-stu-id="77f35-109">1.1</span></span>|
|[<span data-ttu-id="77f35-110">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="77f35-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="77f35-111">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="77f35-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="77f35-112">Propriedades</span><span class="sxs-lookup"><span data-stu-id="77f35-112">Properties</span></span>

| <span data-ttu-id="77f35-113">Propriedade</span><span class="sxs-lookup"><span data-stu-id="77f35-113">Property</span></span> | <span data-ttu-id="77f35-114">Modelos</span><span class="sxs-lookup"><span data-stu-id="77f35-114">Modes</span></span> | <span data-ttu-id="77f35-115">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="77f35-115">Return type</span></span> | <span data-ttu-id="77f35-116">Mínimo</span><span class="sxs-lookup"><span data-stu-id="77f35-116">Minimum</span></span><br><span data-ttu-id="77f35-117">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="77f35-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="77f35-118">context</span><span class="sxs-lookup"><span data-stu-id="77f35-118">context</span></span>](office.context.md) | <span data-ttu-id="77f35-119">Escrever</span><span class="sxs-lookup"><span data-stu-id="77f35-119">Compose</span></span><br><span data-ttu-id="77f35-120">Leitura</span><span class="sxs-lookup"><span data-stu-id="77f35-120">Read</span></span> | [<span data-ttu-id="77f35-121">Context</span><span class="sxs-lookup"><span data-stu-id="77f35-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="77f35-122">1.1</span><span class="sxs-lookup"><span data-stu-id="77f35-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="77f35-123">Enumerações</span><span class="sxs-lookup"><span data-stu-id="77f35-123">Enumerations</span></span>

| <span data-ttu-id="77f35-124">Enumeração</span><span class="sxs-lookup"><span data-stu-id="77f35-124">Enumeration</span></span> | <span data-ttu-id="77f35-125">Modelos</span><span class="sxs-lookup"><span data-stu-id="77f35-125">Modes</span></span> | <span data-ttu-id="77f35-126">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="77f35-126">Return type</span></span> | <span data-ttu-id="77f35-127">Mínimo</span><span class="sxs-lookup"><span data-stu-id="77f35-127">Minimum</span></span><br><span data-ttu-id="77f35-128">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="77f35-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="77f35-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="77f35-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="77f35-130">Escrever</span><span class="sxs-lookup"><span data-stu-id="77f35-130">Compose</span></span><br><span data-ttu-id="77f35-131">Leitura</span><span class="sxs-lookup"><span data-stu-id="77f35-131">Read</span></span> | <span data-ttu-id="77f35-132">String</span><span class="sxs-lookup"><span data-stu-id="77f35-132">String</span></span> | [<span data-ttu-id="77f35-133">1.1</span><span class="sxs-lookup"><span data-stu-id="77f35-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="77f35-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="77f35-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="77f35-135">Escrever</span><span class="sxs-lookup"><span data-stu-id="77f35-135">Compose</span></span><br><span data-ttu-id="77f35-136">Leitura</span><span class="sxs-lookup"><span data-stu-id="77f35-136">Read</span></span> | <span data-ttu-id="77f35-137">String</span><span class="sxs-lookup"><span data-stu-id="77f35-137">String</span></span> | [<span data-ttu-id="77f35-138">1.1</span><span class="sxs-lookup"><span data-stu-id="77f35-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="77f35-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="77f35-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="77f35-140">Escrever</span><span class="sxs-lookup"><span data-stu-id="77f35-140">Compose</span></span><br><span data-ttu-id="77f35-141">Leitura</span><span class="sxs-lookup"><span data-stu-id="77f35-141">Read</span></span> | <span data-ttu-id="77f35-142">String</span><span class="sxs-lookup"><span data-stu-id="77f35-142">String</span></span> | [<span data-ttu-id="77f35-143">1.1</span><span class="sxs-lookup"><span data-stu-id="77f35-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="77f35-144">Namespaces</span><span class="sxs-lookup"><span data-stu-id="77f35-144">Namespaces</span></span>

<span data-ttu-id="77f35-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): inclui uma série de enumerações específicas do Outlook, por exemplo, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,, e `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="77f35-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="77f35-146">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="77f35-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="77f35-147">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="77f35-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="77f35-148">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="77f35-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="77f35-149">Tipo</span><span class="sxs-lookup"><span data-stu-id="77f35-149">Type</span></span>

*   <span data-ttu-id="77f35-150">String</span><span class="sxs-lookup"><span data-stu-id="77f35-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="77f35-151">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="77f35-151">Properties:</span></span>

|<span data-ttu-id="77f35-152">Nome</span><span class="sxs-lookup"><span data-stu-id="77f35-152">Name</span></span>| <span data-ttu-id="77f35-153">Tipo</span><span class="sxs-lookup"><span data-stu-id="77f35-153">Type</span></span>| <span data-ttu-id="77f35-154">Descrição</span><span class="sxs-lookup"><span data-stu-id="77f35-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="77f35-155">String</span><span class="sxs-lookup"><span data-stu-id="77f35-155">String</span></span>|<span data-ttu-id="77f35-156">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="77f35-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="77f35-157">String</span><span class="sxs-lookup"><span data-stu-id="77f35-157">String</span></span>|<span data-ttu-id="77f35-158">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="77f35-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="77f35-159">Requisitos</span><span class="sxs-lookup"><span data-stu-id="77f35-159">Requirements</span></span>

|<span data-ttu-id="77f35-160">Requisito</span><span class="sxs-lookup"><span data-stu-id="77f35-160">Requirement</span></span>| <span data-ttu-id="77f35-161">Valor</span><span class="sxs-lookup"><span data-stu-id="77f35-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="77f35-162">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="77f35-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="77f35-163">1.1</span><span class="sxs-lookup"><span data-stu-id="77f35-163">1.1</span></span>|
|[<span data-ttu-id="77f35-164">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="77f35-164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="77f35-165">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="77f35-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="77f35-166">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="77f35-166">CoercionType: String</span></span>

<span data-ttu-id="77f35-167">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="77f35-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="77f35-168">Tipo</span><span class="sxs-lookup"><span data-stu-id="77f35-168">Type</span></span>

*   <span data-ttu-id="77f35-169">String</span><span class="sxs-lookup"><span data-stu-id="77f35-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="77f35-170">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="77f35-170">Properties:</span></span>

|<span data-ttu-id="77f35-171">Nome</span><span class="sxs-lookup"><span data-stu-id="77f35-171">Name</span></span>| <span data-ttu-id="77f35-172">Tipo</span><span class="sxs-lookup"><span data-stu-id="77f35-172">Type</span></span>| <span data-ttu-id="77f35-173">Descrição</span><span class="sxs-lookup"><span data-stu-id="77f35-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="77f35-174">String</span><span class="sxs-lookup"><span data-stu-id="77f35-174">String</span></span>|<span data-ttu-id="77f35-175">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="77f35-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="77f35-176">String</span><span class="sxs-lookup"><span data-stu-id="77f35-176">String</span></span>|<span data-ttu-id="77f35-177">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="77f35-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="77f35-178">Requisitos</span><span class="sxs-lookup"><span data-stu-id="77f35-178">Requirements</span></span>

|<span data-ttu-id="77f35-179">Requisito</span><span class="sxs-lookup"><span data-stu-id="77f35-179">Requirement</span></span>| <span data-ttu-id="77f35-180">Valor</span><span class="sxs-lookup"><span data-stu-id="77f35-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="77f35-181">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="77f35-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="77f35-182">1.1</span><span class="sxs-lookup"><span data-stu-id="77f35-182">1.1</span></span>|
|[<span data-ttu-id="77f35-183">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="77f35-183">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="77f35-184">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="77f35-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="77f35-185">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="77f35-185">SourceProperty: String</span></span>

<span data-ttu-id="77f35-186">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="77f35-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="77f35-187">Tipo</span><span class="sxs-lookup"><span data-stu-id="77f35-187">Type</span></span>

*   <span data-ttu-id="77f35-188">String</span><span class="sxs-lookup"><span data-stu-id="77f35-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="77f35-189">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="77f35-189">Properties:</span></span>

|<span data-ttu-id="77f35-190">Nome</span><span class="sxs-lookup"><span data-stu-id="77f35-190">Name</span></span>| <span data-ttu-id="77f35-191">Tipo</span><span class="sxs-lookup"><span data-stu-id="77f35-191">Type</span></span>| <span data-ttu-id="77f35-192">Descrição</span><span class="sxs-lookup"><span data-stu-id="77f35-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="77f35-193">String</span><span class="sxs-lookup"><span data-stu-id="77f35-193">String</span></span>|<span data-ttu-id="77f35-194">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="77f35-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="77f35-195">Cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="77f35-195">String</span></span>|<span data-ttu-id="77f35-196">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="77f35-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="77f35-197">Requisitos</span><span class="sxs-lookup"><span data-stu-id="77f35-197">Requirements</span></span>

|<span data-ttu-id="77f35-198">Requisito</span><span class="sxs-lookup"><span data-stu-id="77f35-198">Requirement</span></span>| <span data-ttu-id="77f35-199">Valor</span><span class="sxs-lookup"><span data-stu-id="77f35-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="77f35-200">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="77f35-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="77f35-201">1.1</span><span class="sxs-lookup"><span data-stu-id="77f35-201">1.1</span></span>|
|[<span data-ttu-id="77f35-202">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="77f35-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="77f35-203">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="77f35-203">Compose or Read</span></span>|
