---
title: Namespace do Office – conjunto de requisitos 1,4
description: Membros de namespace do Office disponíveis para suplementos do Outlook usando o conjunto de requisitos de API da caixa de correio 1,4.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: f797fe5281d2031a2182249aeb18d740cd114d43
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430740"
---
# <a name="office-mailbox-requirement-set-14"></a><span data-ttu-id="71794-103">Office (conjunto de requisitos de caixa de correio 1,4)</span><span class="sxs-lookup"><span data-stu-id="71794-103">Office (Mailbox requirement set 1.4)</span></span>

<span data-ttu-id="71794-p101">O namespace do Office fornece interfaces compartilhadas que são usadas pelos suplementos em todos os aplicativos do Office. Esta listagem documenta somente as interfaces que são usadas pelos suplementos do Outlook. Para obter uma lista completa de namespaces do Office, confira [API compartilhada](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="71794-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="71794-106">Requisitos</span><span class="sxs-lookup"><span data-stu-id="71794-106">Requirements</span></span>

|<span data-ttu-id="71794-107">Requisito</span><span class="sxs-lookup"><span data-stu-id="71794-107">Requirement</span></span>| <span data-ttu-id="71794-108">Valor</span><span class="sxs-lookup"><span data-stu-id="71794-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="71794-109">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="71794-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="71794-110">1.1</span><span class="sxs-lookup"><span data-stu-id="71794-110">1.1</span></span>|
|[<span data-ttu-id="71794-111">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="71794-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="71794-112">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="71794-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="71794-113">Propriedades</span><span class="sxs-lookup"><span data-stu-id="71794-113">Properties</span></span>

| <span data-ttu-id="71794-114">Propriedade</span><span class="sxs-lookup"><span data-stu-id="71794-114">Property</span></span> | <span data-ttu-id="71794-115">Modelos</span><span class="sxs-lookup"><span data-stu-id="71794-115">Modes</span></span> | <span data-ttu-id="71794-116">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="71794-116">Return type</span></span> | <span data-ttu-id="71794-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="71794-117">Minimum</span></span><br><span data-ttu-id="71794-118">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="71794-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="71794-119">context</span><span class="sxs-lookup"><span data-stu-id="71794-119">context</span></span>](office.context.md) | <span data-ttu-id="71794-120">Escrever</span><span class="sxs-lookup"><span data-stu-id="71794-120">Compose</span></span><br><span data-ttu-id="71794-121">Ler</span><span class="sxs-lookup"><span data-stu-id="71794-121">Read</span></span> | [<span data-ttu-id="71794-122">Context</span><span class="sxs-lookup"><span data-stu-id="71794-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="71794-123">1.1</span><span class="sxs-lookup"><span data-stu-id="71794-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="71794-124">Enumerações</span><span class="sxs-lookup"><span data-stu-id="71794-124">Enumerations</span></span>

| <span data-ttu-id="71794-125">Enumeração</span><span class="sxs-lookup"><span data-stu-id="71794-125">Enumeration</span></span> | <span data-ttu-id="71794-126">Modelos</span><span class="sxs-lookup"><span data-stu-id="71794-126">Modes</span></span> | <span data-ttu-id="71794-127">Tipo de retorno</span><span class="sxs-lookup"><span data-stu-id="71794-127">Return type</span></span> | <span data-ttu-id="71794-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="71794-128">Minimum</span></span><br><span data-ttu-id="71794-129">conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="71794-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="71794-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="71794-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="71794-131">Escrever</span><span class="sxs-lookup"><span data-stu-id="71794-131">Compose</span></span><br><span data-ttu-id="71794-132">Ler</span><span class="sxs-lookup"><span data-stu-id="71794-132">Read</span></span> | <span data-ttu-id="71794-133">String</span><span class="sxs-lookup"><span data-stu-id="71794-133">String</span></span> | [<span data-ttu-id="71794-134">1.1</span><span class="sxs-lookup"><span data-stu-id="71794-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="71794-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="71794-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="71794-136">Escrever</span><span class="sxs-lookup"><span data-stu-id="71794-136">Compose</span></span><br><span data-ttu-id="71794-137">Ler</span><span class="sxs-lookup"><span data-stu-id="71794-137">Read</span></span> | <span data-ttu-id="71794-138">String</span><span class="sxs-lookup"><span data-stu-id="71794-138">String</span></span> | [<span data-ttu-id="71794-139">1.1</span><span class="sxs-lookup"><span data-stu-id="71794-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="71794-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="71794-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="71794-141">Escrever</span><span class="sxs-lookup"><span data-stu-id="71794-141">Compose</span></span><br><span data-ttu-id="71794-142">Ler</span><span class="sxs-lookup"><span data-stu-id="71794-142">Read</span></span> | <span data-ttu-id="71794-143">String</span><span class="sxs-lookup"><span data-stu-id="71794-143">String</span></span> | [<span data-ttu-id="71794-144">1.1</span><span class="sxs-lookup"><span data-stu-id="71794-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="71794-145">Namespaces</span><span class="sxs-lookup"><span data-stu-id="71794-145">Namespaces</span></span>

<span data-ttu-id="71794-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true): inclui uma série de enumerações específicas do Outlook, por exemplo,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` e `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="71794-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="71794-147">Detalhes da enumeração</span><span class="sxs-lookup"><span data-stu-id="71794-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="71794-148">AsyncResultStatus: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="71794-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="71794-149">Especifica o resultado de uma chamada assíncrona.</span><span class="sxs-lookup"><span data-stu-id="71794-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="71794-150">Tipo</span><span class="sxs-lookup"><span data-stu-id="71794-150">Type</span></span>

*   <span data-ttu-id="71794-151">String</span><span class="sxs-lookup"><span data-stu-id="71794-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="71794-152">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="71794-152">Properties:</span></span>

|<span data-ttu-id="71794-153">Nome</span><span class="sxs-lookup"><span data-stu-id="71794-153">Name</span></span>| <span data-ttu-id="71794-154">Tipo</span><span class="sxs-lookup"><span data-stu-id="71794-154">Type</span></span>| <span data-ttu-id="71794-155">Descrição</span><span class="sxs-lookup"><span data-stu-id="71794-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="71794-156">String</span><span class="sxs-lookup"><span data-stu-id="71794-156">String</span></span>|<span data-ttu-id="71794-157">A chamada foi bem-sucedida.</span><span class="sxs-lookup"><span data-stu-id="71794-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="71794-158">String</span><span class="sxs-lookup"><span data-stu-id="71794-158">String</span></span>|<span data-ttu-id="71794-159">Falha na chamada.</span><span class="sxs-lookup"><span data-stu-id="71794-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71794-160">Requisitos</span><span class="sxs-lookup"><span data-stu-id="71794-160">Requirements</span></span>

|<span data-ttu-id="71794-161">Requisito</span><span class="sxs-lookup"><span data-stu-id="71794-161">Requirement</span></span>| <span data-ttu-id="71794-162">Valor</span><span class="sxs-lookup"><span data-stu-id="71794-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="71794-163">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="71794-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="71794-164">1.1</span><span class="sxs-lookup"><span data-stu-id="71794-164">1.1</span></span>|
|[<span data-ttu-id="71794-165">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="71794-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="71794-166">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="71794-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="71794-167">CoercionType: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="71794-167">CoercionType: String</span></span>

<span data-ttu-id="71794-168">Especifica como forçar dados retornados ou definidos pelo método invocado.</span><span class="sxs-lookup"><span data-stu-id="71794-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="71794-169">Tipo</span><span class="sxs-lookup"><span data-stu-id="71794-169">Type</span></span>

*   <span data-ttu-id="71794-170">String</span><span class="sxs-lookup"><span data-stu-id="71794-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="71794-171">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="71794-171">Properties:</span></span>

|<span data-ttu-id="71794-172">Nome</span><span class="sxs-lookup"><span data-stu-id="71794-172">Name</span></span>| <span data-ttu-id="71794-173">Tipo</span><span class="sxs-lookup"><span data-stu-id="71794-173">Type</span></span>| <span data-ttu-id="71794-174">Descrição</span><span class="sxs-lookup"><span data-stu-id="71794-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="71794-175">String</span><span class="sxs-lookup"><span data-stu-id="71794-175">String</span></span>|<span data-ttu-id="71794-176">Solicita que os dados sejam retornados no formato HTML.</span><span class="sxs-lookup"><span data-stu-id="71794-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="71794-177">String</span><span class="sxs-lookup"><span data-stu-id="71794-177">String</span></span>|<span data-ttu-id="71794-178">Solicita que os dados sejam retornados no formato de texto.</span><span class="sxs-lookup"><span data-stu-id="71794-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71794-179">Requisitos</span><span class="sxs-lookup"><span data-stu-id="71794-179">Requirements</span></span>

|<span data-ttu-id="71794-180">Requisito</span><span class="sxs-lookup"><span data-stu-id="71794-180">Requirement</span></span>| <span data-ttu-id="71794-181">Valor</span><span class="sxs-lookup"><span data-stu-id="71794-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="71794-182">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="71794-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="71794-183">1.1</span><span class="sxs-lookup"><span data-stu-id="71794-183">1.1</span></span>|
|[<span data-ttu-id="71794-184">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="71794-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="71794-185">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="71794-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="71794-186">SourceProperty: cadeia de caracteres</span><span class="sxs-lookup"><span data-stu-id="71794-186">SourceProperty: String</span></span>

<span data-ttu-id="71794-187">Especifica a origem dos dados retornados pelo método chamado.</span><span class="sxs-lookup"><span data-stu-id="71794-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="71794-188">Tipo</span><span class="sxs-lookup"><span data-stu-id="71794-188">Type</span></span>

*   <span data-ttu-id="71794-189">String</span><span class="sxs-lookup"><span data-stu-id="71794-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="71794-190">Propriedades:</span><span class="sxs-lookup"><span data-stu-id="71794-190">Properties:</span></span>

|<span data-ttu-id="71794-191">Nome</span><span class="sxs-lookup"><span data-stu-id="71794-191">Name</span></span>| <span data-ttu-id="71794-192">Tipo</span><span class="sxs-lookup"><span data-stu-id="71794-192">Type</span></span>| <span data-ttu-id="71794-193">Descrição</span><span class="sxs-lookup"><span data-stu-id="71794-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="71794-194">String</span><span class="sxs-lookup"><span data-stu-id="71794-194">String</span></span>|<span data-ttu-id="71794-195">A origem dos dados é o corpo de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="71794-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="71794-196">String</span><span class="sxs-lookup"><span data-stu-id="71794-196">String</span></span>|<span data-ttu-id="71794-197">A origem dos dados é o assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="71794-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="71794-198">Requisitos</span><span class="sxs-lookup"><span data-stu-id="71794-198">Requirements</span></span>

|<span data-ttu-id="71794-199">Requisito</span><span class="sxs-lookup"><span data-stu-id="71794-199">Requirement</span></span>| <span data-ttu-id="71794-200">Valor</span><span class="sxs-lookup"><span data-stu-id="71794-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="71794-201">Versão do conjunto de requisitos mínimos da caixa de correio</span><span class="sxs-lookup"><span data-stu-id="71794-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="71794-202">1.1</span><span class="sxs-lookup"><span data-stu-id="71794-202">1.1</span></span>|
|[<span data-ttu-id="71794-203">Modo do Outlook aplicável</span><span class="sxs-lookup"><span data-stu-id="71794-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="71794-204">Escrever ou Ler</span><span class="sxs-lookup"><span data-stu-id="71794-204">Compose or Read</span></span>|
