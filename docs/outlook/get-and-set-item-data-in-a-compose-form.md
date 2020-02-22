---
title: Obter e definir dados de item em um formulário de composição no Outlook
description: Obtenha ou defina várias propriedades de um item em um suplemento do Outlook em um cenário de redação, incluindo seus destinatários, o assunto, o corpo e o local e a hora do compromisso.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: ff75c6565b6ff49dfb2ad1ac95c75499c9b32284
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165831"
---
# <a name="get-and-set-item-data-in-a-compose-form-in-outlook"></a><span data-ttu-id="8c85f-103">Obter e definir dados de item em um formulário de redação no Outlook</span><span class="sxs-lookup"><span data-stu-id="8c85f-103">Get and set item data in a compose form in Outlook</span></span>

<span data-ttu-id="8c85f-104">Saiba como obter ou definir várias propriedades de um item em um suplemento do Outlook em um cenário de composição, incluindo seus destinatários, o assunto, o corpo e o local e a hora do compromisso.</span><span class="sxs-lookup"><span data-stu-id="8c85f-104">Learn how to get or set various properties of an item in an Outlook add-in in a compose scenario, including its recipients, subject, body, and appointment location and time.</span></span>

## <a name="getting-and-setting-item-properties-for-a-compose-add-in"></a><span data-ttu-id="8c85f-105">Obter e definir propriedades de item de um suplemento de redação</span><span class="sxs-lookup"><span data-stu-id="8c85f-105">Getting and setting item properties for a compose add-in</span></span>

<span data-ttu-id="8c85f-106">Em um formulário de composição, é possível obter a maioria das propriedades que estão expostas no mesmo tipo de item de um formulário de leitura (por exemplo, participantes, destinatários, assunto e corpo) e acessar algumas propriedades adicionais que são relevantes somente no formulário de composição, mas não de leitura (corpo, cco).</span><span class="sxs-lookup"><span data-stu-id="8c85f-106">In a compose form, you can get most of the properties that are exposed on the same kind of item as in a read form (such as attendees, recipients, subject, and body), and you can get a few extra properties that are relevant in only a compose form but not a read form (body, bcc).</span></span>

<span data-ttu-id="8c85f-p101">Para a maioria dessas propriedades, como é possível que um suplemento do Outlook e o usuário estejam modificando a mesma propriedade na interface de usuário ao mesmo tempo, os métodos para obtê-las e defini-las é assíncrono. A Tabela 1 lista as propriedades no nível do item e os métodos assíncronos correspondentes para obtê-los e defini-los em um formulário de redação. As propriedades [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) e [item.conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) são exceções porque os usuários não podem modificá-las. Você pode obtê-las via programação da mesma maneira em um formulário de redação e em um formulário de leitura, diretamente do objeto pai.</span><span class="sxs-lookup"><span data-stu-id="8c85f-p101">For most of these properties, because it's possible that an Outlook add-in and the user can be modifying the same property in the user interface at the same time, the methods to get and set them are asynchronous. Table 1 lists the item-level properties and corresponding asynchronous methods to get and set them in a compose form. The  [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [item.conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) properties are exceptions because users cannot modify them. You can programmatically get them the same way in a compose form as in a read form, directly from the parent object.</span></span>

<span data-ttu-id="8c85f-111">Em vez de acessar as propriedades do item da API JavaScript para Office, você pode acessar as propriedades no nível do item usando os EWS (Serviços Web do Exchange).</span><span class="sxs-lookup"><span data-stu-id="8c85f-111">Other than accessing item properties in the JavaScript API for Office, you can access item-level properties using Exchange Web Services (EWS).</span></span> <span data-ttu-id="8c85f-112">Com a permissão **ReadWriteMailbox**, você pode usar o método [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) para acessar as operações [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) e [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) dos EWS para obter e definir propriedades de um ou mais itens na caixa de correio do usuário.</span><span class="sxs-lookup"><span data-stu-id="8c85f-112">With the **ReadWriteMailbox** permission, you can use the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to access EWS operations, [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) and [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation), to get and set more properties of an item or items in the user's mailbox.</span></span>

<span data-ttu-id="8c85f-113">A função `makeEwsRequestAsync` está disponível nos formulários de leitura e redação.</span><span class="sxs-lookup"><span data-stu-id="8c85f-113">The `makeEwsRequestAsync` function is available in both compose and read forms.</span></span> <span data-ttu-id="8c85f-114">Para saber mais sobre a permissão **ReadWriteMailbox** e acessar os EWS na plataforma de suplementos do Office, confira [Noções básicas sobre permissões de suplemento do Outlook](understanding-outlook-add-in-permissions.md) e [Chamar serviços Web de um suplemento do Outlook](web-services.md).</span><span class="sxs-lookup"><span data-stu-id="8c85f-114">For more information about the **ReadWriteMailbox** permission, and accessing EWS through the Office Add-ins platform, see [Understanding Outlook add-in permissions](understanding-outlook-add-in-permissions.md) and [Call web services from an Outlook add-in](web-services.md).</span></span>

<span data-ttu-id="8c85f-115">**Tabela 1. Métodos assíncronos para obter ou definir propriedades de item em um formulário de redação**</span><span class="sxs-lookup"><span data-stu-id="8c85f-115">**Table 1. Asynchronous methods to get or set item properties in a compose form**</span></span>

<br/>

| <span data-ttu-id="8c85f-116">Propriedade</span><span class="sxs-lookup"><span data-stu-id="8c85f-116">Property</span></span> | <span data-ttu-id="8c85f-117">Tipo de propriedade</span><span class="sxs-lookup"><span data-stu-id="8c85f-117">Property type</span></span> | <span data-ttu-id="8c85f-118">Método assíncrono para obter</span><span class="sxs-lookup"><span data-stu-id="8c85f-118">Asynchronous method to get</span></span> | <span data-ttu-id="8c85f-119">Método(s) assíncrono(s) para definir</span><span class="sxs-lookup"><span data-stu-id="8c85f-119">Asynchronous method(s) to set</span></span> |
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="8c85f-120">bcc</span><span class="sxs-lookup"><span data-stu-id="8c85f-120">bcc</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[<span data-ttu-id="8c85f-121">Destinatários</span><span class="sxs-lookup"><span data-stu-id="8c85f-121">Recipients</span></span>](/javascript/api/outlook/office.Recipients)|[<span data-ttu-id="8c85f-122">Recipients.getAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-122">Recipients.getAsync</span></span>](/javascript/api/outlook/office.Recipients#getasync-options--callback-)|<span data-ttu-id="8c85f-123">[Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-), [Recipients.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)</span><span class="sxs-lookup"><span data-stu-id="8c85f-123">[Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-), [Recipients.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)</span></span>|
|[<span data-ttu-id="8c85f-124">body</span><span class="sxs-lookup"><span data-stu-id="8c85f-124">body</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[<span data-ttu-id="8c85f-125">Body</span><span class="sxs-lookup"><span data-stu-id="8c85f-125">Body</span></span>](/javascript/api/outlook/office.Body)|[<span data-ttu-id="8c85f-126">Body.getAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-126">Body.getAsync</span></span>](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-)|<span data-ttu-id="8c85f-127">[Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-), [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)</span><span class="sxs-lookup"><span data-stu-id="8c85f-127">[Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-), [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)</span></span>|
|[<span data-ttu-id="8c85f-128">cc</span><span class="sxs-lookup"><span data-stu-id="8c85f-128">cc</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|<span data-ttu-id="8c85f-129">Destinatários</span><span class="sxs-lookup"><span data-stu-id="8c85f-129">Recipients</span></span>|<span data-ttu-id="8c85f-130">Recipients.getAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-130">Recipients.getAsync</span></span>|<span data-ttu-id="8c85f-131">Recipients.addAsync Recipients.setAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-131">Recipients.addAsync Recipients.setAsync</span></span>|
|[<span data-ttu-id="8c85f-132">end</span><span class="sxs-lookup"><span data-stu-id="8c85f-132">end</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[<span data-ttu-id="8c85f-133">Time</span><span class="sxs-lookup"><span data-stu-id="8c85f-133">Time</span></span>](/javascript/api/outlook/office.Time)|[<span data-ttu-id="8c85f-134">Time.getAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-134">Time.getAsync</span></span>](/javascript/api/outlook/office.Time#getasync-options--callback-)|[<span data-ttu-id="8c85f-135">Time.setAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-135">Time.setAsync</span></span>](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)|
|[<span data-ttu-id="8c85f-136">location</span><span class="sxs-lookup"><span data-stu-id="8c85f-136">location</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[<span data-ttu-id="8c85f-137">Location</span><span class="sxs-lookup"><span data-stu-id="8c85f-137">Location</span></span>](/javascript/api/outlook/office.Location)|[<span data-ttu-id="8c85f-138">Location.getAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-138">Location.getAsync</span></span>](/javascript/api/outlook/office.Location#getasync-options--callback-)|[<span data-ttu-id="8c85f-139">Location.setAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-139">Location.setAsync</span></span>](/javascript/api/outlook/office.Location#setasync-location--options--callback-)|
|[<span data-ttu-id="8c85f-140">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="8c85f-140">optionalAttendees</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|<span data-ttu-id="8c85f-141">Destinatários</span><span class="sxs-lookup"><span data-stu-id="8c85f-141">Recipients</span></span>|<span data-ttu-id="8c85f-142">Recipients.getAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-142">Recipients.getAsync</span></span>|<span data-ttu-id="8c85f-143">Recipients.addAsync Recipients.setAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-143">Recipients.addAsync Recipients.setAsync</span></span>|
|[<span data-ttu-id="8c85f-144">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="8c85f-144">requiredAttendees</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|<span data-ttu-id="8c85f-145">Destinatários</span><span class="sxs-lookup"><span data-stu-id="8c85f-145">Recipients</span></span>|<span data-ttu-id="8c85f-146">Recipients.getAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-146">Recipients.getAsync</span></span>|<span data-ttu-id="8c85f-147">Recipients.addAsync Recipients.setAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-147">Recipients.addAsync Recipients.setAsync</span></span>|
|[<span data-ttu-id="8c85f-148">start</span><span class="sxs-lookup"><span data-stu-id="8c85f-148">start</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|<span data-ttu-id="8c85f-149">Hora</span><span class="sxs-lookup"><span data-stu-id="8c85f-149">Time</span></span>|<span data-ttu-id="8c85f-150">Time.getAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-150">Time.getAsync</span></span>|<span data-ttu-id="8c85f-151">Time.setAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-151">Time.setAsync</span></span>|
|[<span data-ttu-id="8c85f-152">subject</span><span class="sxs-lookup"><span data-stu-id="8c85f-152">subject</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[<span data-ttu-id="8c85f-153">Subject</span><span class="sxs-lookup"><span data-stu-id="8c85f-153">Subject</span></span>](/javascript/api/outlook/office.Subject)|[<span data-ttu-id="8c85f-154">Subject.getAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-154">Subject.getAsync</span></span>](/javascript/api/outlook/office.Subject#getasync-options--callback-)|[<span data-ttu-id="8c85f-155">Subject.setAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-155">Subject.setAsync</span></span>](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)|
|[<span data-ttu-id="8c85f-156">to</span><span class="sxs-lookup"><span data-stu-id="8c85f-156">to</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|<span data-ttu-id="8c85f-157">Destinatários</span><span class="sxs-lookup"><span data-stu-id="8c85f-157">Recipients</span></span>|<span data-ttu-id="8c85f-158">Recipients.getAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-158">Recipients.getAsync</span></span>|<span data-ttu-id="8c85f-159">Recipients.addAsync Recipients.setAsync</span><span class="sxs-lookup"><span data-stu-id="8c85f-159">Recipients.addAsync Recipients.setAsync</span></span>|

## <a name="see-also"></a><span data-ttu-id="8c85f-160">Confira também</span><span class="sxs-lookup"><span data-stu-id="8c85f-160">See also</span></span>

- [<span data-ttu-id="8c85f-161">Criar suplementos do Outlook para formulários de redação</span><span class="sxs-lookup"><span data-stu-id="8c85f-161">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="8c85f-162">Noções básicas sobre permissões de suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c85f-162">Understanding Outlook add-in permissions</span></span>](understanding-outlook-add-in-permissions.md)
- [<span data-ttu-id="8c85f-163">Chamar serviços Web de um suplemento do Outlook</span><span class="sxs-lookup"><span data-stu-id="8c85f-163">Call web services from an Outlook add-in</span></span>](web-services.md)
- [<span data-ttu-id="8c85f-164">Obter e definir dados de item do Outlook em formulários de leitura ou redação</span><span class="sxs-lookup"><span data-stu-id="8c85f-164">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)