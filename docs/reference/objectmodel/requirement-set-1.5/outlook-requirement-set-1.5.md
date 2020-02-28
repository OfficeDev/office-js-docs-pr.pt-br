---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.5
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: e9dda19a4dd9831ca1bb7924558ac99c26baf763
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324789"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="d81d1-102">Conjunto de requisitos de API para suplementos do Outlook versão 1.5</span><span class="sxs-lookup"><span data-stu-id="d81d1-102">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="d81d1-103">O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="d81d1-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="d81d1-104">Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="d81d1-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="d81d1-105">Novidades na versão 1.5?</span><span class="sxs-lookup"><span data-stu-id="d81d1-105">What's new in 1.5?</span></span>

<span data-ttu-id="d81d1-p101">O conjunto de requisitos 1.5 inclui todos os recursos do [Conjunto de requisitos 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) e contém os seguintes recursos adicionais.</span><span class="sxs-lookup"><span data-stu-id="d81d1-p101">Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.</span></span>

- <span data-ttu-id="d81d1-108">Adicionado suporte para [painéis de tarefas fixáveis](../../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="d81d1-108">Added support for [pinnable task panes](../../../outlook/pinnable-taskpane.md).</span></span>
- <span data-ttu-id="d81d1-109">Adicionado suporte para chamar [APIs REST](../../../outlook/use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="d81d1-109">Added support for calling [REST APIs](../../../outlook/use-rest-api.md).</span></span>
- <span data-ttu-id="d81d1-110">Adicionada a capacidade de marcar um anexo como embutido.</span><span class="sxs-lookup"><span data-stu-id="d81d1-110">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="d81d1-111">Adicionada a capacidade de fechar um painel de tarefas ou uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="d81d1-111">Added ability to close a task pane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="d81d1-112">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="d81d1-112">Change log</span></span>

- <span data-ttu-id="d81d1-113">Adicionado o [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): adiciona um manipulador de eventos para um evento compatível.</span><span class="sxs-lookup"><span data-stu-id="d81d1-113">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="d81d1-114">Foi adicionado o [Office. Context. Mailbox. removeHandlerAsync](office.context.mailbox.md#methods): remove os manipuladores de eventos para um tipo de evento suportado.</span><span class="sxs-lookup"><span data-stu-id="d81d1-114">Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): Removes the event handlers for a supported event type.</span></span>
- <span data-ttu-id="d81d1-115">Adicionado o [Office.EventType](office.md#eventtype-string): especifica o evento associado a um manipulador de eventos e inclui suporte para o evento ItemChanged.</span><span class="sxs-lookup"><span data-stu-id="d81d1-115">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="d81d1-116">Adicionado o [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): obtém a URL do ponto de extremidade REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="d81d1-116">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="d81d1-p102">Modificado o [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): Uma nova versão deste método com uma nova assinatura (`getCallbackTokenAsync([options], callback)`) foi adicionada. A versão original ainda está disponível e não é alterada.</span><span class="sxs-lookup"><span data-stu-id="d81d1-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="d81d1-119">Adicionado [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span><span class="sxs-lookup"><span data-stu-id="d81d1-119">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span></span>
- <span data-ttu-id="d81d1-120">Modificado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): um novo valor no dicionário `options` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="d81d1-120">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="d81d1-121">Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Um novo valor no dicionário do `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="d81d1-121">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="d81d1-122">Modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): um novo valor no dicionário `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="d81d1-122">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="d81d1-123">Confira também</span><span class="sxs-lookup"><span data-stu-id="d81d1-123">See also</span></span>

- [<span data-ttu-id="d81d1-124">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="d81d1-124">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="d81d1-125">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="d81d1-125">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="d81d1-126">Introdução</span><span class="sxs-lookup"><span data-stu-id="d81d1-126">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="d81d1-127">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="d81d1-127">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
