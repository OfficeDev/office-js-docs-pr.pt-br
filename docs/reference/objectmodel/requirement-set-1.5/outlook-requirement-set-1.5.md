---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.5
description: ''
ms.date: 01/16/2019
localization_priority: Normal
ms.openlocfilehash: fde394ff4b75e0f6b160f5d56cb73adc9da9dede
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388372"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="1eddd-102">Conjunto de requisitos de API para suplementos do Outlook versão 1.5</span><span class="sxs-lookup"><span data-stu-id="1eddd-102">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="1eddd-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="1eddd-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="1eddd-104">Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="1eddd-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="1eddd-105">Novidades na versão 1.5?</span><span class="sxs-lookup"><span data-stu-id="1eddd-105">What's new in 1.5?</span></span>

<span data-ttu-id="1eddd-p101">O conjunto de requisitos 1.5 inclui todos os recursos do [Conjunto de requisitos 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) e contém os seguintes recursos adicionais.</span><span class="sxs-lookup"><span data-stu-id="1eddd-p101">Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.</span></span>

- <span data-ttu-id="1eddd-108">Adicionado suporte para [painéis de tarefas fixáveis](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span><span class="sxs-lookup"><span data-stu-id="1eddd-108">Added support for [pinnable task panes](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).</span></span>
- <span data-ttu-id="1eddd-109">Adicionado suporte para chamar [APIs REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="1eddd-109">Added support for calling [REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>
- <span data-ttu-id="1eddd-110">Adicionada a capacidade de marcar um anexo como embutido.</span><span class="sxs-lookup"><span data-stu-id="1eddd-110">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="1eddd-111">Adicionada a capacidade de fechar um painel de tarefas ou uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="1eddd-111">Added ability to close a task pane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="1eddd-112">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="1eddd-112">Change log</span></span>

- <span data-ttu-id="1eddd-113">Adicionado o [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback): adiciona um manipulador de eventos para um evento compatível.</span><span class="sxs-lookup"><span data-stu-id="1eddd-113">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="1eddd-114">Adicionado [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#removehandlerasynceventtype-options-callback): remove os manipuladores de eventos para um tipo de evento aceitos.</span><span class="sxs-lookup"><span data-stu-id="1eddd-114">Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#removehandlerasynceventtype-options-callback): Removes the event handlers for a supported event type.</span></span>
- <span data-ttu-id="1eddd-115">Adicionado o [Office.EventType](office.md#eventtype-string): especifica o evento associado a um manipulador de eventos e inclui suporte para o evento ItemChanged.</span><span class="sxs-lookup"><span data-stu-id="1eddd-115">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="1eddd-116">Adicionado o [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string): obtém a URL do ponto de extremidade REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="1eddd-116">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="1eddd-p102">Modificado o [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback): Uma nova versão deste método com uma nova assinatura (`getCallbackTokenAsync([options], callback)`) foi adicionada. A versão original ainda está disponível e não é alterada.</span><span class="sxs-lookup"><span data-stu-id="1eddd-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="1eddd-119">Adicionado [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span><span class="sxs-lookup"><span data-stu-id="1eddd-119">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span></span>
- <span data-ttu-id="1eddd-120">Modificado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): um novo valor no dicionário `options` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="1eddd-120">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="1eddd-121">Modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata): Um novo valor no dicionário do `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="1eddd-121">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="1eddd-122">Modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata): um novo valor no dicionário `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="1eddd-122">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="1eddd-123">Confira também</span><span class="sxs-lookup"><span data-stu-id="1eddd-123">See also</span></span>

- [<span data-ttu-id="1eddd-124">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="1eddd-124">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="1eddd-125">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="1eddd-125">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="1eddd-126">Introdução</span><span class="sxs-lookup"><span data-stu-id="1eddd-126">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)
