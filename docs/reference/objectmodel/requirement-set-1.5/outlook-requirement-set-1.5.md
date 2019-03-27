---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d0489e4efa763b3963fcdc78ec894db46fa06362
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870860"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="9bc50-102">Conjunto de requisitos de API para suplementos do Outlook versão 1.5</span><span class="sxs-lookup"><span data-stu-id="9bc50-102">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="9bc50-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="9bc50-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="9bc50-104">Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="9bc50-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="9bc50-105">Novidades na versão 1.5?</span><span class="sxs-lookup"><span data-stu-id="9bc50-105">What's new in 1.5?</span></span>

<span data-ttu-id="9bc50-p101">O conjunto de requisitos 1.5 inclui todos os recursos do [Conjunto de requisitos 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) e contém os seguintes recursos adicionais.</span><span class="sxs-lookup"><span data-stu-id="9bc50-p101">Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.</span></span>

- <span data-ttu-id="9bc50-108">Adicionado suporte para [painéis de tarefas fixáveis](/outlook/add-ins/pinnable-taskpane).</span><span class="sxs-lookup"><span data-stu-id="9bc50-108">Added support for [pinnable task panes](/outlook/add-ins/pinnable-taskpane).</span></span>
- <span data-ttu-id="9bc50-109">Adicionado suporte para chamar [APIs REST](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="9bc50-109">Added support for calling [REST APIs](/outlook/add-ins/use-rest-api).</span></span>
- <span data-ttu-id="9bc50-110">Adicionada a capacidade de marcar um anexo como embutido.</span><span class="sxs-lookup"><span data-stu-id="9bc50-110">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="9bc50-111">Adicionada a capacidade de fechar um painel de tarefas ou uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="9bc50-111">Added ability to close a task pane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="9bc50-112">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="9bc50-112">Change log</span></span>

- <span data-ttu-id="9bc50-113">Adicionado o [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback): adiciona um manipulador de eventos para um evento compatível.</span><span class="sxs-lookup"><span data-stu-id="9bc50-113">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="9bc50-114">Foi adicionado o [Office. Context. Mailbox. removeHandlerAsync](office.context.mailbox.md#removehandlerasynceventtype-options-callback): remove os manipuladores de eventos para um tipo de evento suportado.</span><span class="sxs-lookup"><span data-stu-id="9bc50-114">Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#removehandlerasynceventtype-options-callback): Removes the event handlers for a supported event type.</span></span>
- <span data-ttu-id="9bc50-115">Adicionado o [Office.EventType](office.md#eventtype-string): especifica o evento associado a um manipulador de eventos e inclui suporte para o evento ItemChanged.</span><span class="sxs-lookup"><span data-stu-id="9bc50-115">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="9bc50-116">Adicionado o [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string): obtém a URL do ponto de extremidade REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="9bc50-116">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="9bc50-p102">Modificado o [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback): Uma nova versão deste método com uma nova assinatura (`getCallbackTokenAsync([options], callback)`) foi adicionada. A versão original ainda está disponível e não é alterada.</span><span class="sxs-lookup"><span data-stu-id="9bc50-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="9bc50-119">Adicionado [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span><span class="sxs-lookup"><span data-stu-id="9bc50-119">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span></span>
- <span data-ttu-id="9bc50-120">Modificado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): um novo valor no dicionário `options` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="9bc50-120">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="9bc50-121">Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Um novo valor no dicionário do `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="9bc50-121">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="9bc50-122">Modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): um novo valor no dicionário `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="9bc50-122">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="9bc50-123">Confira também</span><span class="sxs-lookup"><span data-stu-id="9bc50-123">See also</span></span>

- [<span data-ttu-id="9bc50-124">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="9bc50-124">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="9bc50-125">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="9bc50-125">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="9bc50-126">Introdução</span><span class="sxs-lookup"><span data-stu-id="9bc50-126">Get started</span></span>](/outlook/add-ins/quick-start)
