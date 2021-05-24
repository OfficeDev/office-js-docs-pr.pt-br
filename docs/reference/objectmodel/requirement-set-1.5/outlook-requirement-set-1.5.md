---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.5
description: Recursos e APIs que foram introduzidos para Outlook e as APIs JavaScript Office como parte da API de Caixa de Correio 1.5.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7d780538a77f54db6f1234a6d29a3bcdea9533b0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590838"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="db398-103">Conjunto de requisitos de API para suplementos do Outlook versão 1.5</span><span class="sxs-lookup"><span data-stu-id="db398-103">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="db398-104">O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.</span><span class="sxs-lookup"><span data-stu-id="db398-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="db398-105">Esta documentação destina-se a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="db398-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="db398-106">Novidades na versão 1.5?</span><span class="sxs-lookup"><span data-stu-id="db398-106">What's new in 1.5?</span></span>

<span data-ttu-id="db398-107">O conjunto de requisitos 1.5 inclui todos os recursos do conjunto [de requisitos 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md).</span><span class="sxs-lookup"><span data-stu-id="db398-107">Requirement set 1.5 includes all of the features of [requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md).</span></span> <span data-ttu-id="db398-108">Ele adicionou os seguintes recursos.</span><span class="sxs-lookup"><span data-stu-id="db398-108">It added the following features.</span></span>

- <span data-ttu-id="db398-109">Adicionado suporte para [painéis de tarefas fixáveis](../../../outlook/pinnable-taskpane.md).</span><span class="sxs-lookup"><span data-stu-id="db398-109">Added support for [pinnable task panes](../../../outlook/pinnable-taskpane.md).</span></span>
- <span data-ttu-id="db398-110">Adicionado suporte para chamar [APIs REST](../../../outlook/use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="db398-110">Added support for calling [REST APIs](../../../outlook/use-rest-api.md).</span></span>
- <span data-ttu-id="db398-111">Adicionada a capacidade de marcar um anexo como embutido.</span><span class="sxs-lookup"><span data-stu-id="db398-111">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="db398-112">Adicionada a capacidade de fechar um painel de tarefas ou uma caixa de diálogo.</span><span class="sxs-lookup"><span data-stu-id="db398-112">Added ability to close a task pane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="db398-113">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="db398-113">Change log</span></span>

- <span data-ttu-id="db398-114">Adicionado o [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): adiciona um manipulador de eventos para um evento compatível.</span><span class="sxs-lookup"><span data-stu-id="db398-114">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="db398-115">Adicionado [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): remove os manipuladores de eventos para um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="db398-115">Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): Removes the event handlers for a supported event type.</span></span>
- <span data-ttu-id="db398-116">Adicionado o [Office.EventType](office.md#eventtype-string): especifica o evento associado a um manipulador de eventos e inclui suporte para o evento ItemChanged.</span><span class="sxs-lookup"><span data-stu-id="db398-116">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="db398-117">Adicionado o [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): obtém a URL do ponto de extremidade REST para esta conta de email.</span><span class="sxs-lookup"><span data-stu-id="db398-117">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="db398-p102">Modificado o [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): Uma nova versão deste método com uma nova assinatura (`getCallbackTokenAsync([options], callback)`) foi adicionada. A versão original ainda está disponível e não é alterada.</span><span class="sxs-lookup"><span data-stu-id="db398-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="db398-120">Adicionado [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span><span class="sxs-lookup"><span data-stu-id="db398-120">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span></span>
- <span data-ttu-id="db398-121">Modificado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): um novo valor no dicionário `options` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="db398-121">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="db398-122">Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Um novo valor no dicionário do `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="db398-122">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="db398-123">Modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): um novo valor no dicionário `formData.attachments` chamado `isInline`, usado para especificar que uma imagem foi usada embutida no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="db398-123">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="db398-124">Confira também</span><span class="sxs-lookup"><span data-stu-id="db398-124">See also</span></span>

- [<span data-ttu-id="db398-125">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="db398-125">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="db398-126">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="db398-126">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="db398-127">Introdução</span><span class="sxs-lookup"><span data-stu-id="db398-127">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="db398-128">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="db398-128">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
