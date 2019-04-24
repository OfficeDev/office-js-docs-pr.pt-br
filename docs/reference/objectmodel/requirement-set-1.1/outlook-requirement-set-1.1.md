---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: cd284a5871139b7f6bf006a9deb3671a937682f6
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450300"
---
# <a name="outlook-add-in-api-requirement-set-11"></a><span data-ttu-id="c50b8-102">Conjunto de requisitos de API para suplementos do Outlook versão 1.1</span><span class="sxs-lookup"><span data-stu-id="c50b8-102">Outlook add-in API requirement set 1.1</span></span>

<span data-ttu-id="c50b8-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="c50b8-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="c50b8-104">Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o mais recente.</span><span class="sxs-lookup"><span data-stu-id="c50b8-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-11"></a><span data-ttu-id="c50b8-105">Novidades na versão 1.1?</span><span class="sxs-lookup"><span data-stu-id="c50b8-105">What's new in 1.1?</span></span>

<span data-ttu-id="c50b8-p101">O conjunto de requisitos 1.1 inclui todos os recursos do Conjunto de requisitos 1.0. Ele adicionou a capacidade de os suplementos para acessarem o corpo de mensagens e os compromissos e a capacidade de modificar o item atual.</span><span class="sxs-lookup"><span data-stu-id="c50b8-p101">Requirement set 1.1 includes all of the features of Requirement set 1.0. It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.</span></span>

### <a name="change-log"></a><span data-ttu-id="c50b8-108">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="c50b8-108">Change log</span></span>

- <span data-ttu-id="c50b8-109">Foi adicionado o objeto [Body](/javascript/api/outlook_1_1/office.body): Fornece métodos para adicionar e atualizar o conteúdo de um item em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c50b8-109">Added [Body](/javascript/api/outlook_1_1/office.body) object: Provides methods for adding and updating the content of an item in an Outlook add-in.</span></span>
- <span data-ttu-id="c50b8-110">Foi adicionado o objeto [Location](/javascript/api/outlook_1_1/office.location): Fornece métodos para obter e definir o local de uma reunião em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c50b8-110">Added [Location](/javascript/api/outlook_1_1/office.location) object: Provides methods to get and set the location of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="c50b8-111">Foi adicionado o objeto [Recipients](/javascript/api/outlook_1_1/office.recipients): Fornece métodos para obter e definir os destinatários de um compromisso ou uma mensagem em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c50b8-111">Added [Recipients](/javascript/api/outlook_1_1/office.recipients) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="c50b8-112">Foi adicionado o objeto [Subject](/javascript/api/outlook_1_1/office.subject): Fornece métodos para obter e definir o assunto de um compromisso ou uma mensagem em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c50b8-112">Added [Subject](/javascript/api/outlook_1_1/office.subject) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="c50b8-113">Foi adicionado o objeto [Time](/javascript/api/outlook_1_1/office.time): Fornece métodos para obter e definir o tempo de início ou fim de uma reunião em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c50b8-113">Added [Time](/javascript/api/outlook_1_1/office.time) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="c50b8-114">Foi adicionado o [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): Adiciona um arquivo a uma mensagem ou um compromisso como um anexo.</span><span class="sxs-lookup"><span data-stu-id="c50b8-114">Added [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): Adds a file to a message or appointment as an attachment.</span></span>
- <span data-ttu-id="c50b8-115">Foi adicionado o [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback): Adiciona um item do Exchange, como uma mensagem, como anexo na mensagem ou no compromisso.</span><span class="sxs-lookup"><span data-stu-id="c50b8-115">Added [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback): Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>
- <span data-ttu-id="c50b8-116">Foi adicionado o [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback): Remove um anexo de uma mensagem ou de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="c50b8-116">Added [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback): Removes an attachment from a message or appointment.</span></span>
- <span data-ttu-id="c50b8-117">Foi adicionado o [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body): Obtém um objeto que fornece métodos para manipular o corpo de um item.</span><span class="sxs-lookup"><span data-stu-id="c50b8-117">Added [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body): Gets an object that provides methods for manipulating the body of an item.</span></span>
- <span data-ttu-id="c50b8-118">Foi adicionada a linha [Office. Context. Mailbox. Item. Bcc](office.context.mailbox.item.md#bcc-recipients) de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="c50b8-118">Added [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#bcc-recipients) line of a message.</span></span>
- <span data-ttu-id="c50b8-119">Adicionado o [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype): especifica o tipo de destinatário para um compromisso.</span><span class="sxs-lookup"><span data-stu-id="c50b8-119">Added [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype): Specifies the type of recipient for an appointment.</span></span>

## <a name="see-also"></a><span data-ttu-id="c50b8-120">Confira também</span><span class="sxs-lookup"><span data-stu-id="c50b8-120">See also</span></span>

- [<span data-ttu-id="c50b8-121">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="c50b8-121">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="c50b8-122">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="c50b8-122">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="c50b8-123">Introdução</span><span class="sxs-lookup"><span data-stu-id="c50b8-123">Get started</span></span>](/outlook/add-ins/quick-start)
