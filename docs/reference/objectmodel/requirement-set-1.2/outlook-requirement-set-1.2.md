---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.2
description: Recursos e APIs que foram introduzidos para os Outlook e as APIs JavaScript Office como parte da API de Caixa de Correio 1.2.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: d643f0fdf07c5f22d8d863075b894cfc05b21363
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590397"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="996b2-103">Conjunto de requisitos de API para suplementos do Outlook versão 1.2</span><span class="sxs-lookup"><span data-stu-id="996b2-103">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="996b2-104">O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.</span><span class="sxs-lookup"><span data-stu-id="996b2-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="996b2-105">Esta documentação destina-se a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="996b2-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-12"></a><span data-ttu-id="996b2-106">Novidades na versão 1.2?</span><span class="sxs-lookup"><span data-stu-id="996b2-106">What's new in 1.2?</span></span>

<span data-ttu-id="996b2-107">O conjunto de requisitos 1.2 inclui todos os recursos do conjunto de requisitos [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md).</span><span class="sxs-lookup"><span data-stu-id="996b2-107">Requirement set 1.2 includes all of the features of [requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md).</span></span> <span data-ttu-id="996b2-108">Ele adicionou a capacidade de os suplementos inserirem texto no cursor do usuário, no assunto ou no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="996b2-108">It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="996b2-109">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="996b2-109">Change log</span></span>

- <span data-ttu-id="996b2-110">Foi adicionado o [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Retorna de forma assíncrona os dados selecionados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="996b2-110">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="996b2-111">Foi adicionado o [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="996b2-111">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="996b2-112">Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Foi adicionada a propriedade `attachments` ao parâmetro `formData`.</span><span class="sxs-lookup"><span data-stu-id="996b2-112">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="996b2-113">Foi modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Foi adicionada a propriedade `attachments` ao parâmetro `formData`.</span><span class="sxs-lookup"><span data-stu-id="996b2-113">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="996b2-114">Confira também</span><span class="sxs-lookup"><span data-stu-id="996b2-114">See also</span></span>

- [<span data-ttu-id="996b2-115">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="996b2-115">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="996b2-116">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="996b2-116">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="996b2-117">Introdução</span><span class="sxs-lookup"><span data-stu-id="996b2-117">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="996b2-118">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="996b2-118">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
