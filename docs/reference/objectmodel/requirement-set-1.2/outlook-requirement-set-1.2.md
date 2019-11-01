---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.2
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 898e768dfc1828ba44f29e9da5c4baa61de186cb
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902092"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="65362-102">Conjunto de requisitos de API para suplementos do Outlook versão 1.2</span><span class="sxs-lookup"><span data-stu-id="65362-102">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="65362-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="65362-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="65362-104">Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="65362-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-12"></a><span data-ttu-id="65362-105">Novidades na versão 1.2?</span><span class="sxs-lookup"><span data-stu-id="65362-105">What's new in 1.2?</span></span>

<span data-ttu-id="65362-p101">O conjunto de requisitos 1.2 inclui todos os recursos do [Conjunto de requisitos 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Ele adicionou a capacidade de os suplementos inserirem texto no cursor do usuário, no assunto ou no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="65362-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="65362-108">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="65362-108">Change log</span></span>

- <span data-ttu-id="65362-109">Foi adicionado o [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): Retorna de forma assíncrona os dados selecionados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="65362-109">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="65362-110">Foi adicionado o [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback): Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="65362-110">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="65362-111">Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Foi adicionada a propriedade `attachments` ao parâmetro `formData`.</span><span class="sxs-lookup"><span data-stu-id="65362-111">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="65362-112">Foi modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): Foi adicionada a propriedade `attachments` ao parâmetro `formData`.</span><span class="sxs-lookup"><span data-stu-id="65362-112">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="65362-113">Confira também</span><span class="sxs-lookup"><span data-stu-id="65362-113">See also</span></span>

- [<span data-ttu-id="65362-114">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="65362-114">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="65362-115">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="65362-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="65362-116">Introdução</span><span class="sxs-lookup"><span data-stu-id="65362-116">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="65362-117">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="65362-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
