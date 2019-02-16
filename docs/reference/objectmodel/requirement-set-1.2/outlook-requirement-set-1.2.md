---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 1767b1b93f13de2c8a0731d2f08a1141b709b734
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068025"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="f5a2b-102">Conjunto de requisitos de API para suplementos do Outlook versão 1.2</span><span class="sxs-lookup"><span data-stu-id="f5a2b-102">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="f5a2b-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="f5a2b-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="f5a2b-104">Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="f5a2b-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-12"></a><span data-ttu-id="f5a2b-105">Novidades na versão 1.2?</span><span class="sxs-lookup"><span data-stu-id="f5a2b-105">What's new in 1.2?</span></span>

<span data-ttu-id="f5a2b-p101">O conjunto de requisitos 1.2 inclui todos os recursos do [Conjunto de requisitos 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Ele adicionou a capacidade de os suplementos inserirem texto no cursor do usuário, no assunto ou no corpo da mensagem.</span><span class="sxs-lookup"><span data-stu-id="f5a2b-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="f5a2b-108">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="f5a2b-108">Change log</span></span>

- <span data-ttu-id="f5a2b-109">Foi adicionado o [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): Retorna de forma assíncrona os dados selecionados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f5a2b-109">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="f5a2b-110">Foi adicionado o [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback): Insere de forma assíncrona os dados no corpo ou no assunto de uma mensagem.</span><span class="sxs-lookup"><span data-stu-id="f5a2b-110">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="f5a2b-111">Foi modificado o [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Foi adicionada a propriedade `attachments` ao parâmetro `formData`.</span><span class="sxs-lookup"><span data-stu-id="f5a2b-111">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="f5a2b-112">Foi modificado o [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): Foi adicionada a propriedade `attachments` ao parâmetro `formData`.</span><span class="sxs-lookup"><span data-stu-id="f5a2b-112">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="f5a2b-113">Confira também</span><span class="sxs-lookup"><span data-stu-id="f5a2b-113">See also</span></span>

- [<span data-ttu-id="f5a2b-114">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="f5a2b-114">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="f5a2b-115">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="f5a2b-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="f5a2b-116">Introdução</span><span class="sxs-lookup"><span data-stu-id="f5a2b-116">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)
