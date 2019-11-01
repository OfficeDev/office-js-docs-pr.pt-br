---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.4
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: ded3aa8cdcde1132f074f55356e64bb4468919ff
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901938"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="24d82-102">Conjunto de requisitos de API para suplementos do Outlook versão 1.4</span><span class="sxs-lookup"><span data-stu-id="24d82-102">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="24d82-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="24d82-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="24d82-104">Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="24d82-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="24d82-105">Novidades na versão 1.4</span><span class="sxs-lookup"><span data-stu-id="24d82-105">What's new in 1.4?</span></span>

<span data-ttu-id="24d82-p101">O conjunto de requisitos 1.4 inclui todos os recursos do [Conjunto de requisitos 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Adicionou acesso ao namespace `Office.ui`.</span><span class="sxs-lookup"><span data-stu-id="24d82-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="24d82-108">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="24d82-108">Change log</span></span>

- <span data-ttu-id="24d82-109">Foi adicionado o [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Exibe uma caixa de diálogo em um host do Office.</span><span class="sxs-lookup"><span data-stu-id="24d82-109">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="24d82-110">Foi adicionado o [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): fornece uma mensagem da caixa de diálogo à sua página pai/de abertura.</span><span class="sxs-lookup"><span data-stu-id="24d82-110">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="24d82-111">Foi adicionado o objeto [Dialog](/javascript/api/office/office.dialog): o objeto retornado quando o método [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) é chamado.</span><span class="sxs-lookup"><span data-stu-id="24d82-111">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="24d82-112">Confira também</span><span class="sxs-lookup"><span data-stu-id="24d82-112">See also</span></span>

- [<span data-ttu-id="24d82-113">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="24d82-113">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="24d82-114">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="24d82-114">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="24d82-115">Introdução</span><span class="sxs-lookup"><span data-stu-id="24d82-115">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="24d82-116">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="24d82-116">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
