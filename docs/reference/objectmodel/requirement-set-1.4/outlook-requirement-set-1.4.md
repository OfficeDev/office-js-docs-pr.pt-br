---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.4
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 7f8297fed8b94f3d949e260c38572284621e3840
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42597037"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="76eb2-102">Conjunto de requisitos de API para suplementos do Outlook versão 1.4</span><span class="sxs-lookup"><span data-stu-id="76eb2-102">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="76eb2-103">O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="76eb2-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="76eb2-104">Esta documentação destina-se a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="76eb2-104">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="76eb2-105">Novidades na versão 1.4</span><span class="sxs-lookup"><span data-stu-id="76eb2-105">What's new in 1.4?</span></span>

<span data-ttu-id="76eb2-p101">O conjunto de requisitos 1.4 inclui todos os recursos do [Conjunto de requisitos 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Adicionou acesso ao namespace `Office.ui`.</span><span class="sxs-lookup"><span data-stu-id="76eb2-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="76eb2-108">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="76eb2-108">Change log</span></span>

- <span data-ttu-id="76eb2-109">Foi adicionado o [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Exibe uma caixa de diálogo em um host do Office.</span><span class="sxs-lookup"><span data-stu-id="76eb2-109">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="76eb2-110">Foi adicionado o [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): fornece uma mensagem da caixa de diálogo à sua página pai/de abertura.</span><span class="sxs-lookup"><span data-stu-id="76eb2-110">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="76eb2-111">Foi adicionado o objeto [Dialog](/javascript/api/office/office.dialog): o objeto retornado quando o método [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) é chamado.</span><span class="sxs-lookup"><span data-stu-id="76eb2-111">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="76eb2-112">Confira também</span><span class="sxs-lookup"><span data-stu-id="76eb2-112">See also</span></span>

- [<span data-ttu-id="76eb2-113">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="76eb2-113">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="76eb2-114">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="76eb2-114">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="76eb2-115">Introdução</span><span class="sxs-lookup"><span data-stu-id="76eb2-115">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="76eb2-116">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="76eb2-116">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
