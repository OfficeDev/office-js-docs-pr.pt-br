---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.4
description: Recursos e APIs que foram introduzidos para os Outlook e as APIs JavaScript Office como parte da API de Caixa de Correio 1.4.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 19d77784926ac09d5620eb36242701da59b39f09
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591013"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="fb00b-103">Conjunto de requisitos de API para suplementos do Outlook versão 1.4</span><span class="sxs-lookup"><span data-stu-id="fb00b-103">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="fb00b-104">O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.</span><span class="sxs-lookup"><span data-stu-id="fb00b-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="fb00b-105">Esta documentação destina-se a um [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md) que não seja o conjunto de requisitos mais recente.</span><span class="sxs-lookup"><span data-stu-id="fb00b-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="fb00b-106">Novidades na versão 1.4</span><span class="sxs-lookup"><span data-stu-id="fb00b-106">What's new in 1.4?</span></span>

<span data-ttu-id="fb00b-107">O conjunto de requisitos 1.4 inclui todos os recursos do conjunto [de requisitos 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md).</span><span class="sxs-lookup"><span data-stu-id="fb00b-107">Requirement set 1.4 includes all of the features of [requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md).</span></span> <span data-ttu-id="fb00b-108">Adicionou acesso ao namespace `Office.ui`.</span><span class="sxs-lookup"><span data-stu-id="fb00b-108">It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="fb00b-109">Log de alterações</span><span class="sxs-lookup"><span data-stu-id="fb00b-109">Change log</span></span>

- <span data-ttu-id="fb00b-110">Adicionado [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): exibe uma caixa de diálogo em um Office aplicativo.</span><span class="sxs-lookup"><span data-stu-id="fb00b-110">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office application.</span></span>
- <span data-ttu-id="fb00b-111">Foi adicionado o [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): fornece uma mensagem da caixa de diálogo à sua página pai/de abertura.</span><span class="sxs-lookup"><span data-stu-id="fb00b-111">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="fb00b-112">Foi adicionado o objeto [Dialog](/javascript/api/office/office.dialog): o objeto retornado quando o método [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) é chamado.</span><span class="sxs-lookup"><span data-stu-id="fb00b-112">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="fb00b-113">Confira também</span><span class="sxs-lookup"><span data-stu-id="fb00b-113">See also</span></span>

- [<span data-ttu-id="fb00b-114">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="fb00b-114">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="fb00b-115">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="fb00b-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="fb00b-116">Introdução</span><span class="sxs-lookup"><span data-stu-id="fb00b-116">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="fb00b-117">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="fb00b-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
