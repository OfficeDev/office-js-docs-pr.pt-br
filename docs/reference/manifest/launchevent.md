---
title: LaunchEvent no arquivo de manifesto (versão prévia)
description: O elemento LaunchEvent configura seu suplemento para ser ativado com base nos eventos com suporte.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: a4f5208ec7f735d926c3a878cae34973c3992cf9
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278522"
---
# <a name="launchevent-element-preview"></a><span data-ttu-id="e0101-103">Elemento LaunchEvent (visualização)</span><span class="sxs-lookup"><span data-stu-id="e0101-103">LaunchEvent element (preview)</span></span>

<span data-ttu-id="e0101-104">Configura o suplemento para que ele seja ativado com base nos eventos com suporte.</span><span class="sxs-lookup"><span data-stu-id="e0101-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="e0101-105">Filho do [`<LaunchEvents>`](launchevents.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="e0101-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="e0101-106">Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="e0101-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="e0101-107">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="e0101-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e0101-108">A ativação baseada em evento está atualmente [em versão prévia](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="e0101-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="e0101-109">Para obter mais informações, consulte [como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="e0101-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="e0101-110">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="e0101-110">Syntax</span></span>

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## <a name="contained-in"></a><span data-ttu-id="e0101-111">Contido em</span><span class="sxs-lookup"><span data-stu-id="e0101-111">Contained in</span></span>

- [<span data-ttu-id="e0101-112">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="e0101-112">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="e0101-113">Atributos</span><span class="sxs-lookup"><span data-stu-id="e0101-113">Attributes</span></span>

|  <span data-ttu-id="e0101-114">Atributo</span><span class="sxs-lookup"><span data-stu-id="e0101-114">Attribute</span></span>  |  <span data-ttu-id="e0101-115">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="e0101-115">Required</span></span>  |  <span data-ttu-id="e0101-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="e0101-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e0101-117">**Type**</span><span class="sxs-lookup"><span data-stu-id="e0101-117">**Type**</span></span>  |  <span data-ttu-id="e0101-118">Sim</span><span class="sxs-lookup"><span data-stu-id="e0101-118">Yes</span></span>  | <span data-ttu-id="e0101-119">Especifica um tipo de evento suportado.</span><span class="sxs-lookup"><span data-stu-id="e0101-119">Specifies a supported event type.</span></span> <span data-ttu-id="e0101-120">Os tipos disponíveis são `OnNewMessageCompose` e `OnNewAppointmentOrganizer` .</span><span class="sxs-lookup"><span data-stu-id="e0101-120">Available types are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> |
|  <span data-ttu-id="e0101-121">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="e0101-121">**FunctionName**</span></span>  |  <span data-ttu-id="e0101-122">Sim</span><span class="sxs-lookup"><span data-stu-id="e0101-122">Yes</span></span>  | <span data-ttu-id="e0101-123">Especifica o nome da função JavaScript para manipular o evento especificado no `Type` atributo.</span><span class="sxs-lookup"><span data-stu-id="e0101-123">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="e0101-124">Confira também</span><span class="sxs-lookup"><span data-stu-id="e0101-124">See also</span></span>

- [<span data-ttu-id="e0101-125">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="e0101-125">LaunchEvents</span></span>](launchevents.md)
