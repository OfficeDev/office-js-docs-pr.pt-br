---
title: LaunchEvents no arquivo de manifesto (versão prévia)
description: O elemento LaunchEvents configura seu suplemento para ser ativado com base nos eventos com suporte.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 2e1ad56d405fca0f85fad500a113fba7d0448caf
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278521"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="c4b6b-103">Elemento LaunchEvents (visualização)</span><span class="sxs-lookup"><span data-stu-id="c4b6b-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="c4b6b-104">Configura o suplemento para que ele seja ativado com base nos eventos com suporte.</span><span class="sxs-lookup"><span data-stu-id="c4b6b-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="c4b6b-105">Filho do [`<ExtensionPoint>`](extensionpoint.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="c4b6b-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="c4b6b-106">Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="c4b6b-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="c4b6b-107">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="c4b6b-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c4b6b-108">A ativação baseada em evento está atualmente [em versão prévia](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="c4b6b-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="c4b6b-109">Para obter mais informações, consulte [como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="c4b6b-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="c4b6b-110">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="c4b6b-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="c4b6b-111">Contido em</span><span class="sxs-lookup"><span data-stu-id="c4b6b-111">Contained in</span></span>

<span data-ttu-id="c4b6b-112">[ExtensionPoint](extensionpoint.md) (suplemento de email do**LaunchEvent** )</span><span class="sxs-lookup"><span data-stu-id="c4b6b-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="c4b6b-113">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="c4b6b-113">Child elements</span></span>

|  <span data-ttu-id="c4b6b-114">Elemento</span><span class="sxs-lookup"><span data-stu-id="c4b6b-114">Element</span></span> |  <span data-ttu-id="c4b6b-115">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="c4b6b-115">Required</span></span>  |  <span data-ttu-id="c4b6b-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="c4b6b-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="c4b6b-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="c4b6b-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="c4b6b-118">Sim</span><span class="sxs-lookup"><span data-stu-id="c4b6b-118">Yes</span></span> |  <span data-ttu-id="c4b6b-119">Mapeie o evento suportado para sua função no arquivo JavaScript para ativação de suplemento.</span><span class="sxs-lookup"><span data-stu-id="c4b6b-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="c4b6b-120">Confira também</span><span class="sxs-lookup"><span data-stu-id="c4b6b-120">See also</span></span>

- [<span data-ttu-id="c4b6b-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="c4b6b-121">LaunchEvent</span></span>](launchevent.md)
