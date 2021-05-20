---
title: LaunchEvent no arquivo manifesto (visualização)
description: O elemento LaunchEvent configura seu complemento para ativar com base em eventos suportados.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 7283e9aba9ca57793019ffe027a7f4d6e3243aa8
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555308"
---
# <a name="launchevent-element-preview"></a><span data-ttu-id="e6222-103">Elemento LaunchEvent (pré-visualização)</span><span class="sxs-lookup"><span data-stu-id="e6222-103">LaunchEvent element (preview)</span></span>

<span data-ttu-id="e6222-104">Configura seu complemento para ativar com base em eventos suportados.</span><span class="sxs-lookup"><span data-stu-id="e6222-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="e6222-105">Filho do [`<LaunchEvents>`](launchevents.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="e6222-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="e6222-106">Para obter mais informações, consulte [Configurar seu Outlook complemento para ativação baseada em eventos](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="e6222-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="e6222-107">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="e6222-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e6222-108">A ativação baseada em eventos está atualmente [em pré-visualização](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas em Outlook na web e em Windows.</span><span class="sxs-lookup"><span data-stu-id="e6222-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and on Windows.</span></span> <span data-ttu-id="e6222-109">Para obter mais informações, consulte [Como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="e6222-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="e6222-110">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="e6222-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="e6222-111">Contido em</span><span class="sxs-lookup"><span data-stu-id="e6222-111">Contained in</span></span>

- [<span data-ttu-id="e6222-112">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="e6222-112">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="e6222-113">Atributos</span><span class="sxs-lookup"><span data-stu-id="e6222-113">Attributes</span></span>

|  <span data-ttu-id="e6222-114">Atributo</span><span class="sxs-lookup"><span data-stu-id="e6222-114">Attribute</span></span>  |  <span data-ttu-id="e6222-115">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="e6222-115">Required</span></span>  |  <span data-ttu-id="e6222-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="e6222-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e6222-117">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="e6222-117">**Type**</span></span>  |  <span data-ttu-id="e6222-118">Sim</span><span class="sxs-lookup"><span data-stu-id="e6222-118">Yes</span></span>  | <span data-ttu-id="e6222-119">Especifica um tipo de evento suportado.</span><span class="sxs-lookup"><span data-stu-id="e6222-119">Specifies a supported event type.</span></span> <span data-ttu-id="e6222-120">Para obter o conjunto de tipos suportados, consulte [Como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#supported-events).</span><span class="sxs-lookup"><span data-stu-id="e6222-120">For the set of supported types, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#supported-events).</span></span> |
|  <span data-ttu-id="e6222-121">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="e6222-121">**FunctionName**</span></span>  |  <span data-ttu-id="e6222-122">Sim</span><span class="sxs-lookup"><span data-stu-id="e6222-122">Yes</span></span>  | <span data-ttu-id="e6222-123">Especifica o nome da função JavaScript para lidar com o evento especificado no `Type` atributo.</span><span class="sxs-lookup"><span data-stu-id="e6222-123">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="e6222-124">Confira também</span><span class="sxs-lookup"><span data-stu-id="e6222-124">See also</span></span>

- [<span data-ttu-id="e6222-125">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="e6222-125">LaunchEvents</span></span>](launchevents.md)
