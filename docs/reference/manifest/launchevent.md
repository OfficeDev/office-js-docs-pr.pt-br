---
title: LaunchEvent no arquivo de manifesto
description: O elemento LaunchEvent configura seu complemento para ser ativado com base em eventos suportados.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: c866a085ed6b7a33c8d7bf02d25e6ec748629e07
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591076"
---
# <a name="launchevent-element"></a><span data-ttu-id="2a433-103">Elemento LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="2a433-103">LaunchEvent element</span></span>

<span data-ttu-id="2a433-104">Configura seu complemento para ser ativado com base em eventos com suporte.</span><span class="sxs-lookup"><span data-stu-id="2a433-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="2a433-105">Filho do [`<LaunchEvents>`](launchevents.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="2a433-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="2a433-106">Para obter mais informações, [consulte Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="2a433-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="2a433-107">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="2a433-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2a433-108">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="2a433-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="2a433-109">Contido em</span><span class="sxs-lookup"><span data-stu-id="2a433-109">Contained in</span></span>

- [<span data-ttu-id="2a433-110">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="2a433-110">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="2a433-111">Atributos</span><span class="sxs-lookup"><span data-stu-id="2a433-111">Attributes</span></span>

|  <span data-ttu-id="2a433-112">Atributo</span><span class="sxs-lookup"><span data-stu-id="2a433-112">Attribute</span></span>  |  <span data-ttu-id="2a433-113">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="2a433-113">Required</span></span>  |  <span data-ttu-id="2a433-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="2a433-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2a433-115">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="2a433-115">**Type**</span></span>  |  <span data-ttu-id="2a433-116">Sim</span><span class="sxs-lookup"><span data-stu-id="2a433-116">Yes</span></span>  | <span data-ttu-id="2a433-117">Especifica um tipo de evento com suporte.</span><span class="sxs-lookup"><span data-stu-id="2a433-117">Specifies a supported event type.</span></span> <span data-ttu-id="2a433-118">Para o conjunto de tipos com suporte, consulte [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md#supported-events).</span><span class="sxs-lookup"><span data-stu-id="2a433-118">For the set of supported types, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md#supported-events).</span></span> |
|  <span data-ttu-id="2a433-119">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="2a433-119">**FunctionName**</span></span>  |  <span data-ttu-id="2a433-120">Sim</span><span class="sxs-lookup"><span data-stu-id="2a433-120">Yes</span></span>  | <span data-ttu-id="2a433-121">Especifica o nome da função JavaScript para manipular o evento especificado no `Type` atributo.</span><span class="sxs-lookup"><span data-stu-id="2a433-121">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="2a433-122">Confira também</span><span class="sxs-lookup"><span data-stu-id="2a433-122">See also</span></span>

- [<span data-ttu-id="2a433-123">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="2a433-123">LaunchEvents</span></span>](launchevents.md)
