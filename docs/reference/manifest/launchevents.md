---
title: LaunchEvents no arquivo de manifesto
description: O elemento LaunchEvents configura seu complemento para ser ativado com base em eventos suportados.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 16d721ca6d9402d2bd5d19787707e146358044f0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590909"
---
# <a name="launchevents-element"></a><span data-ttu-id="8d907-103">Elemento LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="8d907-103">LaunchEvents element</span></span>

<span data-ttu-id="8d907-104">Configura seu complemento para ser ativado com base em eventos com suporte.</span><span class="sxs-lookup"><span data-stu-id="8d907-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="8d907-105">Filho do [`<ExtensionPoint>`](extensionpoint.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="8d907-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="8d907-106">Para obter mais informações, [consulte Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="8d907-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="8d907-107">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="8d907-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8d907-108">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="8d907-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="8d907-109">Contido em</span><span class="sxs-lookup"><span data-stu-id="8d907-109">Contained in</span></span>

<span data-ttu-id="8d907-110">[ExtensionPoint](extensionpoint.md) ( Complemento de email **LaunchEvent)**</span><span class="sxs-lookup"><span data-stu-id="8d907-110">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="8d907-111">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="8d907-111">Child elements</span></span>

|  <span data-ttu-id="8d907-112">Elemento</span><span class="sxs-lookup"><span data-stu-id="8d907-112">Element</span></span> |  <span data-ttu-id="8d907-113">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="8d907-113">Required</span></span>  |  <span data-ttu-id="8d907-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="8d907-114">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="8d907-115">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="8d907-115">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="8d907-116">Sim</span><span class="sxs-lookup"><span data-stu-id="8d907-116">Yes</span></span> |  <span data-ttu-id="8d907-117">Mapeie o evento com suporte para sua função no arquivo JavaScript para ativação do complemento.</span><span class="sxs-lookup"><span data-stu-id="8d907-117">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="8d907-118">Confira também</span><span class="sxs-lookup"><span data-stu-id="8d907-118">See also</span></span>

- [<span data-ttu-id="8d907-119">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="8d907-119">LaunchEvent</span></span>](launchevent.md)
