---
title: LaunchEvents no arquivo de manifesto (visualização)
description: O elemento LaunchEvents configura seu complemento para ser ativado com base em eventos com suporte.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 59c52aa3f60e69e2bdda84718c6123f02942fedc
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237977"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="ff655-103">Elemento LaunchEvents (visualização)</span><span class="sxs-lookup"><span data-stu-id="ff655-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="ff655-104">Configura o seu complemento para ativar com base em eventos com suporte.</span><span class="sxs-lookup"><span data-stu-id="ff655-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="ff655-105">Filho do [`<ExtensionPoint>`](extensionpoint.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="ff655-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="ff655-106">Para saber mais, confira [Configurar seu complemento do Outlook para ativação baseada em eventos.](../../outlook/autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="ff655-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="ff655-107">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="ff655-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ff655-108">A ativação baseada em eventos está [atualmente em visualização](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e só está disponível no Outlook na Web e no Windows.</span><span class="sxs-lookup"><span data-stu-id="ff655-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and on Windows.</span></span> <span data-ttu-id="ff655-109">Para obter mais informações, [consulte Como visualizar o recurso de ativação baseada em eventos.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)</span><span class="sxs-lookup"><span data-stu-id="ff655-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="ff655-110">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="ff655-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="ff655-111">Contido em</span><span class="sxs-lookup"><span data-stu-id="ff655-111">Contained in</span></span>

<span data-ttu-id="ff655-112">[ExtensionPoint](extensionpoint.md) ( add-in de email **LaunchEvent)**</span><span class="sxs-lookup"><span data-stu-id="ff655-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="ff655-113">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ff655-113">Child elements</span></span>

|  <span data-ttu-id="ff655-114">Elemento</span><span class="sxs-lookup"><span data-stu-id="ff655-114">Element</span></span> |  <span data-ttu-id="ff655-115">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ff655-115">Required</span></span>  |  <span data-ttu-id="ff655-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="ff655-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="ff655-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="ff655-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="ff655-118">Sim</span><span class="sxs-lookup"><span data-stu-id="ff655-118">Yes</span></span> |  <span data-ttu-id="ff655-119">Mapeie o evento com suporte para sua função no arquivo JavaScript para ativação do complemento.</span><span class="sxs-lookup"><span data-stu-id="ff655-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="ff655-120">Confira também</span><span class="sxs-lookup"><span data-stu-id="ff655-120">See also</span></span>

- [<span data-ttu-id="ff655-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="ff655-121">LaunchEvent</span></span>](launchevent.md)
