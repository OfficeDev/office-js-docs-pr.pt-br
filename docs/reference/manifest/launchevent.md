---
title: LaunchEvent no arquivo de manifesto (versão prévia)
description: O elemento LaunchEvent configura seu suplemento para ser ativado com base nos eventos com suporte.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 4874b9f4c14e3a999f41ec3fa20a15393b031ea6
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611775"
---
# <a name="launchevent-element-preview"></a><span data-ttu-id="f175d-103">Elemento LaunchEvent (visualização)</span><span class="sxs-lookup"><span data-stu-id="f175d-103">LaunchEvent element (preview)</span></span>

<span data-ttu-id="f175d-104">Configura o suplemento para que ele seja ativado com base nos eventos com suporte.</span><span class="sxs-lookup"><span data-stu-id="f175d-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="f175d-105">Filho do [`<LaunchEvents>`](launchevents.md) elemento.</span><span class="sxs-lookup"><span data-stu-id="f175d-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="f175d-106">Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).</span><span class="sxs-lookup"><span data-stu-id="f175d-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="f175d-107">**Tipo de suplemento:** Email</span><span class="sxs-lookup"><span data-stu-id="f175d-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f175d-108">A ativação baseada em evento está atualmente [em versão prévia](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas no Outlook na Web.</span><span class="sxs-lookup"><span data-stu-id="f175d-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="f175d-109">Para obter mais informações, consulte [como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span><span class="sxs-lookup"><span data-stu-id="f175d-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="f175d-110">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="f175d-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="f175d-111">Contido em</span><span class="sxs-lookup"><span data-stu-id="f175d-111">Contained in</span></span>

- [<span data-ttu-id="f175d-112">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="f175d-112">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="f175d-113">Atributos</span><span class="sxs-lookup"><span data-stu-id="f175d-113">Attributes</span></span>

|  <span data-ttu-id="f175d-114">Atributo</span><span class="sxs-lookup"><span data-stu-id="f175d-114">Attribute</span></span>  |  <span data-ttu-id="f175d-115">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f175d-115">Required</span></span>  |  <span data-ttu-id="f175d-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="f175d-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f175d-117">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="f175d-117">**Type**</span></span>  |  <span data-ttu-id="f175d-118">Sim</span><span class="sxs-lookup"><span data-stu-id="f175d-118">Yes</span></span>  | <span data-ttu-id="f175d-119">Especifica um tipo de evento suportado.</span><span class="sxs-lookup"><span data-stu-id="f175d-119">Specifies a supported event type.</span></span> <span data-ttu-id="f175d-120">Os tipos disponíveis são `OnNewMessageCompose` e `OnNewAppointmentOrganizer` .</span><span class="sxs-lookup"><span data-stu-id="f175d-120">Available types are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> |
|  <span data-ttu-id="f175d-121">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="f175d-121">**FunctionName**</span></span>  |  <span data-ttu-id="f175d-122">Sim</span><span class="sxs-lookup"><span data-stu-id="f175d-122">Yes</span></span>  | <span data-ttu-id="f175d-123">Especifica o nome da função JavaScript para manipular o evento especificado no `Type` atributo.</span><span class="sxs-lookup"><span data-stu-id="f175d-123">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="f175d-124">Confira também</span><span class="sxs-lookup"><span data-stu-id="f175d-124">See also</span></span>

- [<span data-ttu-id="f175d-125">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="f175d-125">LaunchEvents</span></span>](launchevents.md)
