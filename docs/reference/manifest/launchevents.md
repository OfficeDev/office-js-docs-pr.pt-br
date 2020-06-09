---
title: LaunchEvents no arquivo de manifesto (versão prévia)
description: O elemento LaunchEvents configura seu suplemento para ser ativado com base nos eventos com suporte.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 92416f8c646326410a8cd9ee7831e17a5c5f1ffc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611768"
---
# <a name="launchevents-element-preview"></a>Elemento LaunchEvents (visualização)

Configura o suplemento para que ele seja ativado com base nos eventos com suporte. Filho do [`<ExtensionPoint>`](extensionpoint.md) elemento. Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).

**Tipo de suplemento:** Email

> [!IMPORTANT]
> A ativação baseada em evento está atualmente [em versão prévia](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas no Outlook na Web. Para obter mais informações, consulte [como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

## <a name="syntax"></a>Sintaxe

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

## <a name="contained-in"></a>Contido em

[ExtensionPoint](extensionpoint.md) (suplemento de email do**LaunchEvent** )

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | Sim |  Mapeie o evento suportado para sua função no arquivo JavaScript para ativação de suplemento. |

## <a name="see-also"></a>Confira também

- [LaunchEvent](launchevent.md)
