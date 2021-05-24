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
# <a name="launchevents-element"></a>Elemento LaunchEvents

Configura seu complemento para ser ativado com base em eventos com suporte. Filho do [`<ExtensionPoint>`](extensionpoint.md) elemento. Para obter mais informações, [consulte Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).

**Tipo de suplemento:** Email

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

[ExtensionPoint](extensionpoint.md) ( Complemento de email **LaunchEvent)**

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | Sim |  Mapeie o evento com suporte para sua função no arquivo JavaScript para ativação do complemento. |

## <a name="see-also"></a>Confira também

- [LaunchEvent](launchevent.md)
