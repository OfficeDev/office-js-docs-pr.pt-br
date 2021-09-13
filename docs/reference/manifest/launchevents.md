---
title: LaunchEvents no arquivo de manifesto
description: O elemento LaunchEvents configura seu complemento para ser ativado com base em eventos suportados.
ms.date: 05/11/2021
ms.localizationpriority: medium
ms.openlocfilehash: 02e0b21d65733492a783ffb099caf9e76225e53f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148932"
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
