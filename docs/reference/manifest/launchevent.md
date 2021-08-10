---
title: LaunchEvent no arquivo de manifesto
description: O elemento LaunchEvent configura seu complemento para ser ativado com base em eventos suportados.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 763ee6b0776cce0118df35dbbad3488a54fd7599cc774de240d4f1130ed021cb
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089761"
---
# <a name="launchevent-element"></a>Elemento LaunchEvent

Configura seu complemento para ser ativado com base em eventos com suporte. Filho do [`<LaunchEvents>`](launchevents.md) elemento. Para obter mais informações, [consulte Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).

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

- [LaunchEvents](launchevents.md)

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Tipo**  |  Sim  | Especifica um tipo de evento com suporte. Para o conjunto de tipos com suporte, consulte [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md#supported-events). |
|  **FunctionName**  |  Sim  | Especifica o nome da função JavaScript para manipular o evento especificado no `Type` atributo. |

## <a name="see-also"></a>Confira também

- [LaunchEvents](launchevents.md)
