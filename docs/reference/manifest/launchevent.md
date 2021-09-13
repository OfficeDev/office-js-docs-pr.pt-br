---
title: LaunchEvent no arquivo de manifesto
description: O elemento LaunchEvent configura seu complemento para ser ativado com base em eventos suportados.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 23615424e194917a15b20ea4afbf7d9c5b8017e9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151756"
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
