---
title: LaunchEvent no arquivo de manifesto (versão prévia)
description: O elemento LaunchEvent configura seu suplemento para ser ativado com base nos eventos com suporte.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: a4f5208ec7f735d926c3a878cae34973c3992cf9
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278522"
---
# <a name="launchevent-element-preview"></a>Elemento LaunchEvent (visualização)

Configura o suplemento para que ele seja ativado com base nos eventos com suporte. Filho do [`<LaunchEvents>`](launchevents.md) elemento. Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).

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

- [LaunchEvents](launchevents.md)

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Type**  |  Sim  | Especifica um tipo de evento suportado. Os tipos disponíveis são `OnNewMessageCompose` e `OnNewAppointmentOrganizer` . |
|  **FunctionName**  |  Sim  | Especifica o nome da função JavaScript para manipular o evento especificado no `Type` atributo. |

## <a name="see-also"></a>Confira também

- [LaunchEvents](launchevents.md)
