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
|  **Tipo**  |  Sim  | Especifica um tipo de evento suportado. Os tipos disponíveis são `OnNewMessageCompose` e `OnNewAppointmentOrganizer` . |
|  **FunctionName**  |  Sim  | Especifica o nome da função JavaScript para manipular o evento especificado no `Type` atributo. |

## <a name="see-also"></a>Confira também

- [LaunchEvents](launchevents.md)
