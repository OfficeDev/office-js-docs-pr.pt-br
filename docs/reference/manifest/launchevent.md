---
title: LaunchEvent no arquivo manifesto (visualização)
description: O elemento LaunchEvent configura seu complemento para ativar com base em eventos suportados.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 7283e9aba9ca57793019ffe027a7f4d6e3243aa8
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555308"
---
# <a name="launchevent-element-preview"></a>Elemento LaunchEvent (pré-visualização)

Configura seu complemento para ativar com base em eventos suportados. Filho do [`<LaunchEvents>`](launchevents.md) elemento. Para obter mais informações, consulte [Configurar seu Outlook complemento para ativação baseada em eventos](../../outlook/autolaunch.md).

**Tipo de suplemento:** Email

> [!IMPORTANT]
> A ativação baseada em eventos está atualmente [em pré-visualização](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas em Outlook na web e em Windows. Para obter mais informações, consulte [Como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

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
|  **Tipo**  |  Sim  | Especifica um tipo de evento suportado. Para obter o conjunto de tipos suportados, consulte [Como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#supported-events). |
|  **FunctionName**  |  Sim  | Especifica o nome da função JavaScript para lidar com o evento especificado no `Type` atributo. |

## <a name="see-also"></a>Confira também

- [LaunchEvents](launchevents.md)
