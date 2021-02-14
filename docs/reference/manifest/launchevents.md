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
# <a name="launchevents-element-preview"></a>Elemento LaunchEvents (visualização)

Configura o seu complemento para ativar com base em eventos com suporte. Filho do [`<ExtensionPoint>`](extensionpoint.md) elemento. Para saber mais, confira [Configurar seu complemento do Outlook para ativação baseada em eventos.](../../outlook/autolaunch.md)

**Tipo de suplemento:** Email

> [!IMPORTANT]
> A ativação baseada em eventos está [atualmente em visualização](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e só está disponível no Outlook na Web e no Windows. Para obter mais informações, [consulte Como visualizar o recurso de ativação baseada em eventos.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)

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

[ExtensionPoint](extensionpoint.md) ( add-in de email **LaunchEvent)**

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | Sim |  Mapeie o evento com suporte para sua função no arquivo JavaScript para ativação do complemento. |

## <a name="see-also"></a>Confira também

- [LaunchEvent](launchevent.md)
