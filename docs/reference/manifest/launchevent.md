---
title: LaunchEvent no arquivo de manifesto
description: O elemento LaunchEvent configura seu complemento para ser ativado com base em eventos suportados.
ms.date: 03/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 71469693bff7213455582a3247778cabf92c2aa3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745810"
---
# <a name="launchevent-element"></a>Elemento LaunchEvent

Configura seu complemento para ser ativado com base em eventos com suporte. Filho do [`<LaunchEvents>`](launchevents.md) elemento. Para obter mais informações, [consulte Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).

**Tipo de suplemento:** Email

**Válido somente nesses esquemas VersionOverrides**:

- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

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
|  **SendMode** (visualização) |  Não  | Usado por `OnMessageSend` e `OnAppointmentSend` eventos. Especifica as opções disponíveis para o usuário se o seu complemento impedir que um item seja enviado ou se o add-in estiver indisponível. Se a **propriedade SendMode** não estiver incluída, a `SoftBlock` opção será definida por padrão. Para opções disponíveis, consulte [Opções de SendMode disponíveis](#available-sendmode-options-preview). |

## <a name="available-sendmode-options-preview"></a>Opções de SendMode disponíveis (visualização)

Quando você incluir o `OnMessageSend` ou `OnAppointmentSend` evento no manifesto, você também deve definir a **propriedade SendMode** . Se a **propriedade SendMode** não estiver incluída, a `SoftBlock` opção será definida por padrão. A seguir estão as opções disponíveis. Com base nas condições que seu complemento está procurando, o usuário será alertado se o seu complemento encontrar um problema no item que está sendo enviado.

| Opção SendMode | Descrição |
|---|---|
|`PromptUser`|Se o item não atender às condições do complemento, o usuário poderá escolher **Enviar** De qualquer maneira no alerta ou resolver o problema e tentar enviar o item novamente. Se o complemento estiver demorando muito para processar o item, o usuário será solicitado a parar de executar o add-in e escolher Enviar de qualquer **maneira**. Se o complemento não estiver disponível (por exemplo, há um erro ao carregar o complemento), o item será enviado.|
|`SoftBlock`|Opção padrão se a **propriedade SendMode** não estiver incluída. O usuário é alertado de que o item que está enviando não está de acordo com as condições do complemento e deve resolver o problema antes de tentar enviar o item novamente. No entanto, se o complemento não estiver disponível (por exemplo, há um erro ao carregar o complemento), o item será enviado.|
|`Block`|O item não será enviado se ocorrer qualquer uma das seguintes situações.<br>- O item não está de acordo com as condições do complemento.<br>- O complemento não pode se conectar ao servidor.<br>- Há um erro ao carregar o complemento.|

## <a name="see-also"></a>Confira também

- [LaunchEvents](launchevents.md)
- [Configurar seu Outlook para ativação baseada em eventos](../../outlook/autolaunch.md#supported-events)
- [Use Alertas Inteligentes e o evento OnMessageSend em seu Outlook de usuário](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
