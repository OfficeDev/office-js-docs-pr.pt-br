---
title: LaunchEvent no arquivo de manifesto
description: O elemento LaunchEvent configura seu complemento para ser ativado com base em eventos suportados.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="launchevent-element"></a>Elemento LaunchEvent

Configura seu complemento para ser ativado com base em eventos com suporte. Filho do [`<LaunchEvents>`](launchevents.md) elemento. Para obter mais informações, consulte [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).

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
|  **SendMode** (visualização) |  Não  | Obrigatório para e `OnMessageSend` `OnAppointmentSend` eventos. Especifica as opções disponíveis para o usuário se o seu complemento impedir que o item seja enviado. Para opções disponíveis, consulte [Opções de SendMode disponíveis](#available-sendmode-options-preview). |

## <a name="available-sendmode-options-preview"></a>Opções de SendMode disponíveis (visualização)

Ao incluir o evento `OnMessageSend` ou `OnAppointmentSend` no manifesto, você também deve definir a **propriedade SendMode** . A seguir estão as opções disponíveis. Com base nas condições que seu complemento está procurando, o usuário será alertado se o seu complemento encontrar um problema no item que está sendo enviado.

| Opção SendMode | Descrição |
|---|---|
|`PromptUser`|No alerta, o usuário pode optar por **Enviar** de qualquer maneira ou resolver o problema e tentar enviar o item novamente.|
|`SoftBlock`|O usuário deve corrigir o problema antes de tentar enviar o item novamente.|

## <a name="see-also"></a>Confira também

- [LaunchEvents](launchevents.md)
- [Configurar seu Outlook para ativação baseada em eventos](../../outlook/autolaunch.md#supported-events)
- [Use Alertas Inteligentes e o evento OnMessageSend em seu Outlook de usuário](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
