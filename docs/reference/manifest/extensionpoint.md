---
title: Elemento ExtensionPoint no arquivo de manifesto
description: Define onde um suplemento expõe a funcionalidade na interface de usuário do Office.
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1f8ccc08a9c0d42edf89c904b8809a530239be4c
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855629"
---
# <a name="extensionpoint-element"></a>Elemento ExtensionPoint

 Define onde um suplemento expõe a funcionalidade na interface de usuário do Office. O elemento **ExtensionPoint** é um elemento filho de [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Sim  | O tipo de ponto de extensão que está sendo definido. Os valores possíveis dependem do aplicativo host Office definido no valor do elemento **Host do avô**.|

## <a name="extension-points-for-excel-onenote-powerpoint-and-word-add-in-commands"></a>Pontos de extensão para Excel, OneNote, PowerPoint e comandos de complemento do Word

Há três tipos de pontos de extensão disponíveis em alguns ou todos esses hosts.

- [PrimaryCommandSurface](#primarycommandsurface) (Válido para Word, Excel, PowerPoint e OneNote) - A faixa de opções Office.
- [ContextMenu](#contextmenu) (Válido para Word, Excel, PowerPoint e OneNote) - O menu de atalho que aparece quando você seleciona e segura (ou clica com o botão direito do mouse) na interface do usuário Office.
- [CustomFunctions](#customfunctions) (Válido somente para Excel) - Uma função personalizada escrita em JavaScript para Excel.

Consulte as subseções a seguir para os elementos filho e exemplos desses tipos de pontos de extensão.

### <a name="primarycommandsurface"></a>PrimaryCommandSurface

A superfície de comando principal no Word, Excel, PowerPoint e OneNote é a faixa de opções.

#### <a name="child-elements"></a>Elementos filho

|Elemento|Descrição|
|:-----|:-----|
|[CustomTab] (customtab.md|Obrigatório se você quiser adicionar uma guia personalizada à faixa de opções (usando **PrimaryCommandSurface**). Se você usar o elemento **CustomTab**, o elemento **OfficeTab** não poderá ser usado. O atributo **id** é obrigatório. |
|[OfficeTab](officetab.md)|Obrigatório se você quiser estender uma guia padrão Aplicativo do Office faixa de opções (usando **PrimaryCommandSurface**). Se você usar o elemento **OfficeTab**, o elemento **CustomTab** não poderá ser usado.|

#### <a name="example"></a>Exemplo

O exemplo a seguir mostra como usar o elemento **ExtensionPoint** com **PrimaryCommandSurface**. Ele adiciona uma guia personalizada à faixa de opções.

> [!IMPORTANT]
> Forneça uma ID exclusiva para os elementos que contêm um atributo ID.

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="Contoso.MyTab1">
    <Label resid="residLabel4" />
    <Group id="Contoso.Group1">
      <Label resid="residLabel4" />
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Tooltip resid="residToolTip" />
      <Control xsi:type="Button" id="Contoso.Button1">
          <!-- information about the control -->
      </Control>
      <!-- other controls, as needed -->
    </Group>
  </CustomTab>
</ExtensionPoint>
```

### <a name="contextmenu"></a>ContextMenu

Um menu de contexto é um menu de atalho que aparece quando você clica com o botão direito do mouse na interface Office interface do usuário.

#### <a name="child-elements"></a>Elementos filho
 
|Elemento|Descrição|
|:-----|:-----|
|[OfficeMenu](officemenu.md)|Obrigatório se você estiver adicionando comandos de suplemento a um menu de contexto padrão (usando **ContextMenu**). O **atributo id** deve ser definido como uma das seguintes cadeias de caracteres: <br/> - **ContextMenuText** se o menu de contexto deve ser aberto quando um usuário clica com o botão direito do mouse no texto selecionado. <br/> - **ContextMenuCell** se o menu de contexto deve abrir quando o usuário clicar com o botão direito do mouse em uma célula em uma Excel planilha.|

#### <a name="example"></a>Exemplo

A seguir, adiciona um menu de contexto personalizado às células em uma Excel de dados.

```xml
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="Contoso.ContextMenu2">
            <!-- information about the control -->
    </Control>
    <!-- other controls, as needed -->
  </OfficeMenu>
</ExtensionPoint>
```

### <a name="customfunctions"></a>CustomFunctions

Uma função personalizada escrita em JavaScript ou TypeScript para Excel.

#### <a name="child-elements"></a>Elementos filho

|Elemento|Descrição|
|:-----|:-----|
|[Script](script.md)|Obrigatório. Links para o arquivo JavaScript com o código de registro e definição da função personalizada.|
|[Page](page.md)|Obrigatório. Links para a página HTML de suas funções personalizadas.|
|[Metadados](metadata.md)|Obrigatório. Define as configurações de metadados usados por uma função personalizada no Excel.|
|[Namespace](namespace.md)|Opcional. Define o namespace usado por uma função personalizada no Excel.|

#### <a name="example"></a>Exemplo

```xml
<ExtensionPoint xsi:type="CustomFunctions">
  <Script>
    <SourceLocation resid="Functions.Script.Url"/>
  </Script>
  <Page>
    <SourceLocation resid="Shared.Url"/>
  </Page>
  <Metadata>
    <SourceLocation resid="Functions.Metadata.Url"/>
  </Metadata>
  <Namespace resid="Functions.Namespace"/>
</ExtensionPoint>
```

## <a name="extension-points-for-outlook"></a>Pontos de extensão para Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface)
- [MessageComposeCommandSurface](#messagecomposecommandsurface)
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (Só pode ser usado em [DesktopFormFactor](desktopformfactor.md)).
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [MobileOnlineMeetingCommandSurface](#mobileonlinemeetingcommandsurface)
- [LaunchEvent](#launchevent)
- [Eventos](#events)
- [DetectedEntity](#detectedentity)

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface

Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de email. No Outlook para área de trabalho, isso aparece na faixa de opções.

#### <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adiciona os comandos à guia da faixa de opções padrão.  |
|  [CustomTab](customtab.md) |  Adiciona os comandos à guia da faixa de opções personalizada.  |

#### <a name="officetab-example"></a>Exemplo de OfficeTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Exemplo de CustomTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="Contoso.TabCustom2">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a>MessageComposeCommandSurface

Este ponto de extensão coloca botões na faixa de opções para suplementos que usam o formulário de composição de email. 

#### <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adiciona os comandos à guia da faixa de opções padrão.  |
|  [CustomTab](customtab.md) |  Adiciona os comandos à guia da faixa de opções personalizada.  |

#### <a name="officetab-example"></a>Exemplo de OfficeTab

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Exemplo de CustomTab

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="Contoso.TabCustom3">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a>AppointmentOrganizerCommandSurface

Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao organizador da reunião. 

#### <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adiciona os comandos à guia da faixa de opções padrão.  |
|  [CustomTab](customtab.md) |  Adiciona os comandos à guia da faixa de opções personalizada.  |

#### <a name="officetab-example"></a>Exemplo de OfficeTab

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Exemplo de CustomTab

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="Contoso.TabCustom4">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a>AppointmentAttendeeCommandSurface

Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao participante da reunião. 

#### <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adiciona os comandos à guia da faixa de opções padrão.  |
|  [CustomTab](customtab.md) |  Adiciona os comandos à guia da faixa de opções personalizada.  |

#### <a name="officetab-example"></a>Exemplo de OfficeTab

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a>Exemplo de CustomTab

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="Contoso.TabCustom5">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Module

Este ponto de extensão coloca botões na faixa de opções para a extensão do módulo.

> [!IMPORTANT]
> Registrar eventos [de Caixa de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) Correio e [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) não está disponível com esse ponto de extensão.

#### <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
|  [OfficeTab](officetab.md) |  Adiciona os comandos à guia da faixa de opções padrão.  |
|  [CustomTab](customtab.md) |  Adiciona os comandos à guia da faixa de opções personalizada.  |

### <a name="mobilemessagereadcommandsurface"></a>MobileMessageReadCommandSurface

Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de email no fator forma móvel.

#### <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
|  [Group](group.md) |  Adiciona um grupo de botões à superfície de comando.  |

Os elementos **ExtensionPoint** desse tipo só podem ter um elemento filho: um elemento **Group**.

Os elementos **Control** contidos neste ponto de extensão precisam ter o atributo **xsi:type** definido como `MobileButton`.

#### <a name="example"></a>Exemplo

```xml
<ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
  <Group id="Contoso.mobileGroup1">
    <Label resid="residAppName"/>
      <Control  xsi:type="MobileButton id="Contoso.mobileButton1"">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="mobileonlinemeetingcommandsurface"></a>MobileOnlineMeetingCommandSurface

Esse ponto de extensão coloca uma alternância apropriada para o modo na superfície de comando para um compromisso no fator de forma móvel. Um organizador de reunião pode criar uma reunião online. Um participante pode participar posteriormente da reunião online. Para saber mais sobre esse cenário, consulte o artigo Criar um Outlook para um [provedor de reunião](../../outlook/online-meeting.md) online.

> [!NOTE]
> Esse ponto de extensão só é suportado no Android e no iOS com uma assinatura Microsoft 365 assinatura.
>
> Registrar eventos [de Caixa de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) Correio e [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) não está disponível com esse ponto de extensão.

#### <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
|  [Control](control.md) |  Adiciona um botão à superfície de comando.  |

`ExtensionPoint` elementos desse tipo só podem ter um elemento filho: um `Control` elemento.

O `Control` elemento contido neste ponto de extensão deve ter o `xsi:type` atributo definido como `MobileButton`.

As `Icon` imagens devem estar em escala de cinza usando código hexaxa `#919191` ou seu equivalente em [outros formatos de cor](https://convertingcolors.com/hex-color-919191.html).

#### <a name="example"></a>Exemplo

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="Contoso.onlineMeetingFunctionButton1">
    <Label resid="residUILessButton0Name" />
    <Icon>
      <bt:Image resid="UiLessIcon" size="25" scale="1" />
      <bt:Image resid="UiLessIcon" size="25" scale="2" />
      <bt:Image resid="UiLessIcon" size="25" scale="3" />
      <bt:Image resid="UiLessIcon" size="32" scale="1" />
      <bt:Image resid="UiLessIcon" size="32" scale="2" />
      <bt:Image resid="UiLessIcon" size="32" scale="3" />
      <bt:Image resid="UiLessIcon" size="48" scale="1" />
      <bt:Image resid="UiLessIcon" size="48" scale="2" />
      <bt:Image resid="UiLessIcon" size="48" scale="3" />
    </Icon>
    <Action xsi:type="ExecuteFunction">
      <FunctionName>insertContosoMeeting</FunctionName>
    </Action>
  </Control>
</ExtensionPoint>
```

### <a name="launchevent"></a>LaunchEvent

Esse ponto de extensão permite que um complemento seja ativado com base em eventos suportados no fator de formulário da área de trabalho. Para saber mais sobre esse cenário e para obter a lista completa de eventos com suporte, consulte o artigo Configurar seu Outlook [de](../../outlook/autolaunch.md) ativação baseada em eventos.

> [!IMPORTANT]
> Registrar eventos [de Caixa de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) Correio e [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) não está disponível com esse ponto de extensão.

#### <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
| [LaunchEvents](launchevents.md) |  Lista de [LaunchEvent](launchevent.md) para ativação baseada em evento.  |
| [SourceLocation](sourcelocation.md) |  O local do arquivo JavaScript de origem.  |

#### <a name="example"></a>Exemplo

```xml
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

### <a name="events"></a>Eventos

Este ponto de extensão adiciona um manipulador de eventos para um evento especificado. Para obter mais informações sobre como usar esse ponto de extensão, consulte Recurso Ao enviar [para Outlook de complementos](../../outlook/outlook-on-send-addins.md).

> [!IMPORTANT]
> Registrar eventos [de Caixa de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) Correio e [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) não está disponível com esse ponto de extensão.

| Elemento | Descrição  |
|:-----|:-----|
|  [Event](event.md) |  Especifica o evento e a função de manipulador de eventos.  |

#### <a name="itemsend-event-example"></a>Exemplo do evento ItemSend

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a>DetectedEntity

Este ponto extensão adiciona uma ativação do suplemento contextual em um tipo de entidade especificada.

> [!IMPORTANT]
> Registrar eventos [de Caixa de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) Correio e [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) não está disponível com esse ponto de extensão.

O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.

> [!NOTE]
> Este tipo de elemento está disponível para [ clientes do Outlook que ofereçam suporte a conjuntos de requisitos 1.6 e posteriores](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).

|  Elemento |  Descrição  |
|:-----|:-----|
|  [Label](#label) |  Especifica o rótulo para o suplemento na janela contextual.  |
|  [SourceLocation](sourcelocation.md) |  Especifica a URL para a janela contextual.  |
|  [Rule](rule.md) |  Especifica a regra ou regras que determinam quando um suplemento é ativado.  |

#### <a name="label"></a>Label

Obrigatório. O rótulo do grupo. O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no [elemento Resources](resources.md) .

#### <a name="highlight-requirements"></a>Requisitos de realce

A única maneira que um usuário pode ativar um suplemento contextual é interagir com uma entidade realçada. Os desenvolvedores podem controlar quais entidades são realçadas usando o atributo `Highlight` do elemento `Rule` para os tipos de regra `ItemHasKnownEntity` e `ItemHasRegularExpressionMatch`.

No entanto, há algumas limitações que devem ser consideradas. Essas limitações são para garantir que sempre haverá uma entidade realçada em compromissos ou mensagens aplicáveis para oferecer ao usuário uma maneira de ativar o suplemento.

- Os tipos de entidade `EmailAddress` e `Url` não podem ser realçados e, portanto, não podem ser usados para ativar um suplemento.
- Se for usada uma única regra, `Highlight` DEVERÁ ser definido como `all`.
- Se usar um tipo de regra `RuleCollection` com `Mode="AND"` para combinar várias regras, pelo menos uma das regras DEVERÁ ter o `Highlight` definido como `all`.
- Se usar um tipo de regra `RuleCollection` com `Mode="OR"` para combinar várias regras, todas as regras DEVERÃO ter o `Highlight` definido como `all`.

#### <a name="detectedentity-event-example"></a>Exemplo do evento DetectedEntity

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint>
```
