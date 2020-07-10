---
title: Elemento ExtensionPoint no arquivo de manifesto
description: Define onde um suplemento expõe a funcionalidade na interface de usuário do Office.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 44824e0c74b35105833f1f05cdda87bc873a4427
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094453"
---
# <a name="extensionpoint-element"></a>Elemento ExtensionPoint

 Define onde um suplemento expõe a funcionalidade na interface de usuário do Office. O elemento **ExtensionPoint** é um elemento filho de [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Sim  | O tipo de ponto de extensão que está sendo definido.|

## <a name="extension-points-for-excel-only"></a>Pontos de extensão somente para Excel

- **CustomFunctions**: uma função personalizada escrita em JavaScript para Excel.

[Este exemplo de código XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) mostra como usar o elemento **ExtensionPoint** com o valor do atributo **CustomFunctions** e os elementos filhos a serem usados.

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a>Pontos de extensão para comandos de suplemento do Word, Excel, PowerPoint e OneNote

- **PrimaryCommandSurface**, que se refere à faixa de opções no Office.
- **ContextMenu**, que é o menu de atalho exibido ao clicar com o botão direito do mouse na interface de usuário do Office.

Os exemplos a seguir mostram como usar o elemento **ExtensionPoint** com os valores de atributo **PrimaryCommandSurface** e **ContextMenu** e os elementos filho que devem ser usados com cada um.

> [!IMPORTANT]
> Forneça uma ID exclusiva para os elementos que contêm um atributo ID. É recomendável usar o nome de sua empresa com a ID. Por exemplo, use o formato a seguir. <CustomTab id="mycompanyname.mygroupname">

```XML
<ExtensionPoint xsi:type="PrimaryCommandSurface">
          <CustomTab id="Contoso Tab">
          <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
            <!-- <OfficeTab id="TabData"> -->
            <Label resid="residLabel4" />
            <Group id="Group1Id12">
              <Label resid="residLabel4" />
              <Icon>
                <bt:Image size="16" resid="icon1_32x32" />
                <bt:Image size="32" resid="icon1_32x32" />
                <bt:Image size="80" resid="icon1_32x32" />
              </Icon>
              <Tooltip resid="residToolTip" />
              <Control xsi:type="Button" id="Button1Id1">

                  <!-- information about the control -->
              </Control>
              <!-- other controls, as needed -->
            </Group>
          </CustomTab>
        </ExtensionPoint>

      <ExtensionPoint xsi:type="ContextMenu">
        <OfficeMenu id="ContextMenuCell">
          <Control xsi:type="Menu" id="ContextMenu2">
                  <!-- information about the control -->
          </Control>
          <!-- other controls, as needed -->
        </OfficeMenu>
        </ExtensionPoint>
```

#### <a name="child-elements"></a>Elementos filho
 
|**Elemento**|**Descrição**|
|:-----|:-----|
|**CustomTab**|Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.|
|**OfficeTab**|Obrigatório se você deseja estender uma guia padrão da faixa de opções do aplicativo do Office (usando **PrimaryCommandSurface**). Se você usar o elemento **OfficeTab**, o elemento **CustomTab** não poderá ser usado. Para saber mais, confira [OfficeTab](officetab.md).|
|**OfficeMenu**|Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: <br/> - **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. <br/> - **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.|
|**Group**|A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.|
|**Label**|Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.|
|**Icon**|Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.|
|**Tooltip**|Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.|
|**Control**|Cada grupo exige pelo menos um controle. Um elemento **Control** pode ser um **Button** ou um **Menu**. Use **Menu** para especificar uma lista suspensa de controles de botão. Atualmente, há suporte apenas para botões e menus. Confira as seções [Controles de botão](control.md#button-control) e [Controles de menu](control.md#menu-dropdown-button-controls) para saber mais.<br/>**Observação:**  Para facilitar a solução de problemas, recomendamos que um elemento **Control** e os elementos filho de **recursos** relacionados sejam adicionados um de cada vez.|
|**Script**|Links para o arquivo JavaScript com a definição de função personalizada e o código de registro Esse elemento não é usado na Visualização do Desenvolvedor. Em vez disso, a página HTML é responsável por carregar todos os arquivos JavaScript.|
|**Page**|Links para a página HTML de suas funções personalizadas.|

## <a name="extension-points-for-outlook"></a>Pontos de extensão para Outlook

- [MessageReadCommandSurface](#messagereadcommandsurface)
- [MessageComposeCommandSurface](#messagecomposecommandsurface)
- [AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)
- [AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)
- [Module](#module) (Só pode ser usado em [DesktopFormFactor](desktopformfactor.md)).
- [MobileMessageReadCommandSurface](#mobilemessagereadcommandsurface)
- [MobileOnlineMeetingCommandSurface](#mobileonlinemeetingcommandsurface-preview)
- [LaunchEvent](#launchevent-preview)
- [Eventos](#events)
- [DetectedEntity](#detectedentity)

### <a name="messagereadcommandsurface"></a>MessageReadCommandSurface

This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.

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
  <CustomTab id="TabCustom1">
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
  <CustomTab id="TabCustom1">
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
  <CustomTab id="TabCustom1">
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
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a>Module

Este ponto de extensão coloca botões na faixa de opções para a extensão do módulo.

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
  <Group id="mobileGroupID">
    <Label resid="residAppName"/>
      <Control id="mobileButton1" xsi:type="MobileButton">
        <!-- Control definition -->
      </Control>
  </Group>
</ExtensionPoint>
```

### <a name="mobileonlinemeetingcommandsurface-preview"></a>MobileOnlineMeetingCommandSurface (visualização)

> [!NOTE]
> Este ponto de extensão só tem suporte na [Visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Android com uma assinatura do Microsoft 365.

Este ponto de extensão coloca uma alternância apropriada de modo na superfície de comando para um compromisso no fator de forma móvel. Um organizador da reunião pode criar uma reunião online. Um participante pode ingressar na reunião online subsequentemente. Para saber mais sobre esse cenário, confira o artigo [criar um suplemento do Outlook Mobile para um provedor de reunião online](../../outlook/online-meeting.md) .

#### <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
|  [Control](control.md) |  Adiciona um botão à superfície de comando.  |

`ExtensionPoint`elementos desse tipo só podem ter um elemento filho: um `Control` elemento.

O `Control` elemento contido neste ponto de extensão deve ter o `xsi:type` atributo definido como `MobileButton` .

As `Icon` imagens devem estar em escala de cinza usando `#919191` o código hex ou seu equivalente em [outros formatos de cor](https://convertingcolors.com/hex-color-919191.html).

#### <a name="example"></a>Exemplo

```xml
<ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
  <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
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

### <a name="launchevent-preview"></a>LaunchEvent (visualização)

> [!NOTE]
> Este ponto de extensão só tem suporte na [Visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Outlook na Web com uma assinatura do Microsoft 365.

Este ponto de extensão permite que um suplemento seja ativado com base nos eventos suportados no fator forma da área de trabalho. Atualmente, os únicos eventos com suporte são `OnNewMessageCompose` e `OnNewAppointmentOrganizer` . Para saber mais sobre esse cenário, confira o artigo [Configurar o suplemento do Outlook para ativação baseada em eventos](../../outlook/autolaunch.md) .

#### <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
| [LaunchEvents](launchevents.md) |  Lista de [LaunchEvent](launchevent.md) para a ativação baseada em evento.  |
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

Este ponto de extensão adiciona um manipulador de eventos para um evento especificado. Para obter mais informações sobre como usar esse ponto de extensão, consulte o [recurso ao enviar para suplementos do Outlook](../../outlook/outlook-on-send-addins.md).

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

O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.

> [!NOTE]
> Este tipo de elemento está disponível para [ clientes do Outlook que ofereçam suporte a conjuntos de requisitos 1.6 e posteriores](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).

|  Elemento |  Descrição  |
|:-----|:-----|
|  [Label](#label) |  Especifica o rótulo para o suplemento na janela contextual.  |
|  [SourceLocation](sourcelocation.md) |  Especifica a URL para a janela contextual.  |
|  [Rule](rule.md) |  Especifica a regra ou regras que determinam quando um suplemento é ativado.  |

#### <a name="label"></a>Label

Obrigatório. O rótulo do grupo. O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .

#### <a name="highlight-requirements"></a>Requisitos de realce

The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.

However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.

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
