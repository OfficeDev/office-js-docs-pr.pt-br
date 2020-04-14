---
title: Elemento ExtensionPoint no arquivo de manifesto
description: Define onde um suplemento expõe a funcionalidade na interface de usuário do Office.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 40b5ab8c2dcae01238854f5a3bfcc599f6b01a9b
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241032"
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
|**CustomTab**|Obrigatório se você quiser adicionar uma guia personalizada à faixa de opções (usando **PrimaryCommandSurface**). Se você usar o elemento **CustomTab**, o elemento **OfficeTab** não poderá ser usado. O atributo **id** é obrigatório. |
|**OfficeTab**|Obrigatório se você quiser estender uma guia de faixa de opções padrão do Office (usando **PrimaryCommandSurface**). Se você usar o elemento **OfficeTab**, o elemento **CustomTab** não poderá ser usado. Para saber mais, confira [OfficeTab](officetab.md).|
|**OfficeMenu**|Obrigatório se você estiver adicionando comandos de suplemento a um menu de contexto padrão (usando **ContextMenu**). O atributo **id** deve ser definido como: <br/> - **ContextMenuText** para o Excel ou Word. Exibe o item no menu de contexto quando o texto for selecionado e o usuário clicar com o botão direito do mouse no texto selecionado. <br/> - **ContextMenuCell** para Excel. Exibe o item no menu de contexto quando o usuário clica com o botão direito do mouse em uma célula na planilha.|
|**Group**|Um grupo de pontos de extensão de interface do usuário em uma guia. Um grupo pode ter até seis controles. O atributo **id** é obrigatório. É uma cadeia de caracteres com, no máximo, 125 caracteres. |
|**Label**|Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **ShortStrings**, que é elemento filho do elemento **Resources**. |
|**Icon**|Obrigatório. Especifica o ícone do grupo a ser usado em dispositivos de fator forma pequeno ou quando muitos botões são exibidos. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Image**. O elemento **Image** é um elemento filho do elemento **Images**, que é um elemento filho do elemento **Resources**. O atributo **size** fornece o tamanho da imagem em pixels. Três tamanhos de imagem são obrigatórios: 16, 32 e 80 pixels. Também há suporte para cinco tamanhos opcionais: 20, 24, 40, 48 e 64 pixels. |
|**Tooltip**|Opcional. A dica de ferramenta do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **LongStrings**, que é um elemento filho do elemento **Resources**. |
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
> Este ponto de extensão só tem suporte na [Visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Android com uma assinatura do Office 365.

Este ponto de extensão coloca uma alternância apropriada de modo na superfície de comando para um compromisso no fator de forma móvel. Um organizador da reunião pode criar uma reunião online. Um participante pode ingressar na reunião online subsequentemente. Para saber mais sobre esse cenário, confira o artigo [criar um suplemento do Outlook Mobile para um provedor de reunião online](../../outlook/online-meeting.md) .

#### <a name="child-elements"></a>Elementos filho

|  Elemento |  Descrição  |
|:-----|:-----|
|  [Control](control.md) |  Adiciona um botão à superfície de comando.  |

`ExtensionPoint`elementos desse tipo só podem ter um elemento filho: um `Control` elemento.

O `Control` elemento contido neste ponto de extensão deve ter o `xsi:type` atributo definido como `MobileButton`.

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
      <bt:Image resid="UiLessIcon" size="32" scale="2" />
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

### <a name="events"></a>Eventos

Este ponto de extensão adiciona um manipulador de eventos para um evento especificado.

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
