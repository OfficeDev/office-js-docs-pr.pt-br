---
title: Criar comandos de suplemento no manifesto para Excel, Word e PowerPoint
description: Use VersionOverrides no manifesto para definir comandos de suplemento para Excel, Word e PowerPoint. Use comandos de suplemento para criar elementos da interface do usu?rio, adicionar listas ou bot?es e executar a??es.
ms.date: 12/04/2017
ms.openlocfilehash: 95861fe0de6f0f56f6436b98cd7ad8dee510e82d
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-word-and-powerpoint"></a>Criar comandos de suplemento no manifesto para Excel, Word e PowerPoint


Use **[VersionOverrides](https://dev.office.com/reference/add-ins/manifest/versionoverrides)** no manifesto para definir comandos de suplemento para Excel, Word e PowerPoint. Os comandos de suplemento fornecem uma maneira f?cil de personalizar a interface do usu?rio padr?o do Office com elementos de interface do usu?rio especificados que executam a??es. Voc? pode usar comandos de suplemento para:
- Criar elementos de interface do usu?rio ou pontos de entrada que facilitam o uso da funcionalidade dos suplementos.  
  
- Adicionar bot?es ou uma lista suspensa de bot?es ? faixa de op??es.    
  
- Adicionar itens de menu individuais ? cada um contendo submenus opcionais ? aos menus de contexto espec?ficos (atalho).    
  
- Executar a??es quando seu comando de suplemento for escolhido. ? poss?vel:
    
  - Mostrar um ou mais suplementos de painel de tarefa com os quais os usu?rios podem interagir. Dentro do suplemento de painel de tarefa, ? poss?vel exibir HTML que use a malha da interface do usu?rio do Office para criar uma interface do usu?rio personalizada.
    
     *ou* 
      
  - Executar c?digo JavaScript, que normalmente ? executado sem exibir qualquer interface do usu?rio.
      
Este artigo descreve como editar seu manifesto para definir comandos de suplemento. O diagrama a seguir mostra a hierarquia de elementos usada para definir comandos de suplemento. Descrevemos esses elementos com mais detalhes neste artigo. 
      
A imagem a seguir representa uma vis?o geral dos elementos dos comandos de suplemento no manifesto. 
![Vis?o geral dos elementos dos comandos de suplemento no manifesto](../images/version-overrides.png)
 
## <a name="step-1-start-from-a-sample"></a>Etapa 1: iniciar usando uma amostra

? altamente recomend?vel iniciar usando uma das amostras fornecidas em [Amostras de comandos de suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Command-Sample). Como op??o, voc? pode criar seu pr?prio manifesto seguindo as etapas neste guia. ? poss?vel validar o manifesto usando o arquivo XSD no site de Amostras de comandos de suplemento do Office. N?o deixe de ler o artigo [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md), antes de usar os comandos de suplemento.

## <a name="step-2-create-a-task-pane-add-in"></a>Etapa 2: criar um suplemento de painel de tarefas

Para come?ar a usar os comandos de suplemento, primeiramente, ? preciso criar um suplemento de painel de tarefas e modificar o manifesto do suplemento, conforme descrito neste artigo. N?o ? poss?vel usar comandos de suplemento com suplementos de conte?do. Se for atualizar um manifesto existente, voc? dever? adicionar o **XML namespaces** apropriado, al?m do elemento **VersionOverrides** ao manifesto, conforme descrito na [Etapa 3: adicionar o elemento VersionOverrides](#step-3-add-versionoverrides-element).
   
O exemplo a seguir mostra o manifesto de um suplemento do Office 2013. N?o h? comandos de suplemento nesse manifesto porque n?o h? elemento **VersionOverrides**. O Office 2013 n?o d? suporte a comandos de suplemento, mas com a adi??o de **VersionOverrides** a esse manifesto, o suplemento ser? executado no Office 2013 e no Office 2016. No Office 2013, o suplemento n?o exibir? comandos de suplemento e usa o valor de **SourceLocation** para executar seu suplemento como um ?nico suplemento de painel de tarefas. No Office 2016, se nenhum elemento **VersionOverrides** estiver inclu?do, **SourceLocation** ser? usado para executar o suplemento. Entretanto, se voc? incluir **VersionOverrides**, o suplemento exibir? apenas os comandos de suplemento e n?o exibir? o suplemento como um ?nico suplemento de painel de tarefas.
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="~remoteAppUrl/Images/Icon_32.png" />
 
  <AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
    <AppDomain>AppDomain3</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/Pages/Home.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>

 <!-- The VersionOverrides element is inserted at this location in the manifest. -->

</OfficeApp>
```

## <a name="step-3-add-versionoverrides-element"></a>Etapa 3: adicionar o elemento VersionOverrides
O elemento **VersionOverrides** ? o elemento raiz que cont?m a defini??o do comando de suplemento. **VersionOverrides** ? um elemento filho do elemento **OfficeApp** no manifesto. A tabela a seguir lista os atributos do elemento **VersionOverrides**.

|**Atributo**|**Descri??o**|
|:-----|:-----|
|**xmlns** <br/> | Obrigat?rio. O local do esquema, que deve ser "http://schemas.microsoft.com/office/taskpaneappversionoverrides". <br/> |
|**xsi:type** <br/> |Obrigat?rio. A vers?o do esquema. A vers?o descrita neste artigo ? "VersionOverridesV1_0".  <br/> |
   
A tabela a seguir identifica os elementos filho de **VersionOverrides**.
  
|**Elemento**|**Descri??o**|
|:-----|:-----|
|**Descri??o** <br/> |Opcional. Descreve o suplemento. Esse elemento filho **Description** substitui um elemento **Description** anterior na parte pai do manifesto. O atributo **resid** para esse elemento **Description** ? definido como a **id** de um elemento **String**. O elemento **String** cont?m o texto para **Description**. <br/> |
|**Requisitos** <br/> |Opcional. Especifica o conjunto de requisitos m?nimos e a vers?o do Office.js exigida pelo suplemento. Esse elemento filho **Requirements** substitui o elemento **Requirements** na parte pai do manifesto. Para saber mais, confira [Especificar requisitos de API e hosts do Office](../develop/specify-office-hosts-and-api-requirements.md).  <br/> |
|**Hosts** <br/> |Obrigat?rio. Especifica um conjunto de hosts do Office. O elemento filho **Hosts** substitui o elemento **Hosts** na parte pai do manifesto. Voc? deve incluir um conjunto de atributos **xsi:type** como "Pasta de trabalho" ou "Documento". <br/> |
|**Recursos** <br/> |Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) referenciado por outros elementos de manifesto. Por exemplo, o valor do elemento **Description** refere-se a um elemento filho em **Resources**. O elemento **Resources** ? descrito na [Etapa 7: adicionar o elemento Resources](#step-7-add-the-resources-element) mais adiante neste artigo. <br/> |
   
O exemplo a seguir mostra como usar o elemento **VersionOverrides** e seus elementos filho.

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information about requirement sets -->
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
        <!-- add information about form factors -->
      </Host>
      <Host xsi:type="Document">
        <!-- add information about form factors -->
      </Host>
    </Hosts>
    <Resources> 
      <!-- add information about resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="step-4-add-hosts-host-and-desktopformfactor-elements"></a>Etapa 4: adicionar os elementos Hosts, Host e DesktopFormFactor

O elemento **Hosts** cont?m um ou mais elementos **Host**. Um elemento **Host** especifica um determinado host do Office. O elemento **Host** cont?m elementos filho que especificam os comandos de suplemento que ser?o exibidos ap?s a instala??o do suplemento nesse host do Office. Para mostrar os mesmos comandos de suplemento em dois ou mais hosts do Office diferentes, voc? deve duplicar os elementos filho em cada **Host**.
       
O elemento **DesktopFormFactor** especifica as configura??es para um suplemento que ? executado no Office, na ?rea de trabalho do Windows, e no Office Online (no navegador).
      
Veja a seguir um exemplo dos elementos **Hosts**, **Host** e **DesktopFormFactor**.

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
  ...
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>

              <!-- information about FunctionFile and ExtensionPoint -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
  ...
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="step-5-add-the-functionfile-element"></a>Etapa 5: adicionar o elemento FunctionFile

O elemento **FunctionFile** especifica um arquivo que cont?m o c?digo JavaScript a ser executado quando um comando de suplemento usa a a??o **ExecuteFunction** (confira [Controles de bot?o](https://dev.office.com/reference/add-ins/manifest/control#Button-control) para obter uma descri??o). O atributo **resid** do elemento **FunctionFile** ? definido como um arquivo HTML que inclui todos os arquivos JavaScript exigidos por seus comandos de suplemento. Voc? n?o pode criar um v?nculo diretamente com um arquivo JavaScript, mas somente com um arquivo HTML. O nome do arquivo ? especificado como um elemento **Url** no elemento **Resources**.
        
Veja a seguir um exemplo do elemento **FunctionFile**.
  
```xml
<DesktopFormFactor>
    <FunctionFile resid="residDesktopFuncUrl" />
    <ExtensionPoint xsi:type="PrimaryCommandSurface">
      <!-- information about this extension point -->
    </ExtensionPoint> 

    <!-- You can define more than one ExtensionPoint element as needed -->
</DesktopFormFactor>
```

> [!IMPORTANT]
> Verifique se seu c?digo JavaScript chama `Office.initialize`. 
   
O JavaScript no arquivo HTML referenciado pelo elemento **FunctionFile** deve chamar `Office.initialize`. O elemento **FunctionName** (confira [Controles de bot?o](https://dev.office.com/reference/add-ins/manifest/control#Button-control) para obter uma descri??o) usa as fun??es em **FunctionFile**.
     
O c?digo a seguir mostra como implementar a fun??o usada por **FunctionName**.

```javascript

<script>
    // The initialize function must be run each time a new page is loaded.
    (function () {
        Office.initialize = function (reason) {
            // If you need to initialize something you can do so here. 
        };
    })();

    // Your function must be in the global namespace.
    function writeText(event) {

        // Implement your custom code here. The following code is a simple example.  
        Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
            function (asyncResult) {
                var error = asyncResult.error;
                if (asyncResult.status === "failed") {
                    // Show error message. 
                }
                else {
                    // Show success message.
                }
            });
        
        // Calling event.completed is required. event.completed lets the platform know that processing has completed. 
        event.completed();
    }
</script>
```

> [!IMPORTANT]
> A chamada para **event.completed** sinaliza que o evento foi manipulado. Quando uma fun??o ? chamada v?rias vezes, por exemplo, com v?rios cliques no mesmo comando de suplemento, todos os eventos s?o enfileirados automaticamente. O primeiro evento ? executado automaticamente, enquanto os outros eventos permanecem na fila. Quando sua fun??o chama **event.completed**, a pr?xima chamada em fila para essa fun??o ? executada. Voc? deve implementar **event.completed**; caso contr?rio, sua fun??o n?o ser? executada.
 
## <a name="step-6-add-extensionpoint-elements"></a>Etapa 6: adicionar elementos do ExtensionPoint

O elemento **ExtensionPoint** define onde os comandos de suplemento devem aparecer na interface do usu?rio do Office. Voc? pode definir os elementos **ExtensionPoint** com estes valores de **xsi:type**:
   
- **PrimaryCommandSurface**, que se refere ? faixa de op??es no Office.
     
- **ContextMenu**, que ? o menu de atalho exibido quando voc? clica com o bot?o direito na interface do usu?rio do Office.
    
Os exemplos a seguir mostram como usar o elemento **ExtensionPoint** com os valores de atributo **PrimaryCommandSurface** e **ContextMenu** e os elementos filho que devem ser usados com cada um.
    
> [!IMPORTANT]
> Para os elementos que cont?m um atributo ID, forne?a uma ID exclusiva. Recomendamos usar o nome da sua empresa com a ID. Por exemplo, use o seguinte formato:`<CustomTab id="mycompanyname.mygroupname">` 
  
```xml
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

|**Elemento**|**Descri??o**|
|:-----|:-----|
|**CustomTab** <br/> |Obrigat?rio se voc? quiser adicionar uma guia personalizada ? faixa de op??es (usando **PrimaryCommandSurface**). Se voc? usar o elemento **CustomTab**, o elemento **OfficeTab** n?o poder? ser usado. O atributo **id** ? obrigat?rio. <br/> |
|**OfficeTab** <br/> |Obrigat?rio se voc? quiser estender uma guia de faixa de op??es padr?o do Office (usando **PrimaryCommandSurface**). Se voc? usar o elemento **OfficeTab**, o elemento **CustomTab** n?o poder? ser usado. <br/> Para obter mais valores de guia a serem usados com o atributo **id**, confira [Valores de guia para guias de faixa de op??es padr?o do Office](https://dev.office.com/reference/add-ins/manifest/officetab).  <br/> |
|**OfficeMenu** <br/> | Obrigat?rio se voc? estiver adicionando comandos de suplemento a um menu de contexto padr?o (usando **ContextMenu**). O atributo **id** deve ser definido como: <br/> **ContextMenuText** para Excel ou Word. Exibe o item no menu de contexto quando o texto ? selecionado e o usu?rio clica com o bot?o direito do mouse no texto selecionado. <br/> **ContextMenuCell** para Excel. Exibe o item no menu de contexto quando o usu?rio clica com o bot?o direito do mouse em uma c?lula na planilha. <br/> |
|**Grupo** <br/> |Um grupo de pontos de extens?o de interface do usu?rio em uma guia. Um grupo pode ter at? seis controles. O atributo **id** ? obrigat?rio. ? uma cadeia de caracteres com, no m?ximo, 125 caracteres. <br/> |
|**R?tulo** <br/> |Obrigat?rio. O r?tulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** ? um elemento filho do elemento **ShortStrings**, que ? elemento filho do elemento **Resources**. <br/> |
|**?cone** <br/> |Obrigat?rio. Especifica o ?cone do grupo a ser usado em dispositivos de fator forma pequeno ou quando muitos bot?es s?o exibidos. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Image**. O elemento **Image** ? um elemento filho do elemento **Images**, que ? um elemento filho do elemento **Resources**. O atributo **size** fornece o tamanho da imagem em pixels. Tr?s tamanhos de imagem s?o obrigat?rios: 16, 32 e 80 pixels. Tamb?m h? suporte para cinco tamanhos opcionais: 20, 24, 40, 48 e 64 pixels. <br/> |
|**Dica de ferramenta** <br/> |Opcional. A dica de ferramenta do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** ? um elemento filho do elemento **LongStrings**, que ? um elemento filho do elemento **Resources**. <br/> |
|**Controle** <br/> |Cada grupo exige pelo menos um controle. Um elemento **Control** pode ser um **Button** ou um **Menu**. Use **Menu** para especificar uma lista suspensa de controles de bot?o. Atualmente, h? suporte apenas para bot?es e menus. Confira as seguintes se??es [Controles de bot?o](https://dev.office.com/reference/add-ins/manifest/control) e [Controles de menu](https://dev.office.com/reference/add-ins/manifest/control) para saber mais. <br/>**Observa??o:** para facilitar a solu??o de problemas, recomendamos adicionar um elemento **Control** e os elementos filho **Resources** relacionados, um de cada vez.          |
   

### <a name="button-controls"></a>Controles de bot?o
Um bot?o executa uma ?nica a??o quando o usu?rio o seleciona. Pode ser a execu??o de uma fun??o JavaScript ou a exibi??o de um painel de tarefas. O exemplo a seguir mostra como definir dois bot?es. O primeiro bot?o executa uma fun??o JavaScript sem mostrar uma interface do usu?rio e o segundo bot?o mostra um painel de tarefas. No elemento **Control**:        

- O atributo **type** ? obrigat?rio e deve ser definido como **Button**.
    
- O atributo **id** do elemento **Control** ? uma cadeia de caracteres com, no m?ximo, 125 caracteres.
    
```xml
<!-- Define a control that calls a JavaScript function. -->
<Control xsi:type="Button" id="Button1Id1">
  <Label resid="residLabel" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getData</FunctionName>
  </Action>
</Control>

<!-- Define a control that shows a task pane. -->
<Control xsi:type="Button" id="Button2Id1">
  <Label resid="residLabel2" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon2_32x32" />
    <bt:Image size="32" resid="icon2_32x32" />
    <bt:Image size="80" resid="icon2_32x32" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="residUnitConverterUrl" />
  </Action>
</Control>
```

|**Elementos**|**Descri??o**|
|:-----|:-----|
|**R?tulo** <br/> |Obrigat?rio. O texto do bot?o. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** ? um elemento filho do elemento **ShortStrings**, que ? elemento filho do elemento **Resources**. <br/> |
|**Dica de ferramenta** <br/> |Opcional. A dica de ferramenta do bot?o. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** ? um elemento filho do elemento **LongStrings**, que ? um elemento filho do elemento **Resources**. <br/> |
|**Dica detalhada** <br/> | Obrigat?rio. A superdica para esse bot?o, que ? definida pelos seguintes itens: <br/> **T?tulo** <br/>  Obrigat?rio. O texto da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** ? um elemento filho do elemento **ShortStrings**, que ? elemento filho do elemento **Resources**. <br/> **Descri??o** <br/>  Obrigat?rio. A descri??o da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** ? um elemento filho do elemento **LongStrings**, que ? um elemento filho do elemento **Resources**. <br/> |
|**?cone** <br/> | Obrigat?rio. Cont?m os elementos **Image** para o bot?o. Arquivos de imagem devem estar no formato .png. <br/> **Imagem** <br/>  Define uma imagem a ser exibida no bot?o. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Image**. O elemento **Image** ? um elemento filho do elemento **Images**, que ? um elemento filho do elemento **Resources**. O atributo **size** indica o tamanho em pixels da imagem. Tr?s tamanhos de imagem s?o obrigat?rios: 16, 32 e 80 pixels. Tamb?m h? suporte para cinco tamanhos opcionais: 20, 24, 40, 48 e 64 pixels. <br/> |
|**A??o** <br/> | Obrigat?rio. Especifica a a??o a ser executada quando o usu?rio seleciona o bot?o. Voc? pode especificar um dos seguintes valores para o atributo **xsi:type**: <br/> **ExecuteFunction**, que executa uma fun??o JavaScript localizada no arquivo referenciado por **FunctionFile**. **ExecuteFunction** n?o exibe uma interface do usu?rio. O elemento filho **FunctionName** especifica o nome da fun??o a ser executada. <br/> **ShowTaskPane**, que mostra um suplemento de painel de tarefas. O elemento filho **SourceLocation** especifica o local do arquivo de origem do suplemento de painel de tarefas a ser exibido. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Url** no elemento **Urls** do elemento **Resources**. <br/> |
   

### <a name="menu-controls"></a>Controles de menu
Um controle **Menu** pode ser usado com **PrimaryCommandSurface** ou **ContextMenu** e define:
  
- Um item de menu no n?vel raiz.
   
- Uma lista de itens de submenu.
 
Quando usado com **PrimaryCommandSurface**, o item de menu raiz ? exibido como um bot?o na faixa de op??es. Quando o bot?o ? selecionado, o submenu ? exibido como uma lista suspensa. Quando usado com **ContextMenu**, um item de menu com um submenu ? inserido no menu de contexto. Em ambos os casos, cada item de submenu pode executar uma fun??o JavaScript ou mostrar um painel de tarefas. Somente um n?vel de submenus ? compat?vel no momento.
       
O exemplo a seguir mostra como definir um item de menu com dois itens de submenu. O primeiro item do submenu mostra um painel de tarefas e o segundo item do submenu executa uma fun??o JavaScript. No elemento **Control**:
    
- O atributo **xsi:type** ? obrigat?rio e deve ser definido como **Menu**.
  
- O atributo **id** ? uma cadeia de caracteres com, no m?ximo, 125 caracteres.
    
```xml

<Control xsi:type="Menu" id="TestMenu2">
  <Label resid="residLabel3" />
  <Tooltip resid="residToolTip" />
  <Supertip>
    <Title resid="residLabel" />
    <Description resid="residToolTip" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="icon1_32x32" />
    <bt:Image size="32" resid="icon1_32x32" />
    <bt:Image size="80" resid="icon1_32x32" />
  </Icon>
  <Items>
    <Item id="showGallery2">
      <Label resid="residLabel3"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon1_32x32" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_32x32" />
      </Icon>
      <Action xsi:type="ShowTaskpane">
        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
        <SourceLocation resid="residUnitConverterUrl" />
      </Action>
    </Item>
    <Item id="showGallery3">
      <Label resid="residLabel5"/>
      <Supertip>
        <Title resid="residLabel" />
        <Description resid="residToolTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="icon4_32x32" />
        <bt:Image size="32" resid="icon4_32x32" />
        <bt:Image size="80" resid="icon4_32x32" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getButton</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

|**Elementos**|**Descri??o**|
|:-----|:-----|
|**R?tulo** <br/> |Obrigat?rio. O texto do item de menu raiz. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** ? um elemento filho do elemento **ShortStrings**, que ? elemento filho do elemento **Resources**. <br/> |
|**Dica de ferramenta** <br/> |Opcional. A dica de ferramenta do menu. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** ? um elemento filho do elemento **LongStrings**, que ? um elemento filho do elemento **Resources**. <br/> |
|**Dica detalhada** <br/> | Obrigat?rio. A superdica para o menu, que ? definida pelos seguintes itens: <br/> **T?tulo** <br/>  Obrigat?rio. O texto da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** ? um elemento filho do elemento **ShortStrings**, que ? elemento filho do elemento **Resources**. <br/> **Descri??o** <br/>  Obrigat?rio. A descri??o da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** ? um elemento filho do elemento **LongStrings**, que ? um elemento filho do elemento **Resources**. <br/> |
|**?cone** <br/> | Obrigat?rio. Cont?m os elementos **Image** para o menu. Arquivos de imagem devem estar no formato .png. <br/> **Imagem** <br/>  Uma imagem para o menu. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Image**. O elemento **Image** ? um elemento filho do elemento **Images**, que ? um elemento filho do elemento **Resources**. O atributo **size** indica o tamanho em pixels da imagem. Tr?s tamanhos de imagem, em pixels, s?o obrigat?rios: 16, 32 e 80 pixels. Cinco tamanhos opcionais, em pixels, tamb?m t?m suporte: 20, 24, 40, 48 e 64 pixels. <br/> |
|**Itens** <br/> |Obrigat?rio. Cont?m os elementos **Item** para cada item do submenu. Cada elemento **Item** cont?m os mesmos elementos filho que [Controles de bot?o](https://dev.office.com/reference/add-ins/manifest/control).  <br/> |
   
## <a name="step-7-add-the-resources-element"></a>Etapa 7: adicionar o elemento Resources

O elemento **Resources** cont?m recursos usados pelos diferentes elementos filho do elemento **VersionOverrides**. Resources inclui ?cones, cadeias de caracteres e URLs. Um elemento no manifesto pode usar um recurso fazendo refer?ncia a **id** do recurso. O uso da **id** ajuda a organizar o manifesto, especialmente quando h? vers?es diferentes do recurso para localidades diferentes. Uma **id** tem no m?ximo 32 caracteres.
  
    
    
Veja a seguir um exemplo de como usar o elemento **Resources**. Cada recurso pode ter um ou mais elementos filho **Override** para definir um recurso diferente para uma localidade espec?fica.


```xml
<Resources>
  <bt:Images>
    <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp16-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp32-icon_default.png" />
    </bt:Image>
    <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/Images/icon_default.png">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Images/ja-jp80-icon_default.png" />
    </bt:Image>        
  </bt:Images>
  <bt:Urls>
    <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
      <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
    </bt:Url>        
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="residLabel" DefaultValue="GetData">
      <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
    </bt:String>      
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="residToolTip" DefaultValue="Get data for your document.">
      <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
    </bt:String>
  </bt:LongStrings>
</Resources>
```

|**Recurso**|**Descri??o**|
|:-----|:-----|
|**Imagens**/ **Imagem** <br/> | Fornece a URL HTTPS para um arquivo de imagem. Cada imagem precisa definir os tr?s tamanhos de imagem necess?rios: <br/>  16?16 <br/>  32?32 <br/>  80?80 <br/>  Os seguintes tamanhos de imagem tamb?m t?m suporte, mas n?o s?o obrigat?rios: <br/>  20?20 <br/>  24?24 <br/>  40?40 <br/>  48?48 <br/>  64?64 <br/> |
|**Urls**/ **Url** <br/> |Fornece um local para a URL HTTPS. Uma URL pode ter no m?ximo 2048 caracteres.  <br/> |
|**Sequ?ncias de caracteres curtas**/ **Sequ?ncia de caracteres** <br/> |O texto para os elementos **Label** e **Title**. Cada **String** cont?m no m?ximo 125 caracteres. <br/> |
|**Sequ?ncias de caracteres longas**/ **Sequ?ncia de caracteres** <br/> |O texto para os elementos **Tooltip** e **Description**. Cada **String** cont?m no m?ximo 250 caracteres. <br/> |
   
> [!NOTE] 
> Use o protocolo SSL (Secure Sockets Layer) para todas as URLs nos elementos **Image** e **Url**.

### <a name="tab-values-for-default-office-ribbon-tabs"></a>Valores para as guias de faixa de op??es padr?o do Office
No Excel e no Word, ? poss?vel adicionar seus comandos de suplemento na faixa de op??es usando as guias padr?o da interface de usu?rio do Office. A tabela a seguir lista os valores que podem ser usados para o atributo **id** do elemento **OfficeTab**. Os valores da guia diferenciam mai?sculas de min?sculas.

|**Aplicativo host do Office**|**Valores de guia**|
|:-----|:-----|
|Excel  <br/> |**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval** <br/> |
|Word  <br/> |**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation** <br/> |
|PowerPoint  <br/> |**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**          <br/> |
   
## <a name="see-also"></a>Veja tamb?m

-  [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md)      
