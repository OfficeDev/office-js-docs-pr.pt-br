---
title: Criar comandos de suplemento no manifesto para Excel, Word e PowerPoint
description: Use VersionOverrides em seu manifesto para definir comandos de suplemento para Excel, PowerPoint e Word. Use comandos de suplemento para criar elementos da interface do usuário, adicionar listas ou botões e executar ações.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 82e921fef7ba37deaa2b20f9f2aa684304cd44ba
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810180"
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-powerpoint-and-word"></a>Criar comandos de suplemento no manifesto para Excel, Word e PowerPoint

> [!NOTE]
> Os comandos de suplemento também são compatíveis com o Outlook. Para obter mais informações, confira [Comandos de suplemento para o Outlook](../outlook/add-in-commands-for-outlook.md)

Use **[VersionOverrides](/javascript/api/manifest/versionoverrides)** em seu manifesto para definir comandos de suplemento para Excel, PowerPoint e Word. Os comandos de suplemento fornecem uma maneira fácil de personalizar a interface do usuário padrão do Office com elementos de interface do usuário especificados que executam ações. Para obter uma introdução aos comandos de suplemento, consulte [Comandos de suplemento para Excel, PowerPoint e Word](../design/add-in-commands.md).

Este artigo descreve como editar seu manifesto para definir comandos de suplemento e como criar o código para [comandos de função](../design/add-in-commands.md#types-of-add-in-commands). O diagrama a seguir mostra a hierarquia de elementos usada para definir comandos de suplemento. Descrevemos esses elementos com mais detalhes neste artigo.

![Visão geral dos elementos de comandos de suplemento no manifesto. O nó superior aqui é VersionOverrides com hosts e recursos para crianças. Em Hosts estão Host e DesktopFormFactor. Em DesktopFormFactor estão FunctionFile e ExtensionPoint. Em ExtensionPoint estão CustomTab ou OfficeTab e Menu do Office. Em CustomTab ou Guia do Office estão o Grupo e, em seguida, Controle e Ação. Em Menu do Office estão Controle e Ação. Em Recursos (filho de VersionOverrides) estão Imagens, Urls, ShortStrings e LongStrings.](../images/version-overrides.png)

## <a name="step-1-create-the-project"></a>Etapa 1: Criar o projeto

Recomendamos que você crie um projeto seguindo um dos inícios rápidos, como [Criar um suplemento de painel de tarefas do Excel](../quickstarts/excel-quickstart-jquery.md). Cada início rápido para Excel, PowerPoint e Word gera um projeto que já contém um comando de suplemento (botão) para mostrar o painel de tarefas. Verifique se você leu [comandos de suplemento para Excel, PowerPoint e Word](../design/add-in-commands.md) antes de usar comandos de suplemento.

## <a name="step-2-create-a-task-pane-add-in"></a>Etapa 2: criar um suplemento de painel de tarefas

Para começar a usar os comandos de suplemento, primeiramente, é preciso criar um suplemento de painel de tarefas e modificar o manifesto do suplemento, conforme descrito neste artigo. Você não pode usar comandos de suplemento com suplementos de conteúdo. Se você estiver atualizando um manifesto existente, deverá adicionar os **namespaces XML apropriados** , bem como adicionar o **\<VersionOverrides\>** elemento ao manifesto, conforme descrito na [Etapa 3: Adicionar elemento VersionOverrides](#step-3-add-versionoverrides-element).

O exemplo a seguir mostra o manifesto de um suplemento do Office 2013. Não há comandos de suplemento neste manifesto porque não há nenhum **\<VersionOverrides\>** elemento. O Office 2013 não dá suporte a comandos de suplemento, mas, ao adicionar **\<VersionOverrides\>** a esse manifesto, seu suplemento será executado no Office 2013 e no Office 2016. No Office 2013, seu suplemento não exibirá comandos de suplemento e usará o valor de **\<SourceLocation\>** para executar seu suplemento como um único suplemento de painel de tarefas. No Office 2016, se nenhum **\<VersionOverrides\>** elemento for incluído, o painel de tarefas do suplemento será aberto automaticamente para a URL especificada em **\<SourceLocation\>**. No entanto, se você incluir **\<VersionOverrides\>**, o suplemento exibirá apenas os comandos de suplemento e não exibirá inicialmente o painel de tarefas do suplemento.
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="https://www.contoso.com/Images/Icon_32.png" />
  <SupportUrl DefaultValue="https://www.contoso.com/contact" />
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

O **\<VersionOverrides\>** elemento é o elemento raiz que contém a definição do comando de suplemento. **\<VersionOverrides\>** é um elemento filho do **\<OfficeApp\>** elemento no manifesto. A tabela a seguir lista os atributos do **\<VersionOverrides\>** elemento.

|Atributo|Descrição|
|:-----|:-----|
|**xmlns** <br/> | Obrigatório. O local do esquema, que deve ser `http://schemas.microsoft.com/office/taskpaneappversionoverrides`. <br/> |
|**xsi:type** <br/> |Obrigatório. A versão do esquema. A versão descrita neste artigo é "VersionOverridesV1_0".  <br/> |

A tabela a seguir identifica os elementos filho de **\<VersionOverrides\>**.
  
|Elemento|Descrição|
|:-----|:-----|
|**\<Description\>** <br/> |Opcional. Descreve o suplemento. Esse elemento filho **\<Description\>** substitui um elemento anterior **\<Description\>** na parte pai do manifesto. O atributo **resid** para esse **\<Description\>** elemento é definido como a **ID** de um **\<String\>** elemento. O **\<String\>** elemento contém o texto para **\<Description\>**. <br/> |
|**\<Requirements\>** <br/> |Opcional. Especifica o conjunto de requisitos mínimos e a versão do Office.js exigida pelo suplemento. Esse elemento filho **\<Requirements\>** substitui o **\<Requirements\>** elemento na parte pai do manifesto. Para obter mais informações, consulte [Especificar aplicativos do Office e requisitos de API](../develop/specify-office-hosts-and-api-requirements.md).  <br/> |
|**\<Hosts\>** <br/> |Obrigatório. Especifica uma coleção de aplicativos do Office. O elemento filho **\<Hosts\>** substitui o **\<Hosts\>** elemento na parte pai do manifesto. Você deve incluir um conjunto de atributos **xsi:type** como "Pasta de trabalho" ou "Documento". <br/> |
|**\<Resources\>** <br/> |Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) consultado por outros elementos do manifesto. Por exemplo, o **\<Description\>** valor do elemento refere-se a um elemento filho em **\<Resources\>**. O **\<Resources\>** elemento é descrito na [Etapa 7: Adicionar o elemento Resources](#step-7-add-the-resources-element) mais adiante neste artigo. <br/> |

O exemplo a seguir mostra como usar o **\<VersionOverrides\>** elemento e seus elementos filho.

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

O **\<Hosts\>** elemento contém um ou mais **\<Host\>** elementos. Um **\<Host\>** elemento especifica um aplicativo específico do Office. O **\<Host\>** elemento contém elementos filho que especificam os comandos de suplemento a serem exibidos após a instalação do suplemento nesse aplicativo do Office. Para mostrar os mesmos comandos de suplemento em dois ou mais aplicativos diferentes do Office, você deve duplicar os elementos filho em cada **\<Host\>**.

O **\<DesktopFormFactor\>** elemento especifica as configurações de um suplemento que é executado em Office na Web (em um navegador) e no Windows.

A seguir está um exemplo de , **\<Host\>** e **\<DesktopFormFactor\>** elementos **\<Hosts\>**.

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

O **\<FunctionFile\>** elemento especifica um arquivo que contém código JavaScript a ser executado quando um comando de suplemento usa a ação **ExecuteFunction** (consulte [Controles de botão](/javascript/api/manifest/control-button) para obter uma descrição). O **\<FunctionFile\>** atributo **resid do** elemento é definido como um arquivo HTML que inclui todos os arquivos JavaScript que seus comandos de suplemento exigem. Você não pode vincular diretamente a um arquivo JavaScript. Você só pode vincular a um arquivo HTML. O nome do arquivo é especificado como um **\<Url\>** elemento no **\<Resources\>** elemento.

A seguir está um exemplo do **\<FunctionFile\>** elemento.
  
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
> Verifique se seu código JavaScript chama `Office.initialize`.

O JavaScript no arquivo HTML referenciado pelo **\<FunctionFile\>** elemento deve chamar `Office.initialize`. O **\<FunctionName\>** elemento (consulte [Controles de botão](/javascript/api/manifest/control-button) para uma descrição) usa as funções em **\<FunctionFile\>**.

O código a seguir mostra como implementar a função usada por **\<FunctionName\>**.

```html
<script>
    // The initialize function must be run each time a new page is loaded.
    (function () {
        Office.initialize = function (reason) {
            // If you need to initialize something you can do so here.
        };
    })();

    // Define the function.
    function writeText(event) {

        // Implement your custom code here. The following code is a simple example.  
        Office.context.document.setSelectedDataAsync("Function command works. Button ID=" + event.source.id,
            function (asyncResult) {
                const error = asyncResult.error;
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    // Show error message.
                }
                else {
                    // Show success message.
                }
            });

        // Calling event.completed is required. event.completed lets the platform know that processing has completed.
        event.completed();
    }
    
    // You must register the function with the following line.
    Office.actions.associate("writeText", writeText);
</script>
```

> [!IMPORTANT]
> The call to **event.completed** signals that you have successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls **event.completed**, the next queued call to that function runs. You must implement **event.completed**, otherwise your function will not run.

## <a name="step-6-add-extensionpoint-elements"></a>Etapa 6: adicionar elementos do ExtensionPoint

O **\<ExtensionPoint\>** elemento define onde os comandos de suplemento devem aparecer na interface do usuário do Office. Você pode definir **\<ExtensionPoint\>** elementos com esses valores **xsi:type** .

- **PrimaryCommandSurface**, que se refere à faixa de opções no Office.

- **ContextMenu**, que é o menu de atalho exibido quando você clica com o botão direito na interface do usuário do Office.

Os exemplos a seguir mostram como usar o **\<ExtensionPoint\>** elemento com valores de atributo **PrimaryCommandSurface** e **ContextMenu** e os elementos filho que devem ser usados com cada um deles.

> [!IMPORTANT]
> For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format: `<CustomTab id="mycompanyname.mygroupname">`.
  
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

|Elemento|Descrição|
|:-----|:-----|
|**\<CustomTab\>** <br/> |Obrigatório se você quiser adicionar uma guia personalizada à faixa de opções (usando **PrimaryCommandSurface**). Se você usar o **\<CustomTab\>** elemento, não poderá usar o **\<OfficeTab\>** elemento. O atributo **id** é obrigatório. <br/> |
|**\<OfficeTab\>** <br/> |Necessário se você quiser estender uma guia de faixa de opções de aplicativo do Office padrão (usando **PrimaryCommandSurface**). Se você usar o **\<OfficeTab\>** elemento, não poderá usar o **\<CustomTab\>** elemento. <br/> Para obter mais valores de guia a serem usados com o atributo **id** , consulte [Valores de guia para guias de faixa de opções de aplicativo do Office padrão](/javascript/api/manifest/officetab).  <br/> |
|**\<OfficeMenu\>** <br/> | Obrigatório se você estiver adicionando comandos de suplemento a um menu de contexto padrão (usando **ContextMenu**). O atributo **id** deve ser definido como: <br/> **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. <br/> **ContextMenuCell** para Excel. Exibe o item no menu de contexto quando o usuário clica com o botão direito do mouse em uma célula na planilha. <br/> |
|**\<Group\>** <br/> |A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters. <br/> |
|**\<Label\>** <br/> |Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<String\>** elemento. O **\<String\>** elemento é um elemento filho do **\<ShortStrings\>** elemento, que é um elemento filho do **\<Resources\>** elemento. <br/> |
|**\<Icon\>** <br/> |Obrigatório. Especifica o ícone do grupo a ser usado em dispositivos de fator forma pequeno, ou quando muitos botões forem exibidos. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<Image\>** elemento. O **\<Image\>** elemento é um elemento filho do **\<Images\>** elemento, que é um elemento filho do **\<Resources\>** elemento. O atributo **size** fornece o tamanho, em pixels, da imagem. Três tamanhos de imagem são obrigatórios: 16, 32 e 80 pixels. Também há suporte para cinco tamanhos opcionais: 20, 24, 40, 48 e 64 pixels. <br/> |
|**\<Tooltip\>** <br/> |Opcional. A dica de ferramenta do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<String\>** elemento. O **\<String\>** elemento é um elemento filho do **\<LongStrings\>** elemento, que é um elemento filho do **\<Resources\>** elemento. <br/> |
|**\<Control\>** <br/> |Cada grupo exige pelo menos um controle. Um **\<Control\>** elemento pode ser um **Botão** ou um **Menu**. Use **Menu** para especificar uma lista suspensa de controles de botão. Atualmente, há suporte apenas para botões e menus. Consulte [Controles de botão](/javascript/api/manifest/control-button) e [controles de menu](/javascript/api/manifest/control-menu) para obter mais informações. <br/>**Nota:** Para facilitar a solução de problemas, recomendamos que você adicione um **\<Control\>** elemento e os elementos filho relacionados **\<Resources\>** um de cada vez.          |

### <a name="button-controls"></a>Controles de botão

Um botão executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função JavaScript ou a exibição de um painel de tarefas. O exemplo a seguir mostra como definir dois botões. O primeiro botão executa uma função JavaScript sem mostrar uma interface do usuário e o segundo botão mostra um painel de tarefas. **\<Control\>** No elemento:

- O atributo **type** é obrigatório e deve ser definido como **Button**.

- O atributo **id** do **\<Control\>** elemento é uma cadeia de caracteres com no máximo 125 caracteres.

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

|Elementos|Descrição|
|:-----|:-----|
|**\<Label\>** <br/> |Obrigatório. O texto do botão. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<String\>** elemento. O **\<String\>** elemento é um elemento filho do **\<ShortStrings\>** elemento, que é um elemento filho do **\<Resources\>** elemento. <br/> |
|**\<Tooltip\>** <br/> |Opcional. A dica de ferramenta do botão. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<String\>** elemento. O **\<String\>** elemento é um elemento filho do **\<LongStrings\>** elemento, que é um elemento filho do **\<Resources\>** elemento. <br/> |
|**\<Supertip\>** <br/> | Obrigatório. A superdica para esse botão, que é definida pelos seguintes itens: <br/> **Título** <br/>  Obrigatório. O texto da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<String\>** elemento. O **\<String\>** elemento é um elemento filho do **\<ShortStrings\>** elemento, que é um elemento filho do **\<Resources\>** elemento. <br/> **\<Description\>** <br/>  Obrigatório. A descrição da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<String\>** elemento. O **\<String\>** elemento é um elemento filho do **\<LongStrings\>** elemento, que é um elemento filho do **\<Resources\>** elemento. <br/> |
|**\<Icon\>** <br/> | Obrigatório. Contém os **\<Image\>** elementos para o botão. Arquivos de imagem devem estar no formato .png. <br/> **\<Image\>** <br/>  Define uma imagem a ser exibida no botão. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<Image\>** elemento. O **\<Image\>** elemento é um elemento filho do **\<Images\>** elemento, que é um elemento filho do **\<Resources\>** elemento. O atributo **size** indica o tamanho em pixels da imagem. Três tamanhos de imagem são obrigatórios: 16, 32 e 80 pixels. Também há suporte para cinco tamanhos opcionais: 20, 24, 40, 48 e 64 pixels. <br/> |
|**\<Action\>** <br/> | Required. Specifies the action to perform when the user selects the button. You can specify one of the following values for the **xsi:type** attribute: <br/> **ExecuteFunction**, que executa uma função JavaScript localizada no arquivo referenciado por **\<FunctionFile\>**. O **\<FunctionName\>** elemento filho especifica o nome da função a ser executada. <br/> **ShowTaskPane**, que mostra o painel de tarefas do suplemento. O **\<SourceLocation\>** elemento filho especifica o local do arquivo de origem da página a ser exibida. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<Url\>** elemento no **\<Urls\>** elemento no **\<Resources\>** elemento. <br/> |

### <a name="menu-controls"></a>Controles de menu

Um controle **Menu** pode ser usado com **PrimaryCommandSurface** ou **ContextMenu** e define:
  
- Um item de menu no nível raiz.
- Uma lista de itens de submenu.

When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.

O exemplo a seguir mostra como definir um item de menu com dois itens de submenu. O primeiro item do submenu mostra um painel de tarefas e o segundo item executa uma função JavaScript. **\<Control\>** No elemento:

- O atributo **xsi:type** é obrigatório e deve ser definido como **Menu**.
- O atributo **id** é uma cadeia de caracteres com, no máximo, 125 caracteres.

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

|Elementos|Descrição|
|:-----|:-----|
|**\<Label\>** <br/> |Obrigatório. O texto do item de menu raiz. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<String\>** elemento. O **\<String\>** elemento é um elemento filho do **\<ShortStrings\>** elemento, que é um elemento filho do **\<Resources\>** elemento. <br/> |
|**\<Tooltip\>** <br/> |Opcional. A dica de ferramenta do menu. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<String\>** elemento. O **\<String\>** elemento é um elemento filho do **\<LongStrings\>** elemento, que é um elemento filho do **\<Resources\>** elemento. <br/> |
|**\<SuperTip\>** <br/> | Obrigatório. A superdica para o menu, que é definida pelos seguintes itens: <br/> **\<Title\>** <br/>  Obrigatório. O texto da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<String\>** elemento. O **\<String\>** elemento é um elemento filho do **\<ShortStrings\>** elemento, que é um elemento filho do **\<Resources\>** elemento. <br/> **\<Description\>** <br/>  Obrigatório. A descrição da superdica. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<String\>** elemento. O **\<String\>** elemento é um elemento filho do **\<LongStrings\>** elemento, que é um elemento filho do **\<Resources\>** elemento. <br/> |
|**\<Icon\>** <br/> | Obrigatório. Contém os **\<Image\>** elementos do menu. Arquivos de imagem devem estar no formato .png. <br/> **\<Image\>** <br/>  Uma imagem para o menu. O atributo **resid** deve ser definido como o valor do atributo **id** de um **\<Image\>** elemento. O **\<Image\>** elemento é um elemento filho do **\<Images\>** elemento, que é um elemento filho do **\<Resources\>** elemento. O atributo **size** indica o tamanho em pixels da imagem. Três tamanhos de imagem, em pixels, são obrigatórios: 16, 32 e 80 pixels. Cinco tamanhos opcionais, em pixels, também têm suporte: 20, 24, 40, 48 e 64 pixels. <br/> |
|**\<Items\>** <br/> |Obrigatório. Contém os **\<Item\>** elementos para cada item de submenu. Cada **\<Item\>** elemento contém os mesmos elementos filho que [os controles button](/javascript/api/manifest/control-button).  <br/> |

## <a name="step-7-add-the-resources-element"></a>Etapa 7: adicionar o elemento Resources

O **\<Resources\>** elemento contém recursos usados pelos diferentes elementos filho do **\<VersionOverrides\>** elemento. Resources inclui ícones, cadeias de caracteres e URLs. Um elemento no manifesto pode usar um recurso fazendo referência a **id** do recurso. O uso da **id** ajuda a organizar o manifesto, especialmente quando há versões diferentes do recurso para localidades diferentes. Uma **id** tem no máximo 32 caracteres.
  
O seguinte mostra um exemplo de como usar o **\<Resources\>** elemento. Cada recurso pode ter um ou mais **\<Override\>** elementos filho para definir um recurso diferente para uma localidade específica.

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

|Recurso|Descrição|
|:-----|:-----|
|**\<Images\>**/ **\<Image\>** <br/> | Fornece a URL HTTPS para um arquivo de imagem. Cada imagem precisa definir os três tamanhos de imagem necessários: <br/>  16×16 <br/>  32×32 <br/>  80×80 <br/>  Os seguintes tamanhos de imagem também têm suporte, mas não são obrigatórios: <br/>  20×20 <br/>  24×24 <br/>  40×40 <br/>  48×48 <br/>  64×64 <br/> |
|**\<Urls\>**/ **\<Url\>** <br/> |Fornece um local para a URL HTTPS. Uma URL pode ter no máximo 2048 caracteres.  <br/> |
|**\<ShortStrings\>**/ **\<String\>** <br/> |O texto para **\<Label\>** e **\<Title\>** elementos. Cada **\<String\>** um contém um máximo de 125 caracteres. <br/> |
|**\<LongStrings\>**/ **\<String\>** <br/> |O texto para **\<Tooltip\>** e **\<Description\>** elementos. Cada **\<String\>** um contém um máximo de 250 caracteres. <br/> |

> [!NOTE]
> Você deve usar a SSL (Secure Sockets Layer) para todas as URLs nos **\<Image\>** elementos e **\<Url\>** .

### <a name="tab-values-for-default-office-app-ribbon-tabs"></a>Valores de guia para guias padrão de faixa de opções de aplicativo do Office

No Excel e no Word, é possível adicionar seus comandos de suplemento na faixa de opções usando as guias padrão da interface de usuário do Office. A tabela a seguir lista os valores que você pode usar para o atributo **id** do **\<OfficeTab\>** elemento. Os valores da guia diferenciam maiúsculas de minúsculas.

|Aplicativo cliente do Office|Valores de guia|
|:-----|:-----|
|Excel  <br/> |**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval** <br/> |
|Word  <br/> |**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation** <br/> |
|PowerPoint  <br/> |**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**          <br/> |

## <a name="see-also"></a>Confira também

- [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md)
- [Exemplo: criar um suplemento do Excel com botões de comando](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/excel)
- [Exemplo: criar um suplemento do Word com botões de comando](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/word)
- [Exemplo: criar um suplemento do PowerPoint com botões de comando](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-add-in-commands/powerpoint)
