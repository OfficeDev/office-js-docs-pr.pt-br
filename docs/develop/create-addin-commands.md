---
title: Criar comandos de suplemento em seu manifesto para Excel, PowerPoint e Word
description: Use VersionOverrides in your manifest to define add-in commands for Excel, PowerPoint, and Word. Use add-in commands to create UI elements, add buttons or lists, and perform actions.
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 3bcd3c6e07cdb9899601403e68e80e8d609d2e6e
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093711"
---
# <a name="create-add-in-commands-in-your-manifest-for-excel-powerpoint-and-word"></a>Criar comandos de suplemento em seu manifesto para Excel, PowerPoint e Word

Use **[VersionOverrides](../reference/manifest/versionoverrides.md)** in your manifest to define add-in commands for Excel, PowerPoint, and Word. Add-in commands provide an easy way to customize the default Office user interface (UI) with specified UI elements that perform actions. You can use add-in commands to:

- Criar elementos de interface do usuário ou pontos de entrada que facilitam o uso da funcionalidade dos suplementos.
- Adicionar botões ou uma lista suspensa de botões à faixa de opções.
- Adicionar itens de menu individuais — cada um contendo submenus opcionais — aos menus de contexto específicos (atalho).
- Perform actions when your add-in command is chosen. You can:
  - Show one or more task pane add-ins for users to interact with. Inside your task pane add-in, you can display HTML that uses Office UI Fabric to create a custom UI.

     *ou*

  - Executar código JavaScript, que normalmente é executado sem exibir qualquer interface do usuário.

This article describes how to edit your manifest to define add-in commands. The following diagram shows the hierarchy of elements used to define add-in commands. These elements are described in more detail in this article.

> [!NOTE]
> Os comandos de suplemento também são compatíveis com o Outlook. Para obter mais informações, consulte [comandos de suplemento para o Outlook](../outlook/add-in-commands-for-outlook.md)

The following image is an overview of add-in commands elements in the manifest.
![Overview of add-in commands elements in the manifest](../images/version-overrides.png)

## <a name="step-1-start-from-a-sample"></a>Etapa 1: iniciar usando uma amostra

We strongly recommend that you start from one of the samples we provide in  [Office Add-in Commands Samples](https://github.com/OfficeDev/Office-Add-in-Command-Sample). Optionally, you can create your own manifest by following the steps in this guide. You can validate your manifest using the XSD file in the Office Add-in Commands Samples site. Ensure that you have read  [Add-in commands for Excel, Word and PowerPoint](../design/add-in-commands.md) before using add-in commands.

## <a name="step-2-create-a-task-pane-add-in"></a>Etapa 2: criar um suplemento de painel de tarefas

To start using add-in commands, you must first create a task pane add-in, and then modify the add-in's manifest as described in this article. You can't use add-in commands with content add-ins. If you're updating an existing manifest, you must add the appropiate **XML namespaces** as well as add the **VersionOverrides** element to the manifest as described in [Step 3: Add VersionOverrides element](#step-3-add-versionoverrides-element).

The following example shows an Office 2013 add-in's manifest. There are no add-in commands in this manifest because there is no **VersionOverrides** element. Office 2013 doesn't support add-in commands, but by adding **VersionOverrides** to this manifest, your add-in will run in both Office 2013 and Office 2016. In Office 2013, your add-in won't display add-in commands, and uses the value of **SourceLocation** to run your add-in as a single task pane add-in. In Office 2016, if no **VersionOverrides** element is included, **SourceLocation** is used to run your add-in. If you include **VersionOverrides**, however, your add-in displays the add-in commands only, and doesn't display your add-in as a single task pane add-in.
  
```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>657a32a9-ab8a-4579-ac9f-df1a11a64e52</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Contoso Add-in Commands" />
  <Description DefaultValue="Contoso Add-in Commands"/>
  <IconUrl DefaultValue="~remoteAppUrl/Images/Icon_32.png" />
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
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

The **VersionOverrides** element is the root element that contains the definition of your add-in command. **VersionOverrides** is a child element of the **OfficeApp** element in the manifest. The following table lists the attributes of the **VersionOverrides** element.

|**Atributo**|**Descrição**|
|:-----|:-----|
|**xmlns** <br/> | Obrigatório. O local do esquema, que deve ser `http://schemas.microsoft.com/office/taskpaneappversionoverrides`. <br/> |
|**xsi:type** <br/> |Required. The schema version. The version described in this article is "VersionOverridesV1_0".  <br/> |

A tabela a seguir identifica os elementos filho de **VersionOverrides**.
  
|**Elemento**|**Descrição**|
|:-----|:-----|
|**Descrição** <br/> |Optional. Describes the add-in. This child **Description** element overrides a previous **Description** element in the parent portion of the manifest. The **resid** attribute for this **Description** element is set to the **id** of a **String** element. The **String** element contains the text for **Description**. <br/> |
|**Requisitos** <br/> |Optional. Specifies the minimum requirement set and version of Office.js that the add-in requires. This child **Requirements** element overrides the **Requirements** element in the parent portion of the manifest. For more information, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md).  <br/> |
|**Hosts** <br/> |Required. Specifies a collection of Office hosts. The child **Hosts** element overrides the **Hosts** element in the parent portion of the manifest. You must include a **xsi:type** attribute set to "Workbook" or "Document". <br/> |
|**Resources** <br/> |Defines a collection of resources (strings, URLs, and images) that other manifest elements reference. For example, the **Description** element's value refers to a child element in **Resources**. The **Resources** element is described in [Step 7: Add the Resources element](#step-7-add-the-resources-element) later in this article. <br/> |

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

The **Hosts** element contains one or more **Host** elements. A **Host** element specifies a particular Office host. The **Host** element contains child elements that specify the add-in commands to display after your add-in is installed in that Office host. To show the same add-in commands in two or more different Office hosts, you must duplicate the child elements in each **Host**.

O elemento **DesktopFormFactor** especifica as configurações para um suplemento que é executado no Office Online (em um navegador) e no Windows.

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

The **FunctionFile** element specifies a file that contains JavaScript code to run when an add-in command uses the **ExecuteFunction** action (see [Button controls](../reference/manifest/control.md#button-control) for a description). The **FunctionFile** element's **resid** attribute is set to a HTML file that includes all the JavaScript files your add-in commands require. You can't link directly to a JavaScript file. You can only link to an HTML file. The file name is specified as a **Url** element in the **Resources** element.

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
> Verifique se seu código JavaScript chama `Office.initialize`.

The JavaScript in the HTML file referenced by the **FunctionFile** element must call `Office.initialize`. The **FunctionName** element (see [Button controls](../reference/manifest/control.md#button-control) for a description) uses the functions in **FunctionFile**.

O código a seguir mostra como implementar a função usada por **FunctionName**.

```js
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
</script>
```

> [!IMPORTANT]
> The call to **event.completed** signals that you have successfully handled the event. When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued. The first event runs automatically, while the other events remain on the queue. When your function calls **event.completed**, the next queued call to that function runs. You must implement **event.completed**, otherwise your function will not run.

## <a name="step-6-add-extensionpoint-elements"></a>Etapa 6: adicionar elementos do ExtensionPoint

The **ExtensionPoint** element defines where add-in commands should appear in the Office UI. You can define **ExtensionPoint** elements with these **xsi:type** values:

- **PrimaryCommandSurface**, que se refere à faixa de opções no Office.

- **ContextMenu**, que é o menu de atalho exibido quando você clica com o botão direito na interface do usuário do Office.

Os exemplos a seguir mostram como usar o elemento **ExtensionPoint** com os valores de atributo **PrimaryCommandSurface** e **ContextMenu** e os elementos filho que devem ser usados com cada um.

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

|**Elemento**|**Descrição**|
|:-----|:-----|
|**CustomTab** <br/> |Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required. <br/> |
|**OfficeTab** <br/> |Obrigatório se você deseja estender uma guia padrão da faixa de opções do aplicativo do Office (usando **PrimaryCommandSurface**). Se você usar o elemento **OfficeTab**, o elemento **CustomTab** não poderá ser usado. <br/> Para obter mais valores de tabulação a serem usados com o atributo **ID** , consulte [Tab Values for the Office app default Ribbon Tabs](../reference/manifest/officetab.md).  <br/> |
|**OfficeMenu** <br/> | Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: <br/> **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. <br/> **ContextMenuCell** for Excel. Displays the item on the context menu when the user right-clicks on a cell on the spreadsheet. <br/> |
|**Group** <br/> |A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters. <br/> |
|**Label** <br/> |Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. <br/> |
|**Icon** <br/> |Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. <br/> |
|**Tooltip** <br/> |Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. <br/> |
|**Control** <br/> |Each group requires at least one control. A **Control** element can be either a **Button** or a **Menu**. Use **Menu** to specify a drop-down list of button controls. Currently, only buttons and menus are supported. See the  [Button controls](../reference/manifest/control.md#button-control) and [Menu controls](../reference/manifest/control.md#menu-dropdown-button-controls) sections for more information. <br/>**Observação:** para facilitar a solução de problemas, recomendamos adicionar um elemento **Control** e os elementos filho **Resources** relacionados, um de cada vez.          |

### <a name="button-controls"></a>Controles de botão

A button performs a single action when the user selects it. It can either execute a JavaScript function or show a task pane. The following example shows how to define two buttons. The first button runs a JavaScript function without showing a UI, and the second button shows a task pane. In the **Control** element:

- O atributo **type** é obrigatório e deve ser definido como **Button**.

- O atributo **id** do elemento **Control** é uma cadeia de caracteres com, no máximo, 125 caracteres.

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

|**Elementos**|**Descrição**|
|:-----|:-----|
|**Label** <br/> |Required. The text for the button. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. <br/> |
|**Tooltip** <br/> |Optional. The tooltip for the button. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. <br/> |
|**Supertip** <br/> | Required. The supertip for this button, which is defined by the following: <br/> **Título** <br/>  Required. The text for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. <br/> **Descrição** <br/>  Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. <br/> |
|**Icon** <br/> | Required. Contains the **Image** elements for the button. Image files must be .png format. <br/> **Imagem** <br/>  Defines an image to display on the button. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute indicates the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64. <br/> |
|**Action** <br/> | Required. Specifies the action to perform when the user selects the button. You can specify one of the following values for the **xsi:type** attribute: <br/> **ExecuteFunction**, which runs a JavaScript function located in the file referenced by **FunctionFile**. **ExecuteFunction** does not display a UI. The **FunctionName** child element specifies the name of the function to execute. <br/> **ShowTaskPane**, which shows a task pane add-in. The **SourceLocation** child element specifies the source file location of the task pane add-in to display. The **resid** attribute must be set to the value of the **id** attribute of a **Url** element in the **Urls** element in the **Resources** element. <br/> |

### <a name="menu-controls"></a>Controles de menu

Um controle **Menu** pode ser usado com **PrimaryCommandSurface** ou **ContextMenu** e define:
  
- Um item de menu no nível raiz.
- Uma lista de itens de submenu.

When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.

The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function. In the **Control** element:

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

|**Elementos**|**Descrição**|
|:-----|:-----|
|**Label** <br/> |Required. The text of the root menu item. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. <br/> |
|**Tooltip** <br/> |Optional. The tooltip for the menu. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. <br/> |
|**SuperTip** <br/> | Required. The supertip for the menu, which is defined by the following: <br/> **Título** <br/>  Required. The text of the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element. <br/> **Descrição** <br/>  Required. The description for the supertip. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element. <br/> |
|**Icon** <br/> | Required. Contains the **Image** elements for the menu. Image files must be .png format. <br/> **Image** <br/>  An image for the menu. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute indicates the size in pixels of the image. Three image sizes, in pixels, are required: 16, 32, and 80. Five optional sizes, in pixels, are also supported: 20, 24, 40, 48, and 64. <br/> |
|**Items** <br/> |Required. Contains the **Item** elements for each submenu item. Each **Item** element contains the same child elements as [Button controls](../reference/manifest/control.md#button-control).  <br/> |

## <a name="step-7-add-the-resources-element"></a>Etapa 7: adicionar o elemento Resources

The **Resources** element contains resources used by the different child elements of the **VersionOverrides** element. Resources include icons, strings, and URLs. An element in the manifest can use a resource by referencing the **id** of the resource. Using the **id** helps organize the manifest, especially when there are different versions of the resource for different locales. An **id** has a maximum of 32 characters.
  
The following shows an example of how to use the **Resources** element. Each resource can have one or more **Override** child elements to define a different resource for a specific locale.

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

|**Recurso**|**Descrição**|
|:-----|:-----|
|**Images**/ **Image** <br/> | Provides the HTTPS URL to an image file. Each image must define the three required image sizes: <br/>  16×16 <br/>  32×32 <br/>  80×80 <br/>  Os seguintes tamanhos de imagem também têm suporte, mas não são obrigatórios: <br/>  20×20 <br/>  24×24 <br/>  40×40 <br/>  48×48 <br/>  64×64 <br/> |
|**Urls**/ **Url** <br/> |Provides an HTTPS URL location. A URL can be a maximum of 2048 characters.  <br/> |
|**ShortStrings**/ **String** <br/> |The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters. <br/> |
|**LongStrings**/ **String** <br/> |The text for **Tooltip** and **Description** elements. Each **String** contains a maximum of 250 characters. <br/> |

> [!NOTE]
> Use o protocolo SSL (Secure Sockets Layer) para todas as URLs nos elementos **Image** e **Url**.

### <a name="tab-values-for-default-office-app-ribbon-tabs"></a>Valores de guia para as guias da faixa de opções do aplicativo do Office padrão

In Excel and Word, you can add your add-in commands to the ribbon by using the default Office UI tabs. The following table lists the values that you can use for the **id** attribute of the **OfficeTab** element. The tab values are case sensitive.

|**Aplicativo host do Office**|**Valores de guia**|
|:-----|:-----|
|Excel  <br/> |**TabHome**         **TabInsert**         **TabPageLayoutExcel**         **TabFormulas**         **TabData**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabBackgroundRemoval** <br/> |
|Word  <br/> |**TabHome**         **TabInsert**         **TabWordDesign**         **TabPageLayoutWord**         **TabReferences**         **TabMailings**         **TabReviewWord**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabBlogPost**         **TabBlogInsert**         **TabPrintPreview**         **TabOutlining**         **TabConflicts**         **TabBackgroundRemoval**         **TabBroadcastPresentation** <br/> |
|PowerPoint  <br/> |**TabHome**         **TabInsert**         **TabDesign**         **TabTransitions**         **TabAnimations**         **TabSlideShow**         **TabReview**         **TabView**         **TabDeveloper**         **TabAddIns**         **TabPrintPreview**         **TabMerge**         **TabGrayscale**         **TabBlackAndWhite**         **TabBroadcastPresentation**         **TabSlideMaster**         **TabHandoutMaster**         **TabNotesMaster**         **TabBackgroundRemoval**         **TabSlideMasterHome**          <br/> |

## <a name="see-also"></a>Confira também

- [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md)
