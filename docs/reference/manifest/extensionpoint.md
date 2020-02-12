---
title: Elemento ExtensionPoint no arquivo de manifesto
description: ''
ms.date: 09/05/2019
localization_priority: Normal
ms.openlocfilehash: 2ad9e0ccb0393e562ca0bb6951ec9bc4eb9c3eab
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950702"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="3ec83-102">Elemento ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="3ec83-102">ExtensionPoint element</span></span>

 <span data-ttu-id="3ec83-103">Define onde um suplemento expõe a funcionalidade na interface de usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="3ec83-103">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="3ec83-104">O elemento **ExtensionPoint** é um elemento filho de [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="3ec83-104">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="3ec83-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="3ec83-105">Attributes</span></span>

|  <span data-ttu-id="3ec83-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="3ec83-106">Attribute</span></span>  |  <span data-ttu-id="3ec83-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3ec83-107">Required</span></span>  |  <span data-ttu-id="3ec83-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ec83-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3ec83-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="3ec83-109">**xsi:type**</span></span>  |  <span data-ttu-id="3ec83-110">Sim</span><span class="sxs-lookup"><span data-stu-id="3ec83-110">Yes</span></span>  | <span data-ttu-id="3ec83-111">O tipo de ponto de extensão que está sendo definido.</span><span class="sxs-lookup"><span data-stu-id="3ec83-111">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="3ec83-112">Pontos de extensão somente para Excel</span><span class="sxs-lookup"><span data-stu-id="3ec83-112">Extension points for Excel only</span></span>

- <span data-ttu-id="3ec83-113">**CustomFunctions**: uma função personalizada escrita em JavaScript para Excel.</span><span class="sxs-lookup"><span data-stu-id="3ec83-113">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="3ec83-114">[Este exemplo de código XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) mostra como usar o elemento **ExtensionPoint** com o valor do atributo **CustomFunctions** e os elementos filhos a serem usados.</span><span class="sxs-lookup"><span data-stu-id="3ec83-114">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="3ec83-115">Pontos de extensão para comandos de suplemento do Word, Excel, PowerPoint e OneNote</span><span class="sxs-lookup"><span data-stu-id="3ec83-115">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="3ec83-116">**PrimaryCommandSurface**, que se refere à faixa de opções no Office.</span><span class="sxs-lookup"><span data-stu-id="3ec83-116">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="3ec83-117">**ContextMenu**, que é o menu de atalho exibido ao clicar com o botão direito do mouse na interface de usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="3ec83-117">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="3ec83-118">Os exemplos a seguir mostram como usar o elemento **ExtensionPoint** com os valores de atributo **PrimaryCommandSurface** e **ContextMenu** e os elementos filhos que devem ser usados com cada um.</span><span class="sxs-lookup"><span data-stu-id="3ec83-118">The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="3ec83-119">Forneça uma ID exclusiva para os elementos que contêm um atributo ID.</span><span class="sxs-lookup"><span data-stu-id="3ec83-119">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="3ec83-120">É recomendável usar o nome de sua empresa com a ID.</span><span class="sxs-lookup"><span data-stu-id="3ec83-120">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="3ec83-121">Por exemplo, use o formato a seguir.</span><span class="sxs-lookup"><span data-stu-id="3ec83-121">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

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

#### <a name="child-elements"></a><span data-ttu-id="3ec83-122">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3ec83-122">Child elements</span></span>
 
|<span data-ttu-id="3ec83-123">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="3ec83-123">**Element**</span></span>|<span data-ttu-id="3ec83-124">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="3ec83-124">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="3ec83-125">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="3ec83-125">**CustomTab**</span></span>|<span data-ttu-id="3ec83-p103">Obrigatório se você quiser adicionar uma guia personalizada à faixa de opções (usando **PrimaryCommandSurface**). Se você usar o elemento **CustomTab**, não será possível usar o elemento **OfficeTab**. O atributo **id** é obrigatório.</span><span class="sxs-lookup"><span data-stu-id="3ec83-p103">Required if you want to add a custom tab to the ribbon (using  **PrimaryCommandSurface**). If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="3ec83-129">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="3ec83-129">**OfficeTab**</span></span>|<span data-ttu-id="3ec83-p104">Obrigatório se você quiser estender uma guia padrão da faixa de opções do Office (usando **PrimaryCommandSurface**). Se você usar o elemento **OfficeTab**, não poderá usar o elemento **CustomTab**. Para saber mais, confira [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="3ec83-p104">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="3ec83-133">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="3ec83-133">**OfficeMenu**</span></span>|<span data-ttu-id="3ec83-p105">Obrigatório se você estiver adicionando comandos de suplemento a um menu de contexto padrão (usando **ContextMenu**). O atributo **id** deve ser definido como: </span><span class="sxs-lookup"><span data-stu-id="3ec83-p105">Required if you're adding add-in commands to a default context menu (using  **ContextMenu**). The  **id** attribute must be set to: </span></span><br/> <span data-ttu-id="3ec83-p106">- **ContextMenuText** para o Excel ou Word. Exibe o item no menu de contexto quando o texto for selecionado e o usuário clicar com o botão direito do mouse no texto selecionado. </span><span class="sxs-lookup"><span data-stu-id="3ec83-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="3ec83-p107">- **ContextMenuCell** para Excel. Exibe o item no menu de contexto quando o usuário clica com o botão direito do mouse em uma célula na planilha.</span><span class="sxs-lookup"><span data-stu-id="3ec83-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="3ec83-140">**Group**</span><span class="sxs-lookup"><span data-stu-id="3ec83-140">**Group**</span></span>|<span data-ttu-id="3ec83-p108">Um grupo de pontos de extensão de interface do usuário em uma guia. O grupo pode ter até seis controles. O atributo **id** é obrigatório. É uma cadeia de caracteres com, no máximo, 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="3ec83-p108">A group of user interface extension points on a tab. A group can have up to six controls. The  **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="3ec83-144">**Label**</span><span class="sxs-lookup"><span data-stu-id="3ec83-144">**Label**</span></span>|<span data-ttu-id="3ec83-p109">Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **ShortStrings**, que é elemento filho do elemento **Resources**.</span><span class="sxs-lookup"><span data-stu-id="3ec83-p109">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="3ec83-149">**Icon**</span><span class="sxs-lookup"><span data-stu-id="3ec83-149">**Icon**</span></span>|<span data-ttu-id="3ec83-p110">Obrigatório. Especifica o ícone do grupo a ser usado em dispositivos de fator forma pequeno, ou quando muitos botões forem exibidos. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Image**. O elemento **Image** é elemento filho do elemento **Images**, que é elemento filho do elemento **Resources**. O atributo **size** fornece o tamanho da imagem em pixels. Três tamanhos de imagem são obrigatórios: 16, 32 e 80 pixels. Também há suporte para cinco tamanhos opcionais: 20, 24, 40, 48 e 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="3ec83-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="3ec83-157">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="3ec83-157">**Tooltip**</span></span>|<span data-ttu-id="3ec83-p111">Opcional. A dica de ferramenta do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **LongStrings**, que, por sua vez, é um elemento filho do elemento **Resources**.</span><span class="sxs-lookup"><span data-stu-id="3ec83-p111">Optional. The tooltip of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="3ec83-162">**Control**</span><span class="sxs-lookup"><span data-stu-id="3ec83-162">**Control**</span></span>|<span data-ttu-id="3ec83-163">Cada grupo exige pelo menos um controle.</span><span class="sxs-lookup"><span data-stu-id="3ec83-163">Each group requires at least one control.</span></span> <span data-ttu-id="3ec83-164">Um elemento **Control** pode ser um **Button** ou um **Menu**.</span><span class="sxs-lookup"><span data-stu-id="3ec83-164">A  **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="3ec83-165">Use **Menu** para especificar uma lista suspensa de controles de botão.</span><span class="sxs-lookup"><span data-stu-id="3ec83-165">Use  **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="3ec83-166">Atualmente, há suporte apenas para botões e menus.</span><span class="sxs-lookup"><span data-stu-id="3ec83-166">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="3ec83-167">Confira as seções [Controles de botão](control.md#button-control) e [Controles de menu](control.md#menu-dropdown-button-controls) para saber mais.</span><span class="sxs-lookup"><span data-stu-id="3ec83-167">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="3ec83-168">**Observação:** para facilitar a solução de problemas, é recomendável que um elemento **Control** e os elementos filho **Resources** associados sejam adicionados um de cada vez.</span><span class="sxs-lookup"><span data-stu-id="3ec83-168">**Note:**  To make troubleshooting easier, we recommend that a  **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="3ec83-169">**Script**</span><span class="sxs-lookup"><span data-stu-id="3ec83-169">**Script**</span></span>|<span data-ttu-id="3ec83-170">Links para o arquivo JavaScript com a definição de função personalizada e o código de registro</span><span class="sxs-lookup"><span data-stu-id="3ec83-170">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="3ec83-171">Esse elemento não é usado na Visualização do Desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="3ec83-171">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="3ec83-172">Em vez disso, a página HTML é responsável por carregar todos os arquivos JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3ec83-172">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="3ec83-173">**Page**</span><span class="sxs-lookup"><span data-stu-id="3ec83-173">**Page**</span></span>|<span data-ttu-id="3ec83-174">Links para a página HTML de suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="3ec83-174">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="3ec83-175">Pontos de extensão para Outlook</span><span class="sxs-lookup"><span data-stu-id="3ec83-175">Extension points for Outlook</span></span>

- [<span data-ttu-id="3ec83-176">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3ec83-176">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="3ec83-177">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3ec83-177">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="3ec83-178">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3ec83-178">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="3ec83-179">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3ec83-179">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="3ec83-180">[Module](#module) (Só pode ser usado em [DesktopFormFactor](desktopformfactor.md)).</span><span class="sxs-lookup"><span data-stu-id="3ec83-180">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="3ec83-181">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3ec83-181">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="3ec83-182">Eventos</span><span class="sxs-lookup"><span data-stu-id="3ec83-182">Events</span></span>](#events)
- [<span data-ttu-id="3ec83-183">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="3ec83-183">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="3ec83-184">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3ec83-184">MessageReadCommandSurface</span></span>
<span data-ttu-id="3ec83-p114">Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de email. No Outlook para área de trabalho, isso aparece na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="3ec83-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="3ec83-187">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3ec83-187">Child elements</span></span>

|  <span data-ttu-id="3ec83-188">Elemento</span><span class="sxs-lookup"><span data-stu-id="3ec83-188">Element</span></span> |  <span data-ttu-id="3ec83-189">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ec83-189">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3ec83-190">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-190">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="3ec83-191">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="3ec83-191">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="3ec83-192">CustomTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-192">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="3ec83-193">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="3ec83-193">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="3ec83-194">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-194">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="3ec83-195">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-195">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="3ec83-196">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3ec83-196">MessageComposeCommandSurface</span></span>
<span data-ttu-id="3ec83-197">Este ponto de extensão coloca botões na faixa de opções para suplementos que usam o formulário de composição de email.</span><span class="sxs-lookup"><span data-stu-id="3ec83-197">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="3ec83-198">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3ec83-198">Child elements</span></span>

|  <span data-ttu-id="3ec83-199">Elemento</span><span class="sxs-lookup"><span data-stu-id="3ec83-199">Element</span></span> |  <span data-ttu-id="3ec83-200">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ec83-200">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3ec83-201">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-201">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="3ec83-202">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="3ec83-202">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="3ec83-203">CustomTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-203">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="3ec83-204">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="3ec83-204">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="3ec83-205">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-205">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="3ec83-206">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-206">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="3ec83-207">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3ec83-207">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="3ec83-208">Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="3ec83-208">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="3ec83-209">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3ec83-209">Child elements</span></span>

|  <span data-ttu-id="3ec83-210">Elemento</span><span class="sxs-lookup"><span data-stu-id="3ec83-210">Element</span></span> |  <span data-ttu-id="3ec83-211">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ec83-211">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3ec83-212">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-212">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="3ec83-213">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="3ec83-213">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="3ec83-214">CustomTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-214">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="3ec83-215">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="3ec83-215">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="3ec83-216">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-216">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="3ec83-217">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-217">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="3ec83-218">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3ec83-218">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="3ec83-219">Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao participante da reunião.</span><span class="sxs-lookup"><span data-stu-id="3ec83-219">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="3ec83-220">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3ec83-220">Child elements</span></span>

|  <span data-ttu-id="3ec83-221">Elemento</span><span class="sxs-lookup"><span data-stu-id="3ec83-221">Element</span></span> |  <span data-ttu-id="3ec83-222">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ec83-222">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3ec83-223">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-223">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="3ec83-224">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="3ec83-224">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="3ec83-225">CustomTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-225">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="3ec83-226">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="3ec83-226">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="3ec83-227">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-227">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="3ec83-228">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-228">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="3ec83-229">Module</span><span class="sxs-lookup"><span data-stu-id="3ec83-229">Module</span></span>

<span data-ttu-id="3ec83-230">Este ponto de extensão coloca botões na faixa de opções para a extensão do módulo.</span><span class="sxs-lookup"><span data-stu-id="3ec83-230">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="3ec83-231">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3ec83-231">Child elements</span></span>

|  <span data-ttu-id="3ec83-232">Elemento</span><span class="sxs-lookup"><span data-stu-id="3ec83-232">Element</span></span> |  <span data-ttu-id="3ec83-233">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ec83-233">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3ec83-234">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-234">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="3ec83-235">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="3ec83-235">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="3ec83-236">CustomTab</span><span class="sxs-lookup"><span data-stu-id="3ec83-236">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="3ec83-237">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="3ec83-237">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="3ec83-238">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="3ec83-238">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="3ec83-239">Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de email no fator forma móvel.</span><span class="sxs-lookup"><span data-stu-id="3ec83-239">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="3ec83-240">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3ec83-240">Child elements</span></span>

|  <span data-ttu-id="3ec83-241">Elemento</span><span class="sxs-lookup"><span data-stu-id="3ec83-241">Element</span></span> |  <span data-ttu-id="3ec83-242">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ec83-242">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3ec83-243">Group</span><span class="sxs-lookup"><span data-stu-id="3ec83-243">Group</span></span>](group.md) |  <span data-ttu-id="3ec83-244">Adiciona um grupo de botões à superfície de comando.</span><span class="sxs-lookup"><span data-stu-id="3ec83-244">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="3ec83-245">Os elementos **ExtensionPoint** desse tipo só podem ter um elemento filho: um elemento **Group**.</span><span class="sxs-lookup"><span data-stu-id="3ec83-245">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="3ec83-246">Os elementos **Control** contidos neste ponto de extensão precisam ter o atributo **xsi:type** definido como `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="3ec83-246">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="3ec83-247">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3ec83-247">Example</span></span>
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

### <a name="events"></a><span data-ttu-id="3ec83-248">Eventos</span><span class="sxs-lookup"><span data-stu-id="3ec83-248">Events</span></span>

<span data-ttu-id="3ec83-249">Este ponto de extensão adiciona um manipulador de eventos para um evento especificado.</span><span class="sxs-lookup"><span data-stu-id="3ec83-249">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="3ec83-250">Esse tipo de elemento tem suporte no Outlook clássico na Web e na [visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Windows, Mac e Outlook moderno na Web.</span><span class="sxs-lookup"><span data-stu-id="3ec83-250">This element type is supported by classic Outlook on the web, and in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Windows, Mac, and modern Outlook on the web.</span></span> <span data-ttu-id="3ec83-251">Uma assinatura do Office 365 também é necessária.</span><span class="sxs-lookup"><span data-stu-id="3ec83-251">An Office 365 subscription is also required.</span></span>

| <span data-ttu-id="3ec83-252">Elemento</span><span class="sxs-lookup"><span data-stu-id="3ec83-252">Element</span></span> | <span data-ttu-id="3ec83-253">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ec83-253">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3ec83-254">Event</span><span class="sxs-lookup"><span data-stu-id="3ec83-254">Event</span></span>](event.md) |  <span data-ttu-id="3ec83-255">Especifica o evento e a função de manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="3ec83-255">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="3ec83-256">Exemplo do evento ItemSend</span><span class="sxs-lookup"><span data-stu-id="3ec83-256">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="3ec83-257">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="3ec83-257">DetectedEntity</span></span>

<span data-ttu-id="3ec83-258">Este ponto extensão adiciona uma ativação do suplemento contextual em um tipo de entidade especificada.</span><span class="sxs-lookup"><span data-stu-id="3ec83-258">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="3ec83-259">O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="3ec83-259">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="3ec83-260">Este tipo de elemento está disponível para [ clientes do Outlook que ofereçam suporte a conjuntos de requisitos 1.6 e posteriores](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="3ec83-260">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="3ec83-261">Elemento</span><span class="sxs-lookup"><span data-stu-id="3ec83-261">Element</span></span> |  <span data-ttu-id="3ec83-262">Descrição</span><span class="sxs-lookup"><span data-stu-id="3ec83-262">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="3ec83-263">Label</span><span class="sxs-lookup"><span data-stu-id="3ec83-263">Label</span></span>](#label) |  <span data-ttu-id="3ec83-264">Especifica o rótulo para o suplemento na janela contextual.</span><span class="sxs-lookup"><span data-stu-id="3ec83-264">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="3ec83-265">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="3ec83-265">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="3ec83-266">Especifica a URL para a janela contextual.</span><span class="sxs-lookup"><span data-stu-id="3ec83-266">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="3ec83-267">Rule</span><span class="sxs-lookup"><span data-stu-id="3ec83-267">Rule</span></span>](rule.md) |  <span data-ttu-id="3ec83-268">Especifica a regra ou regras que determinam quando um suplemento é ativado.</span><span class="sxs-lookup"><span data-stu-id="3ec83-268">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="3ec83-269">Label</span><span class="sxs-lookup"><span data-stu-id="3ec83-269">Label</span></span>

<span data-ttu-id="3ec83-p116">Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="3ec83-p116">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="3ec83-273">Requisitos de realce</span><span class="sxs-lookup"><span data-stu-id="3ec83-273">Highlight requirements</span></span>

<span data-ttu-id="3ec83-p117">A única maneira que um usuário pode ativar um suplemento contextual é interagir com uma entidade realçada. Os desenvolvedores podem controlar quais entidades são realçadas usando o atributo `Highlight` do elemento `Rule` para os tipos de regra `ItemHasKnownEntity` e `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="3ec83-p117">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="3ec83-p118">No entanto, há algumas limitações que devem ser consideradas. Essas limitações são para garantir que sempre haverá uma entidade realçada em compromissos ou mensagens aplicáveis para oferecer ao usuário uma maneira de ativar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="3ec83-p118">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="3ec83-278">Os tipos de entidade `EmailAddress` e `Url` não podem ser realçados e, portanto, não podem ser usados para ativar um suplemento.</span><span class="sxs-lookup"><span data-stu-id="3ec83-278">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="3ec83-279">Se for usada uma única regra, `Highlight` DEVERÁ ser definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="3ec83-279">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="3ec83-280">Se usar um tipo de regra `RuleCollection` com `Mode="AND"` para combinar várias regras, pelo menos uma das regras DEVERÁ ter o `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="3ec83-280">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="3ec83-281">Se usar um tipo de regra `RuleCollection` com `Mode="OR"` para combinar várias regras, todas as regras DEVERÃO ter o `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="3ec83-281">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="3ec83-282">Exemplo do evento DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="3ec83-282">DetectedEntity event example</span></span>

```xml
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="residLabelName"/>
  <SourceLocation resid="residDetectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" Highlight="all" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" Highlight="none" />
  </Rule>
</ExtensionPoint> 
```
