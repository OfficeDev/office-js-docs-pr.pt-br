---
title: Elemento ExtensionPoint no arquivo de manifesto
description: Define onde um suplemento expõe a funcionalidade na interface de usuário do Office.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 89040ff84afd7e4c33bea8af04255ef19a8e0f50
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215086"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="4acca-103">Elemento ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="4acca-103">ExtensionPoint element</span></span>

 <span data-ttu-id="4acca-104">Define onde um suplemento expõe a funcionalidade na interface de usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="4acca-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="4acca-105">O elemento **ExtensionPoint** é um elemento filho de [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="4acca-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="4acca-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="4acca-106">Attributes</span></span>

|  <span data-ttu-id="4acca-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="4acca-107">Attribute</span></span>  |  <span data-ttu-id="4acca-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="4acca-108">Required</span></span>  |  <span data-ttu-id="4acca-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="4acca-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="4acca-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="4acca-110">**xsi:type**</span></span>  |  <span data-ttu-id="4acca-111">Sim</span><span class="sxs-lookup"><span data-stu-id="4acca-111">Yes</span></span>  | <span data-ttu-id="4acca-112">O tipo de ponto de extensão que está sendo definido.</span><span class="sxs-lookup"><span data-stu-id="4acca-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="4acca-113">Pontos de extensão somente para Excel</span><span class="sxs-lookup"><span data-stu-id="4acca-113">Extension points for Excel only</span></span>

- <span data-ttu-id="4acca-114">**CustomFunctions**: uma função personalizada escrita em JavaScript para Excel.</span><span class="sxs-lookup"><span data-stu-id="4acca-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="4acca-115">[Este exemplo de código XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) mostra como usar o elemento **ExtensionPoint** com o valor do atributo **CustomFunctions** e os elementos filhos a serem usados.</span><span class="sxs-lookup"><span data-stu-id="4acca-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="4acca-116">Pontos de extensão para comandos de suplemento do Word, Excel, PowerPoint e OneNote</span><span class="sxs-lookup"><span data-stu-id="4acca-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="4acca-117">**PrimaryCommandSurface**, que se refere à faixa de opções no Office.</span><span class="sxs-lookup"><span data-stu-id="4acca-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="4acca-118">**ContextMenu**, que é o menu de atalho exibido ao clicar com o botão direito do mouse na interface de usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="4acca-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="4acca-119">Os exemplos a seguir mostram como usar o elemento **ExtensionPoint** com os valores de atributo **PrimaryCommandSurface** e **ContextMenu** e os elementos filho que devem ser usados com cada um.</span><span class="sxs-lookup"><span data-stu-id="4acca-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4acca-120">Forneça uma ID exclusiva para os elementos que contêm um atributo ID.</span><span class="sxs-lookup"><span data-stu-id="4acca-120">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="4acca-121">É recomendável usar o nome de sua empresa com a ID.</span><span class="sxs-lookup"><span data-stu-id="4acca-121">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="4acca-122">Por exemplo, use o formato a seguir.</span><span class="sxs-lookup"><span data-stu-id="4acca-122">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

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

#### <a name="child-elements"></a><span data-ttu-id="4acca-123">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4acca-123">Child elements</span></span>
 
|<span data-ttu-id="4acca-124">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="4acca-124">**Element**</span></span>|<span data-ttu-id="4acca-125">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="4acca-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="4acca-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="4acca-126">**CustomTab**</span></span>|<span data-ttu-id="4acca-p103">Obrigatório se você quiser adicionar uma guia personalizada à faixa de opções (usando **PrimaryCommandSurface**). Se você usar o elemento **CustomTab**, o elemento **OfficeTab** não poderá ser usado. O atributo **id** é obrigatório. </span><span class="sxs-lookup"><span data-stu-id="4acca-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="4acca-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="4acca-130">**OfficeTab**</span></span>|<span data-ttu-id="4acca-131">Obrigatório se você quiser estender uma guia de faixa de opções padrão do Office (usando **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="4acca-131">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="4acca-132">Se você usar o elemento **OfficeTab**, o elemento **CustomTab** não poderá ser usado.</span><span class="sxs-lookup"><span data-stu-id="4acca-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="4acca-133">Para saber mais, confira [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="4acca-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="4acca-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="4acca-134">**OfficeMenu**</span></span>|<span data-ttu-id="4acca-p105">Obrigatório se você estiver adicionando comandos de suplemento a um menu de contexto padrão (usando **ContextMenu**). O atributo **id** deve ser definido como: </span><span class="sxs-lookup"><span data-stu-id="4acca-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="4acca-p106">- **ContextMenuText** para o Excel ou Word. Exibe o item no menu de contexto quando o texto for selecionado e o usuário clicar com o botão direito do mouse no texto selecionado. </span><span class="sxs-lookup"><span data-stu-id="4acca-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="4acca-p107">- **ContextMenuCell** para Excel. Exibe o item no menu de contexto quando o usuário clica com o botão direito do mouse em uma célula na planilha.</span><span class="sxs-lookup"><span data-stu-id="4acca-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="4acca-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="4acca-141">**Group**</span></span>|<span data-ttu-id="4acca-p108">Um grupo de pontos de extensão de interface do usuário em uma guia. Um grupo pode ter até seis controles. O atributo **id** é obrigatório. É uma cadeia de caracteres com, no máximo, 125 caracteres. </span><span class="sxs-lookup"><span data-stu-id="4acca-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="4acca-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="4acca-145">**Label**</span></span>|<span data-ttu-id="4acca-p109">Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **ShortStrings**, que é elemento filho do elemento **Resources**. </span><span class="sxs-lookup"><span data-stu-id="4acca-p109">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="4acca-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="4acca-150">**Icon**</span></span>|<span data-ttu-id="4acca-p110">Obrigatório. Especifica o ícone do grupo a ser usado em dispositivos de fator forma pequeno ou quando muitos botões são exibidos. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Image**. O elemento **Image** é um elemento filho do elemento **Images**, que é um elemento filho do elemento **Resources**. O atributo **size** fornece o tamanho da imagem em pixels. Três tamanhos de imagem são obrigatórios: 16, 32 e 80 pixels. Também há suporte para cinco tamanhos opcionais: 20, 24, 40, 48 e 64 pixels. </span><span class="sxs-lookup"><span data-stu-id="4acca-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="4acca-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="4acca-158">**Tooltip**</span></span>|<span data-ttu-id="4acca-p111">Opcional. A dica de ferramenta do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **LongStrings**, que é um elemento filho do elemento **Resources**. </span><span class="sxs-lookup"><span data-stu-id="4acca-p111">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="4acca-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="4acca-163">**Control**</span></span>|<span data-ttu-id="4acca-164">Cada grupo exige pelo menos um controle.</span><span class="sxs-lookup"><span data-stu-id="4acca-164">Each group requires at least one control.</span></span> <span data-ttu-id="4acca-165">Um elemento **Control** pode ser um **Button** ou um **Menu**.</span><span class="sxs-lookup"><span data-stu-id="4acca-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="4acca-166">Use **Menu** para especificar uma lista suspensa de controles de botão.</span><span class="sxs-lookup"><span data-stu-id="4acca-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="4acca-167">Atualmente, há suporte apenas para botões e menus.</span><span class="sxs-lookup"><span data-stu-id="4acca-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="4acca-168">Confira as seções [Controles de botão](control.md#button-control) e [Controles de menu](control.md#menu-dropdown-button-controls) para saber mais.</span><span class="sxs-lookup"><span data-stu-id="4acca-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="4acca-169">**Observação:**  Para facilitar a solução de problemas, recomendamos que um elemento **Control** e os elementos filho de **recursos** relacionados sejam adicionados um de cada vez.</span><span class="sxs-lookup"><span data-stu-id="4acca-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="4acca-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="4acca-170">**Script**</span></span>|<span data-ttu-id="4acca-171">Links para o arquivo JavaScript com a definição de função personalizada e o código de registro</span><span class="sxs-lookup"><span data-stu-id="4acca-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="4acca-172">Esse elemento não é usado na Visualização do Desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="4acca-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="4acca-173">Em vez disso, a página HTML é responsável por carregar todos os arquivos JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4acca-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="4acca-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="4acca-174">**Page**</span></span>|<span data-ttu-id="4acca-175">Links para a página HTML de suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="4acca-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="4acca-176">Pontos de extensão para Outlook</span><span class="sxs-lookup"><span data-stu-id="4acca-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="4acca-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4acca-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="4acca-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4acca-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="4acca-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4acca-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="4acca-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4acca-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="4acca-181">[Module](#module) (Só pode ser usado em [DesktopFormFactor](desktopformfactor.md)).</span><span class="sxs-lookup"><span data-stu-id="4acca-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="4acca-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4acca-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="4acca-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4acca-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface-preview)
- [<span data-ttu-id="4acca-184">Eventos</span><span class="sxs-lookup"><span data-stu-id="4acca-184">Events</span></span>](#events)
- [<span data-ttu-id="4acca-185">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="4acca-185">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="4acca-186">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4acca-186">MessageReadCommandSurface</span></span>

<span data-ttu-id="4acca-p114">Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de email. No Outlook para área de trabalho, isso aparece na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="4acca-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="4acca-189">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4acca-189">Child elements</span></span>

|  <span data-ttu-id="4acca-190">Elemento</span><span class="sxs-lookup"><span data-stu-id="4acca-190">Element</span></span> |  <span data-ttu-id="4acca-191">Descrição</span><span class="sxs-lookup"><span data-stu-id="4acca-191">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4acca-192">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4acca-192">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="4acca-193">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="4acca-193">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="4acca-194">CustomTab</span><span class="sxs-lookup"><span data-stu-id="4acca-194">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="4acca-195">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="4acca-195">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="4acca-196">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4acca-196">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="4acca-197">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="4acca-197">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="4acca-198">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4acca-198">MessageComposeCommandSurface</span></span>

<span data-ttu-id="4acca-199">Este ponto de extensão coloca botões na faixa de opções para suplementos que usam o formulário de composição de email.</span><span class="sxs-lookup"><span data-stu-id="4acca-199">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="4acca-200">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4acca-200">Child elements</span></span>

|  <span data-ttu-id="4acca-201">Elemento</span><span class="sxs-lookup"><span data-stu-id="4acca-201">Element</span></span> |  <span data-ttu-id="4acca-202">Descrição</span><span class="sxs-lookup"><span data-stu-id="4acca-202">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4acca-203">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4acca-203">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="4acca-204">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="4acca-204">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="4acca-205">CustomTab</span><span class="sxs-lookup"><span data-stu-id="4acca-205">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="4acca-206">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="4acca-206">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="4acca-207">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4acca-207">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="4acca-208">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="4acca-208">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="4acca-209">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4acca-209">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="4acca-210">Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="4acca-210">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="4acca-211">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4acca-211">Child elements</span></span>

|  <span data-ttu-id="4acca-212">Elemento</span><span class="sxs-lookup"><span data-stu-id="4acca-212">Element</span></span> |  <span data-ttu-id="4acca-213">Descrição</span><span class="sxs-lookup"><span data-stu-id="4acca-213">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4acca-214">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4acca-214">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="4acca-215">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="4acca-215">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="4acca-216">CustomTab</span><span class="sxs-lookup"><span data-stu-id="4acca-216">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="4acca-217">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="4acca-217">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="4acca-218">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4acca-218">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="4acca-219">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="4acca-219">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="4acca-220">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4acca-220">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="4acca-221">Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao participante da reunião.</span><span class="sxs-lookup"><span data-stu-id="4acca-221">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="4acca-222">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4acca-222">Child elements</span></span>

|  <span data-ttu-id="4acca-223">Elemento</span><span class="sxs-lookup"><span data-stu-id="4acca-223">Element</span></span> |  <span data-ttu-id="4acca-224">Descrição</span><span class="sxs-lookup"><span data-stu-id="4acca-224">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4acca-225">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4acca-225">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="4acca-226">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="4acca-226">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="4acca-227">CustomTab</span><span class="sxs-lookup"><span data-stu-id="4acca-227">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="4acca-228">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="4acca-228">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="4acca-229">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4acca-229">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="4acca-230">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="4acca-230">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="4acca-231">Module</span><span class="sxs-lookup"><span data-stu-id="4acca-231">Module</span></span>

<span data-ttu-id="4acca-232">Este ponto de extensão coloca botões na faixa de opções para a extensão do módulo.</span><span class="sxs-lookup"><span data-stu-id="4acca-232">This extension point puts buttons on the ribbon for the module extension.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="4acca-233">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4acca-233">Child elements</span></span>

|  <span data-ttu-id="4acca-234">Elemento</span><span class="sxs-lookup"><span data-stu-id="4acca-234">Element</span></span> |  <span data-ttu-id="4acca-235">Descrição</span><span class="sxs-lookup"><span data-stu-id="4acca-235">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4acca-236">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4acca-236">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="4acca-237">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="4acca-237">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="4acca-238">CustomTab</span><span class="sxs-lookup"><span data-stu-id="4acca-238">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="4acca-239">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="4acca-239">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="4acca-240">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="4acca-240">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="4acca-241">Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de email no fator forma móvel.</span><span class="sxs-lookup"><span data-stu-id="4acca-241">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="4acca-242">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4acca-242">Child elements</span></span>

|  <span data-ttu-id="4acca-243">Elemento</span><span class="sxs-lookup"><span data-stu-id="4acca-243">Element</span></span> |  <span data-ttu-id="4acca-244">Descrição</span><span class="sxs-lookup"><span data-stu-id="4acca-244">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4acca-245">Group</span><span class="sxs-lookup"><span data-stu-id="4acca-245">Group</span></span>](group.md) |  <span data-ttu-id="4acca-246">Adiciona um grupo de botões à superfície de comando.</span><span class="sxs-lookup"><span data-stu-id="4acca-246">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="4acca-247">Os elementos **ExtensionPoint** desse tipo só podem ter um elemento filho: um elemento **Group**.</span><span class="sxs-lookup"><span data-stu-id="4acca-247">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="4acca-248">Os elementos **Control** contidos neste ponto de extensão precisam ter o atributo **xsi:type** definido como `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="4acca-248">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="4acca-249">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4acca-249">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface-preview"></a><span data-ttu-id="4acca-250">MobileOnlineMeetingCommandSurface (visualização)</span><span class="sxs-lookup"><span data-stu-id="4acca-250">MobileOnlineMeetingCommandSurface (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="4acca-251">Este ponto de extensão só tem suporte na [Visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Android com uma assinatura do Office 365.</span><span class="sxs-lookup"><span data-stu-id="4acca-251">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with an Office 365 subscription.</span></span>

<span data-ttu-id="4acca-252">Este ponto de extensão coloca uma alternância apropriada de modo na superfície de comando para um compromisso no fator de forma móvel.</span><span class="sxs-lookup"><span data-stu-id="4acca-252">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="4acca-253">Um organizador da reunião pode criar uma reunião online.</span><span class="sxs-lookup"><span data-stu-id="4acca-253">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="4acca-254">Um participante pode ingressar na reunião online subsequentemente.</span><span class="sxs-lookup"><span data-stu-id="4acca-254">An attendee can subsequently join the online meeting.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="4acca-255">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4acca-255">Child elements</span></span>

|  <span data-ttu-id="4acca-256">Elemento</span><span class="sxs-lookup"><span data-stu-id="4acca-256">Element</span></span> |  <span data-ttu-id="4acca-257">Descrição</span><span class="sxs-lookup"><span data-stu-id="4acca-257">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4acca-258">Control</span><span class="sxs-lookup"><span data-stu-id="4acca-258">Control</span></span>](control.md) |  <span data-ttu-id="4acca-259">Adiciona um botão à superfície de comando.</span><span class="sxs-lookup"><span data-stu-id="4acca-259">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="4acca-260">Os elementos **ExtensionPoint** desse tipo só podem ter um elemento filho: um elemento **Control** .</span><span class="sxs-lookup"><span data-stu-id="4acca-260">**ExtensionPoint** elements of this type can only have one child element: a **Control** element.</span></span>

<span data-ttu-id="4acca-261">O elemento de **controle** contido neste ponto de extensão deve ter o atributo **xsi: Type** definido `MobileButton`como.</span><span class="sxs-lookup"><span data-stu-id="4acca-261">The **Control** element contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="4acca-262">Exemplo</span><span class="sxs-lookup"><span data-stu-id="4acca-262">Example</span></span>

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
      <FunctionName>UiLessInsertMeetingLink</FunctionName>
    </Action>
  </Control>
</ExtensionPoint>
```

### <a name="events"></a><span data-ttu-id="4acca-263">Eventos</span><span class="sxs-lookup"><span data-stu-id="4acca-263">Events</span></span>

<span data-ttu-id="4acca-264">Este ponto de extensão adiciona um manipulador de eventos para um evento especificado.</span><span class="sxs-lookup"><span data-stu-id="4acca-264">This extension point adds an event handler for a specified event.</span></span>

| <span data-ttu-id="4acca-265">Elemento</span><span class="sxs-lookup"><span data-stu-id="4acca-265">Element</span></span> | <span data-ttu-id="4acca-266">Descrição</span><span class="sxs-lookup"><span data-stu-id="4acca-266">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4acca-267">Event</span><span class="sxs-lookup"><span data-stu-id="4acca-267">Event</span></span>](event.md) |  <span data-ttu-id="4acca-268">Especifica o evento e a função de manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="4acca-268">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="4acca-269">Exemplo do evento ItemSend</span><span class="sxs-lookup"><span data-stu-id="4acca-269">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="4acca-270">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="4acca-270">DetectedEntity</span></span>

<span data-ttu-id="4acca-271">Este ponto extensão adiciona uma ativação do suplemento contextual em um tipo de entidade especificada.</span><span class="sxs-lookup"><span data-stu-id="4acca-271">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="4acca-272">O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="4acca-272">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="4acca-273">Este tipo de elemento está disponível para [ clientes do Outlook que ofereçam suporte a conjuntos de requisitos 1.6 e posteriores](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="4acca-273">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="4acca-274">Elemento</span><span class="sxs-lookup"><span data-stu-id="4acca-274">Element</span></span> |  <span data-ttu-id="4acca-275">Descrição</span><span class="sxs-lookup"><span data-stu-id="4acca-275">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="4acca-276">Label</span><span class="sxs-lookup"><span data-stu-id="4acca-276">Label</span></span>](#label) |  <span data-ttu-id="4acca-277">Especifica o rótulo para o suplemento na janela contextual.</span><span class="sxs-lookup"><span data-stu-id="4acca-277">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="4acca-278">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="4acca-278">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="4acca-279">Especifica a URL para a janela contextual.</span><span class="sxs-lookup"><span data-stu-id="4acca-279">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="4acca-280">Rule</span><span class="sxs-lookup"><span data-stu-id="4acca-280">Rule</span></span>](rule.md) |  <span data-ttu-id="4acca-281">Especifica a regra ou regras que determinam quando um suplemento é ativado.</span><span class="sxs-lookup"><span data-stu-id="4acca-281">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="4acca-282">Label</span><span class="sxs-lookup"><span data-stu-id="4acca-282">Label</span></span>

<span data-ttu-id="4acca-283">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="4acca-283">Required.</span></span> <span data-ttu-id="4acca-284">O rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="4acca-284">The label of the group.</span></span> <span data-ttu-id="4acca-285">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="4acca-285">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="4acca-286">Requisitos de realce</span><span class="sxs-lookup"><span data-stu-id="4acca-286">Highlight requirements</span></span>

<span data-ttu-id="4acca-p117">A única maneira que um usuário pode ativar um suplemento contextual é interagir com uma entidade realçada. Os desenvolvedores podem controlar quais entidades são realçadas usando o atributo `Highlight` do elemento `Rule` para os tipos de regra `ItemHasKnownEntity` e `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="4acca-p117">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="4acca-p118">No entanto, há algumas limitações que devem ser consideradas. Essas limitações são para garantir que sempre haverá uma entidade realçada em compromissos ou mensagens aplicáveis para oferecer ao usuário uma maneira de ativar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="4acca-p118">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="4acca-291">Os tipos de entidade `EmailAddress` e `Url` não podem ser realçados e, portanto, não podem ser usados para ativar um suplemento.</span><span class="sxs-lookup"><span data-stu-id="4acca-291">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="4acca-292">Se for usada uma única regra, `Highlight` DEVERÁ ser definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="4acca-292">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="4acca-293">Se usar um tipo de regra `RuleCollection` com `Mode="AND"` para combinar várias regras, pelo menos uma das regras DEVERÁ ter o `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="4acca-293">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="4acca-294">Se usar um tipo de regra `RuleCollection` com `Mode="OR"` para combinar várias regras, todas as regras DEVERÃO ter o `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="4acca-294">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="4acca-295">Exemplo do evento DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="4acca-295">DetectedEntity event example</span></span>

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
