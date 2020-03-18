---
title: Elemento ExtensionPoint no arquivo de manifesto
description: Define onde um suplemento expõe a funcionalidade na interface de usuário do Office.
ms.date: 09/05/2019
localization_priority: Normal
ms.openlocfilehash: c945875140fdbdb7ba6aaeed7bb0a7bf5d06e050
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720565"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="76fde-103">Elemento ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="76fde-103">ExtensionPoint element</span></span>

 <span data-ttu-id="76fde-104">Define onde um suplemento expõe a funcionalidade na interface de usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="76fde-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="76fde-105">O elemento **ExtensionPoint** é um elemento filho de [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="76fde-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="76fde-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="76fde-106">Attributes</span></span>

|  <span data-ttu-id="76fde-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="76fde-107">Attribute</span></span>  |  <span data-ttu-id="76fde-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="76fde-108">Required</span></span>  |  <span data-ttu-id="76fde-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="76fde-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="76fde-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="76fde-110">**xsi:type**</span></span>  |  <span data-ttu-id="76fde-111">Sim</span><span class="sxs-lookup"><span data-stu-id="76fde-111">Yes</span></span>  | <span data-ttu-id="76fde-112">O tipo de ponto de extensão que está sendo definido.</span><span class="sxs-lookup"><span data-stu-id="76fde-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="76fde-113">Pontos de extensão somente para Excel</span><span class="sxs-lookup"><span data-stu-id="76fde-113">Extension points for Excel only</span></span>

- <span data-ttu-id="76fde-114">**CustomFunctions**: uma função personalizada escrita em JavaScript para Excel.</span><span class="sxs-lookup"><span data-stu-id="76fde-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="76fde-115">[Este exemplo de código XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) mostra como usar o elemento **ExtensionPoint** com o valor do atributo **CustomFunctions** e os elementos filhos a serem usados.</span><span class="sxs-lookup"><span data-stu-id="76fde-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="76fde-116">Pontos de extensão para comandos de suplemento do Word, Excel, PowerPoint e OneNote</span><span class="sxs-lookup"><span data-stu-id="76fde-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="76fde-117">**PrimaryCommandSurface**, que se refere à faixa de opções no Office.</span><span class="sxs-lookup"><span data-stu-id="76fde-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="76fde-118">**ContextMenu**, que é o menu de atalho exibido ao clicar com o botão direito do mouse na interface de usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="76fde-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="76fde-119">Os exemplos a seguir mostram como usar o elemento **ExtensionPoint** com os valores de atributo **PrimaryCommandSurface** e **ContextMenu** e os elementos filho que devem ser usados com cada um.</span><span class="sxs-lookup"><span data-stu-id="76fde-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="76fde-120">Forneça uma ID exclusiva para os elementos que contêm um atributo ID.</span><span class="sxs-lookup"><span data-stu-id="76fde-120">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="76fde-121">É recomendável usar o nome de sua empresa com a ID.</span><span class="sxs-lookup"><span data-stu-id="76fde-121">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="76fde-122">Por exemplo, use o formato a seguir.</span><span class="sxs-lookup"><span data-stu-id="76fde-122">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

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

#### <a name="child-elements"></a><span data-ttu-id="76fde-123">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="76fde-123">Child elements</span></span>
 
|<span data-ttu-id="76fde-124">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="76fde-124">**Element**</span></span>|<span data-ttu-id="76fde-125">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="76fde-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="76fde-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="76fde-126">**CustomTab**</span></span>|<span data-ttu-id="76fde-p103">Obrigatório se você quiser adicionar uma guia personalizada à faixa de opções (usando **PrimaryCommandSurface**). Se você usar o elemento **CustomTab**, o elemento **OfficeTab** não poderá ser usado. O atributo **id** é obrigatório. </span><span class="sxs-lookup"><span data-stu-id="76fde-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="76fde-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="76fde-130">**OfficeTab**</span></span>|<span data-ttu-id="76fde-131">Obrigatório se você quiser estender uma guia de faixa de opções padrão do Office (usando **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="76fde-131">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="76fde-132">Se você usar o elemento **OfficeTab**, o elemento **CustomTab** não poderá ser usado.</span><span class="sxs-lookup"><span data-stu-id="76fde-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="76fde-133">Para saber mais, confira [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="76fde-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="76fde-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="76fde-134">**OfficeMenu**</span></span>|<span data-ttu-id="76fde-p105">Obrigatório se você estiver adicionando comandos de suplemento a um menu de contexto padrão (usando **ContextMenu**). O atributo **id** deve ser definido como: </span><span class="sxs-lookup"><span data-stu-id="76fde-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="76fde-p106">- **ContextMenuText** para o Excel ou Word. Exibe o item no menu de contexto quando o texto for selecionado e o usuário clicar com o botão direito do mouse no texto selecionado. </span><span class="sxs-lookup"><span data-stu-id="76fde-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="76fde-p107">- **ContextMenuCell** para Excel. Exibe o item no menu de contexto quando o usuário clica com o botão direito do mouse em uma célula na planilha.</span><span class="sxs-lookup"><span data-stu-id="76fde-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="76fde-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="76fde-141">**Group**</span></span>|<span data-ttu-id="76fde-p108">Um grupo de pontos de extensão de interface do usuário em uma guia. Um grupo pode ter até seis controles. O atributo **id** é obrigatório. É uma cadeia de caracteres com, no máximo, 125 caracteres. </span><span class="sxs-lookup"><span data-stu-id="76fde-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="76fde-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="76fde-145">**Label**</span></span>|<span data-ttu-id="76fde-p109">Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **ShortStrings**, que é elemento filho do elemento **Resources**. </span><span class="sxs-lookup"><span data-stu-id="76fde-p109">Required. The label of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="76fde-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="76fde-150">**Icon**</span></span>|<span data-ttu-id="76fde-p110">Obrigatório. Especifica o ícone do grupo a ser usado em dispositivos de fator forma pequeno ou quando muitos botões são exibidos. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Image**. O elemento **Image** é um elemento filho do elemento **Images**, que é um elemento filho do elemento **Resources**. O atributo **size** fornece o tamanho da imagem em pixels. Três tamanhos de imagem são obrigatórios: 16, 32 e 80 pixels. Também há suporte para cinco tamanhos opcionais: 20, 24, 40, 48 e 64 pixels. </span><span class="sxs-lookup"><span data-stu-id="76fde-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="76fde-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="76fde-158">**Tooltip**</span></span>|<span data-ttu-id="76fde-p111">Opcional. A dica de ferramenta do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **LongStrings**, que é um elemento filho do elemento **Resources**. </span><span class="sxs-lookup"><span data-stu-id="76fde-p111">Optional. The tooltip of the group. The **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="76fde-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="76fde-163">**Control**</span></span>|<span data-ttu-id="76fde-164">Cada grupo exige pelo menos um controle.</span><span class="sxs-lookup"><span data-stu-id="76fde-164">Each group requires at least one control.</span></span> <span data-ttu-id="76fde-165">Um elemento **Control** pode ser um **Button** ou um **Menu**.</span><span class="sxs-lookup"><span data-stu-id="76fde-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="76fde-166">Use **Menu** para especificar uma lista suspensa de controles de botão.</span><span class="sxs-lookup"><span data-stu-id="76fde-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="76fde-167">Atualmente, há suporte apenas para botões e menus.</span><span class="sxs-lookup"><span data-stu-id="76fde-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="76fde-168">Confira as seções [Controles de botão](control.md#button-control) e [Controles de menu](control.md#menu-dropdown-button-controls) para saber mais.</span><span class="sxs-lookup"><span data-stu-id="76fde-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="76fde-169">**Observação:**  Para facilitar a solução de problemas, recomendamos que um elemento **Control** e os elementos filho de **recursos** relacionados sejam adicionados um de cada vez.</span><span class="sxs-lookup"><span data-stu-id="76fde-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="76fde-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="76fde-170">**Script**</span></span>|<span data-ttu-id="76fde-171">Links para o arquivo JavaScript com a definição de função personalizada e o código de registro</span><span class="sxs-lookup"><span data-stu-id="76fde-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="76fde-172">Esse elemento não é usado na Visualização do Desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="76fde-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="76fde-173">Em vez disso, a página HTML é responsável por carregar todos os arquivos JavaScript.</span><span class="sxs-lookup"><span data-stu-id="76fde-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="76fde-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="76fde-174">**Page**</span></span>|<span data-ttu-id="76fde-175">Links para a página HTML de suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="76fde-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="76fde-176">Pontos de extensão para Outlook</span><span class="sxs-lookup"><span data-stu-id="76fde-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="76fde-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76fde-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="76fde-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76fde-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="76fde-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76fde-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="76fde-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76fde-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="76fde-181">[Module](#module) (Só pode ser usado em [DesktopFormFactor](desktopformfactor.md)).</span><span class="sxs-lookup"><span data-stu-id="76fde-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="76fde-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76fde-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="76fde-183">Eventos</span><span class="sxs-lookup"><span data-stu-id="76fde-183">Events</span></span>](#events)
- [<span data-ttu-id="76fde-184">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="76fde-184">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="76fde-185">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76fde-185">MessageReadCommandSurface</span></span>
<span data-ttu-id="76fde-p114">Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de email. No Outlook para área de trabalho, isso aparece na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="76fde-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="76fde-188">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="76fde-188">Child elements</span></span>

|  <span data-ttu-id="76fde-189">Elemento</span><span class="sxs-lookup"><span data-stu-id="76fde-189">Element</span></span> |  <span data-ttu-id="76fde-190">Descrição</span><span class="sxs-lookup"><span data-stu-id="76fde-190">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76fde-191">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76fde-191">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="76fde-192">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="76fde-192">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="76fde-193">CustomTab</span><span class="sxs-lookup"><span data-stu-id="76fde-193">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="76fde-194">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="76fde-194">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="76fde-195">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76fde-195">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="76fde-196">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="76fde-196">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="76fde-197">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76fde-197">MessageComposeCommandSurface</span></span>
<span data-ttu-id="76fde-198">Este ponto de extensão coloca botões na faixa de opções para suplementos que usam o formulário de composição de email.</span><span class="sxs-lookup"><span data-stu-id="76fde-198">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="76fde-199">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="76fde-199">Child elements</span></span>

|  <span data-ttu-id="76fde-200">Elemento</span><span class="sxs-lookup"><span data-stu-id="76fde-200">Element</span></span> |  <span data-ttu-id="76fde-201">Descrição</span><span class="sxs-lookup"><span data-stu-id="76fde-201">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76fde-202">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76fde-202">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="76fde-203">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="76fde-203">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="76fde-204">CustomTab</span><span class="sxs-lookup"><span data-stu-id="76fde-204">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="76fde-205">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="76fde-205">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="76fde-206">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76fde-206">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="76fde-207">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="76fde-207">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="76fde-208">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76fde-208">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="76fde-209">Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="76fde-209">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="76fde-210">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="76fde-210">Child elements</span></span>

|  <span data-ttu-id="76fde-211">Elemento</span><span class="sxs-lookup"><span data-stu-id="76fde-211">Element</span></span> |  <span data-ttu-id="76fde-212">Descrição</span><span class="sxs-lookup"><span data-stu-id="76fde-212">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76fde-213">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76fde-213">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="76fde-214">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="76fde-214">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="76fde-215">CustomTab</span><span class="sxs-lookup"><span data-stu-id="76fde-215">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="76fde-216">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="76fde-216">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="76fde-217">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76fde-217">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="76fde-218">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="76fde-218">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="76fde-219">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76fde-219">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="76fde-220">Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao participante da reunião.</span><span class="sxs-lookup"><span data-stu-id="76fde-220">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="76fde-221">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="76fde-221">Child elements</span></span>

|  <span data-ttu-id="76fde-222">Elemento</span><span class="sxs-lookup"><span data-stu-id="76fde-222">Element</span></span> |  <span data-ttu-id="76fde-223">Descrição</span><span class="sxs-lookup"><span data-stu-id="76fde-223">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76fde-224">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76fde-224">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="76fde-225">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="76fde-225">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="76fde-226">CustomTab</span><span class="sxs-lookup"><span data-stu-id="76fde-226">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="76fde-227">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="76fde-227">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="76fde-228">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76fde-228">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="76fde-229">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="76fde-229">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="76fde-230">Module</span><span class="sxs-lookup"><span data-stu-id="76fde-230">Module</span></span>

<span data-ttu-id="76fde-231">Este ponto de extensão coloca botões na faixa de opções para a extensão do módulo.</span><span class="sxs-lookup"><span data-stu-id="76fde-231">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="76fde-232">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="76fde-232">Child elements</span></span>

|  <span data-ttu-id="76fde-233">Elemento</span><span class="sxs-lookup"><span data-stu-id="76fde-233">Element</span></span> |  <span data-ttu-id="76fde-234">Descrição</span><span class="sxs-lookup"><span data-stu-id="76fde-234">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76fde-235">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="76fde-235">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="76fde-236">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="76fde-236">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="76fde-237">CustomTab</span><span class="sxs-lookup"><span data-stu-id="76fde-237">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="76fde-238">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="76fde-238">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="76fde-239">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="76fde-239">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="76fde-240">Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de email no fator forma móvel.</span><span class="sxs-lookup"><span data-stu-id="76fde-240">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="76fde-241">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="76fde-241">Child elements</span></span>

|  <span data-ttu-id="76fde-242">Elemento</span><span class="sxs-lookup"><span data-stu-id="76fde-242">Element</span></span> |  <span data-ttu-id="76fde-243">Descrição</span><span class="sxs-lookup"><span data-stu-id="76fde-243">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76fde-244">Group</span><span class="sxs-lookup"><span data-stu-id="76fde-244">Group</span></span>](group.md) |  <span data-ttu-id="76fde-245">Adiciona um grupo de botões à superfície de comando.</span><span class="sxs-lookup"><span data-stu-id="76fde-245">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="76fde-246">Os elementos **ExtensionPoint** desse tipo só podem ter um elemento filho: um elemento **Group**.</span><span class="sxs-lookup"><span data-stu-id="76fde-246">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="76fde-247">Os elementos **Control** contidos neste ponto de extensão precisam ter o atributo **xsi:type** definido como `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="76fde-247">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="76fde-248">Exemplo</span><span class="sxs-lookup"><span data-stu-id="76fde-248">Example</span></span>
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

### <a name="events"></a><span data-ttu-id="76fde-249">Eventos</span><span class="sxs-lookup"><span data-stu-id="76fde-249">Events</span></span>

<span data-ttu-id="76fde-250">Este ponto de extensão adiciona um manipulador de eventos para um evento especificado.</span><span class="sxs-lookup"><span data-stu-id="76fde-250">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="76fde-251">Esse tipo de elemento tem suporte no Outlook clássico na Web e na [visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Windows, Mac e Outlook moderno na Web.</span><span class="sxs-lookup"><span data-stu-id="76fde-251">This element type is supported by classic Outlook on the web, and in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Windows, Mac, and modern Outlook on the web.</span></span> <span data-ttu-id="76fde-252">Uma assinatura do Office 365 também é necessária.</span><span class="sxs-lookup"><span data-stu-id="76fde-252">An Office 365 subscription is also required.</span></span>

| <span data-ttu-id="76fde-253">Elemento</span><span class="sxs-lookup"><span data-stu-id="76fde-253">Element</span></span> | <span data-ttu-id="76fde-254">Descrição</span><span class="sxs-lookup"><span data-stu-id="76fde-254">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76fde-255">Event</span><span class="sxs-lookup"><span data-stu-id="76fde-255">Event</span></span>](event.md) |  <span data-ttu-id="76fde-256">Especifica o evento e a função de manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="76fde-256">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="76fde-257">Exemplo do evento ItemSend</span><span class="sxs-lookup"><span data-stu-id="76fde-257">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="76fde-258">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="76fde-258">DetectedEntity</span></span>

<span data-ttu-id="76fde-259">Este ponto extensão adiciona uma ativação do suplemento contextual em um tipo de entidade especificada.</span><span class="sxs-lookup"><span data-stu-id="76fde-259">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="76fde-260">O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="76fde-260">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="76fde-261">Este tipo de elemento está disponível para [ clientes do Outlook que ofereçam suporte a conjuntos de requisitos 1.6 e posteriores](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="76fde-261">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="76fde-262">Elemento</span><span class="sxs-lookup"><span data-stu-id="76fde-262">Element</span></span> |  <span data-ttu-id="76fde-263">Descrição</span><span class="sxs-lookup"><span data-stu-id="76fde-263">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="76fde-264">Label</span><span class="sxs-lookup"><span data-stu-id="76fde-264">Label</span></span>](#label) |  <span data-ttu-id="76fde-265">Especifica o rótulo para o suplemento na janela contextual.</span><span class="sxs-lookup"><span data-stu-id="76fde-265">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="76fde-266">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="76fde-266">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="76fde-267">Especifica a URL para a janela contextual.</span><span class="sxs-lookup"><span data-stu-id="76fde-267">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="76fde-268">Rule</span><span class="sxs-lookup"><span data-stu-id="76fde-268">Rule</span></span>](rule.md) |  <span data-ttu-id="76fde-269">Especifica a regra ou regras que determinam quando um suplemento é ativado.</span><span class="sxs-lookup"><span data-stu-id="76fde-269">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="76fde-270">Label</span><span class="sxs-lookup"><span data-stu-id="76fde-270">Label</span></span>

<span data-ttu-id="76fde-271">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="76fde-271">Required.</span></span> <span data-ttu-id="76fde-272">O rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="76fde-272">The label of the group.</span></span> <span data-ttu-id="76fde-273">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="76fde-273">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="76fde-274">Requisitos de realce</span><span class="sxs-lookup"><span data-stu-id="76fde-274">Highlight requirements</span></span>

<span data-ttu-id="76fde-p117">A única maneira que um usuário pode ativar um suplemento contextual é interagir com uma entidade realçada. Os desenvolvedores podem controlar quais entidades são realçadas usando o atributo `Highlight` do elemento `Rule` para os tipos de regra `ItemHasKnownEntity` e `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="76fde-p117">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="76fde-p118">No entanto, há algumas limitações que devem ser consideradas. Essas limitações são para garantir que sempre haverá uma entidade realçada em compromissos ou mensagens aplicáveis para oferecer ao usuário uma maneira de ativar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="76fde-p118">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="76fde-279">Os tipos de entidade `EmailAddress` e `Url` não podem ser realçados e, portanto, não podem ser usados para ativar um suplemento.</span><span class="sxs-lookup"><span data-stu-id="76fde-279">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="76fde-280">Se for usada uma única regra, `Highlight` DEVERÁ ser definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="76fde-280">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="76fde-281">Se usar um tipo de regra `RuleCollection` com `Mode="AND"` para combinar várias regras, pelo menos uma das regras DEVERÁ ter o `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="76fde-281">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="76fde-282">Se usar um tipo de regra `RuleCollection` com `Mode="OR"` para combinar várias regras, todas as regras DEVERÃO ter o `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="76fde-282">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="76fde-283">Exemplo do evento DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="76fde-283">DetectedEntity event example</span></span>

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
