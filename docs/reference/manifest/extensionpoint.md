---
title: Elemento ExtensionPoint no arquivo de manifesto
description: Define onde um suplemento expõe a funcionalidade na interface de usuário do Office.
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: e5b638969730be47c30c98d4fc231e58d492ac36
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505462"
---
# <a name="extensionpoint-element"></a><span data-ttu-id="99a25-103">Elemento ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="99a25-103">ExtensionPoint element</span></span>

 <span data-ttu-id="99a25-104">Define onde um suplemento expõe a funcionalidade na interface de usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="99a25-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="99a25-105">O elemento **ExtensionPoint** é um elemento filho de [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="99a25-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="99a25-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="99a25-106">Attributes</span></span>

|  <span data-ttu-id="99a25-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="99a25-107">Attribute</span></span>  |  <span data-ttu-id="99a25-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="99a25-108">Required</span></span>  |  <span data-ttu-id="99a25-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="99a25-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="99a25-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="99a25-110">**xsi:type**</span></span>  |  <span data-ttu-id="99a25-111">Sim</span><span class="sxs-lookup"><span data-stu-id="99a25-111">Yes</span></span>  | <span data-ttu-id="99a25-112">O tipo de ponto de extensão que está sendo definido.</span><span class="sxs-lookup"><span data-stu-id="99a25-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="99a25-113">Pontos de extensão somente para Excel</span><span class="sxs-lookup"><span data-stu-id="99a25-113">Extension points for Excel only</span></span>

- <span data-ttu-id="99a25-114">**CustomFunctions**: uma função personalizada escrita em JavaScript para Excel.</span><span class="sxs-lookup"><span data-stu-id="99a25-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="99a25-115">[Este exemplo de código XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) mostra como usar o elemento **ExtensionPoint** com o valor do atributo **CustomFunctions** e os elementos filhos a serem usados.</span><span class="sxs-lookup"><span data-stu-id="99a25-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="99a25-116">Pontos de extensão para comandos de suplemento do Word, Excel, PowerPoint e OneNote</span><span class="sxs-lookup"><span data-stu-id="99a25-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="99a25-117">**PrimaryCommandSurface**, que se refere à faixa de opções no Office.</span><span class="sxs-lookup"><span data-stu-id="99a25-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="99a25-118">**ContextMenu**, que é o menu de atalho exibido ao clicar com o botão direito do mouse na interface de usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="99a25-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="99a25-119">Os exemplos a seguir mostram como usar o elemento **ExtensionPoint** com os valores de atributo **PrimaryCommandSurface** e **ContextMenu** e os elementos filho que devem ser usados com cada um.</span><span class="sxs-lookup"><span data-stu-id="99a25-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="99a25-120">Forneça uma ID exclusiva para os elementos que contêm um atributo ID.</span><span class="sxs-lookup"><span data-stu-id="99a25-120">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="99a25-121">É recomendável usar o nome de sua empresa com a ID.</span><span class="sxs-lookup"><span data-stu-id="99a25-121">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="99a25-122">Por exemplo, use o formato a seguir.</span><span class="sxs-lookup"><span data-stu-id="99a25-122">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

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

#### <a name="child-elements"></a><span data-ttu-id="99a25-123">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="99a25-123">Child elements</span></span>
 
|<span data-ttu-id="99a25-124">Elemento</span><span class="sxs-lookup"><span data-stu-id="99a25-124">Element</span></span>|<span data-ttu-id="99a25-125">Descrição</span><span class="sxs-lookup"><span data-stu-id="99a25-125">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="99a25-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="99a25-126">**CustomTab**</span></span>|<span data-ttu-id="99a25-p103">Obrigatório se você quiser adicionar uma guia personalizada à faixa de opções (usando **PrimaryCommandSurface**). Se você usar o elemento **CustomTab**, o elemento **OfficeTab** não poderá ser usado. O atributo **id** é obrigatório. </span><span class="sxs-lookup"><span data-stu-id="99a25-p103">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**). If you use the **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="99a25-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="99a25-130">**OfficeTab**</span></span>|<span data-ttu-id="99a25-131">Obrigatório se você quiser estender uma guia padrão da faixa de opções do aplicativo do Office (usando **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="99a25-131">Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="99a25-132">Se você usar o elemento **OfficeTab**, o elemento **CustomTab** não poderá ser usado.</span><span class="sxs-lookup"><span data-stu-id="99a25-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="99a25-133">Para saber mais, confira [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="99a25-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="99a25-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="99a25-134">**OfficeMenu**</span></span>|<span data-ttu-id="99a25-p105">Obrigatório se você estiver adicionando comandos de suplemento a um menu de contexto padrão (usando **ContextMenu**). O atributo **id** deve ser definido como: </span><span class="sxs-lookup"><span data-stu-id="99a25-p105">Required if you're adding add-in commands to a default context menu (using **ContextMenu**). The **id** attribute must be set to: </span></span><br/> <span data-ttu-id="99a25-p106">- **ContextMenuText** para o Excel ou Word. Exibe o item no menu de contexto quando o texto for selecionado e o usuário clicar com o botão direito do mouse no texto selecionado. </span><span class="sxs-lookup"><span data-stu-id="99a25-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="99a25-p107">- **ContextMenuCell** para Excel. Exibe o item no menu de contexto quando o usuário clica com o botão direito do mouse em uma célula na planilha.</span><span class="sxs-lookup"><span data-stu-id="99a25-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="99a25-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="99a25-141">**Group**</span></span>|<span data-ttu-id="99a25-p108">Um grupo de pontos de extensão de interface do usuário em uma guia. Um grupo pode ter até seis controles. O atributo **id** é obrigatório. É uma cadeia de caracteres com, no máximo, 125 caracteres. </span><span class="sxs-lookup"><span data-stu-id="99a25-p108">A group of user interface extension points on a tab. A group can have up to six controls. The **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="99a25-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="99a25-145">**Label**</span></span>|<span data-ttu-id="99a25-146">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="99a25-146">Required.</span></span> <span data-ttu-id="99a25-147">O rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="99a25-147">The label of the group.</span></span> <span data-ttu-id="99a25-148">O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um **elemento String.**</span><span class="sxs-lookup"><span data-stu-id="99a25-148">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="99a25-149">O elemento **String** é um elemento filho do elemento **ShortStrings**, que é elemento filho do elemento **Resources**.</span><span class="sxs-lookup"><span data-stu-id="99a25-149">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="99a25-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="99a25-150">**Icon**</span></span>|<span data-ttu-id="99a25-151">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="99a25-151">Required.</span></span> <span data-ttu-id="99a25-152">Especifica o ícone do grupo a ser usado em dispositivos de fator forma pequeno, ou quando muitos botões forem exibidos.</span><span class="sxs-lookup"><span data-stu-id="99a25-152">Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed.</span></span> <span data-ttu-id="99a25-153">O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um **elemento Image.**</span><span class="sxs-lookup"><span data-stu-id="99a25-153">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of an **Image** element.</span></span> <span data-ttu-id="99a25-154">O elemento **Image** é elemento filho do elemento **Images**, que é elemento filho do elemento **Resources**.</span><span class="sxs-lookup"><span data-stu-id="99a25-154">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span></span> <span data-ttu-id="99a25-155">O atributo **size** fornece o tamanho, em pixels, da imagem.</span><span class="sxs-lookup"><span data-stu-id="99a25-155">The **size** attribute gives the size, in pixels, of the image.</span></span> <span data-ttu-id="99a25-156">Três tamanhos de imagem são obrigatórios: 16, 32 e 80 pixels.</span><span class="sxs-lookup"><span data-stu-id="99a25-156">Three image sizes are required: 16, 32, and 80.</span></span> <span data-ttu-id="99a25-157">Também há suporte para cinco tamanhos opcionais: 20, 24, 40, 48 e 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="99a25-157">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="99a25-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="99a25-158">**Tooltip**</span></span>|<span data-ttu-id="99a25-159">Opcional.</span><span class="sxs-lookup"><span data-stu-id="99a25-159">Optional.</span></span> <span data-ttu-id="99a25-160">A dica de ferramenta do grupo.</span><span class="sxs-lookup"><span data-stu-id="99a25-160">The tooltip of the group.</span></span> <span data-ttu-id="99a25-161">O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um **elemento String.**</span><span class="sxs-lookup"><span data-stu-id="99a25-161">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="99a25-162">O elemento **String** é um elemento filho do elemento **LongStrings**, que, por sua vez, é um elemento filho do elemento **Resources**.</span><span class="sxs-lookup"><span data-stu-id="99a25-162">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="99a25-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="99a25-163">**Control**</span></span>|<span data-ttu-id="99a25-164">Cada grupo exige pelo menos um controle.</span><span class="sxs-lookup"><span data-stu-id="99a25-164">Each group requires at least one control.</span></span> <span data-ttu-id="99a25-165">Um elemento **Control** pode ser um **Button** ou um **Menu**.</span><span class="sxs-lookup"><span data-stu-id="99a25-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="99a25-166">Use **Menu** para especificar uma lista suspensa de controles de botão.</span><span class="sxs-lookup"><span data-stu-id="99a25-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="99a25-167">Atualmente, há suporte apenas para botões e menus.</span><span class="sxs-lookup"><span data-stu-id="99a25-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="99a25-168">Confira as seções [Controles de botão](control.md#button-control) e [Controles de menu](control.md#menu-dropdown-button-controls) para saber mais.</span><span class="sxs-lookup"><span data-stu-id="99a25-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="99a25-169">**Observação:**  Para facilitar a solução de problemas, recomendamos que um elemento **Control** e os elementos filho **dos Recursos** relacionados sejam adicionados um de cada vez.</span><span class="sxs-lookup"><span data-stu-id="99a25-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="99a25-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="99a25-170">**Script**</span></span>|<span data-ttu-id="99a25-171">Links para o arquivo JavaScript com a definição de função personalizada e o código de registro</span><span class="sxs-lookup"><span data-stu-id="99a25-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="99a25-172">Esse elemento não é usado na Visualização do Desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="99a25-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="99a25-173">Em vez disso, a página HTML é responsável por carregar todos os arquivos JavaScript.</span><span class="sxs-lookup"><span data-stu-id="99a25-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="99a25-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="99a25-174">**Page**</span></span>|<span data-ttu-id="99a25-175">Links para a página HTML de suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="99a25-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="99a25-176">Pontos de extensão para Outlook</span><span class="sxs-lookup"><span data-stu-id="99a25-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="99a25-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="99a25-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="99a25-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="99a25-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="99a25-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="99a25-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="99a25-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="99a25-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="99a25-181">[Module](#module) (Só pode ser usado em [DesktopFormFactor](desktopformfactor.md)).</span><span class="sxs-lookup"><span data-stu-id="99a25-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="99a25-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="99a25-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="99a25-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="99a25-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface)
- [<span data-ttu-id="99a25-184">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="99a25-184">LaunchEvent</span></span>](#launchevent-preview)
- [<span data-ttu-id="99a25-185">Eventos</span><span class="sxs-lookup"><span data-stu-id="99a25-185">Events</span></span>](#events)
- [<span data-ttu-id="99a25-186">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="99a25-186">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="99a25-187">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="99a25-187">MessageReadCommandSurface</span></span>

<span data-ttu-id="99a25-p114">Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de email. No Outlook para área de trabalho, isso aparece na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="99a25-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="99a25-190">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="99a25-190">Child elements</span></span>

|  <span data-ttu-id="99a25-191">Elemento</span><span class="sxs-lookup"><span data-stu-id="99a25-191">Element</span></span> |  <span data-ttu-id="99a25-192">Descrição</span><span class="sxs-lookup"><span data-stu-id="99a25-192">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="99a25-193">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="99a25-193">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="99a25-194">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="99a25-194">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="99a25-195">CustomTab</span><span class="sxs-lookup"><span data-stu-id="99a25-195">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="99a25-196">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="99a25-196">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="99a25-197">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="99a25-197">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="99a25-198">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="99a25-198">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="99a25-199">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="99a25-199">MessageComposeCommandSurface</span></span>

<span data-ttu-id="99a25-200">Este ponto de extensão coloca botões na faixa de opções para suplementos que usam o formulário de composição de email.</span><span class="sxs-lookup"><span data-stu-id="99a25-200">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="99a25-201">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="99a25-201">Child elements</span></span>

|  <span data-ttu-id="99a25-202">Elemento</span><span class="sxs-lookup"><span data-stu-id="99a25-202">Element</span></span> |  <span data-ttu-id="99a25-203">Descrição</span><span class="sxs-lookup"><span data-stu-id="99a25-203">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="99a25-204">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="99a25-204">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="99a25-205">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="99a25-205">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="99a25-206">CustomTab</span><span class="sxs-lookup"><span data-stu-id="99a25-206">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="99a25-207">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="99a25-207">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="99a25-208">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="99a25-208">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="99a25-209">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="99a25-209">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="99a25-210">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="99a25-210">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="99a25-211">Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="99a25-211">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="99a25-212">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="99a25-212">Child elements</span></span>

|  <span data-ttu-id="99a25-213">Elemento</span><span class="sxs-lookup"><span data-stu-id="99a25-213">Element</span></span> |  <span data-ttu-id="99a25-214">Descrição</span><span class="sxs-lookup"><span data-stu-id="99a25-214">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="99a25-215">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="99a25-215">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="99a25-216">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="99a25-216">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="99a25-217">CustomTab</span><span class="sxs-lookup"><span data-stu-id="99a25-217">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="99a25-218">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="99a25-218">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="99a25-219">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="99a25-219">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="99a25-220">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="99a25-220">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="99a25-221">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="99a25-221">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="99a25-222">Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao participante da reunião.</span><span class="sxs-lookup"><span data-stu-id="99a25-222">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="99a25-223">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="99a25-223">Child elements</span></span>

|  <span data-ttu-id="99a25-224">Elemento</span><span class="sxs-lookup"><span data-stu-id="99a25-224">Element</span></span> |  <span data-ttu-id="99a25-225">Descrição</span><span class="sxs-lookup"><span data-stu-id="99a25-225">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="99a25-226">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="99a25-226">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="99a25-227">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="99a25-227">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="99a25-228">CustomTab</span><span class="sxs-lookup"><span data-stu-id="99a25-228">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="99a25-229">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="99a25-229">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="99a25-230">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="99a25-230">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="99a25-231">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="99a25-231">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="99a25-232">Module</span><span class="sxs-lookup"><span data-stu-id="99a25-232">Module</span></span>

<span data-ttu-id="99a25-233">Este ponto de extensão coloca botões na faixa de opções para a extensão do módulo.</span><span class="sxs-lookup"><span data-stu-id="99a25-233">This extension point puts buttons on the ribbon for the module extension.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="99a25-234">Registrar eventos [de Caixa de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) Correio e [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) não está disponível com esse ponto de extensão.</span><span class="sxs-lookup"><span data-stu-id="99a25-234">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="99a25-235">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="99a25-235">Child elements</span></span>

|  <span data-ttu-id="99a25-236">Elemento</span><span class="sxs-lookup"><span data-stu-id="99a25-236">Element</span></span> |  <span data-ttu-id="99a25-237">Descrição</span><span class="sxs-lookup"><span data-stu-id="99a25-237">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="99a25-238">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="99a25-238">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="99a25-239">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="99a25-239">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="99a25-240">CustomTab</span><span class="sxs-lookup"><span data-stu-id="99a25-240">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="99a25-241">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="99a25-241">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="99a25-242">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="99a25-242">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="99a25-243">Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de email no fator forma móvel.</span><span class="sxs-lookup"><span data-stu-id="99a25-243">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="99a25-244">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="99a25-244">Child elements</span></span>

|  <span data-ttu-id="99a25-245">Elemento</span><span class="sxs-lookup"><span data-stu-id="99a25-245">Element</span></span> |  <span data-ttu-id="99a25-246">Descrição</span><span class="sxs-lookup"><span data-stu-id="99a25-246">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="99a25-247">Group</span><span class="sxs-lookup"><span data-stu-id="99a25-247">Group</span></span>](group.md) |  <span data-ttu-id="99a25-248">Adiciona um grupo de botões à superfície de comando.</span><span class="sxs-lookup"><span data-stu-id="99a25-248">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="99a25-249">Os elementos **ExtensionPoint** desse tipo só podem ter um elemento filho: um elemento **Group**.</span><span class="sxs-lookup"><span data-stu-id="99a25-249">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="99a25-250">Os elementos **Control** contidos neste ponto de extensão precisam ter o atributo **xsi:type** definido como `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="99a25-250">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="99a25-251">Exemplo</span><span class="sxs-lookup"><span data-stu-id="99a25-251">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface"></a><span data-ttu-id="99a25-252">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="99a25-252">MobileOnlineMeetingCommandSurface</span></span>

<span data-ttu-id="99a25-253">Esse ponto de extensão coloca uma alternância apropriada para o modo na superfície de comando para um compromisso no fator de forma móvel.</span><span class="sxs-lookup"><span data-stu-id="99a25-253">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="99a25-254">Um organizador de reunião pode criar uma reunião online.</span><span class="sxs-lookup"><span data-stu-id="99a25-254">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="99a25-255">Um participante pode participar posteriormente da reunião online.</span><span class="sxs-lookup"><span data-stu-id="99a25-255">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="99a25-256">Para saber mais sobre esse cenário, consulte o artigo Criar um complemento [móvel](../../outlook/online-meeting.md) do Outlook para um provedor de reunião online.</span><span class="sxs-lookup"><span data-stu-id="99a25-256">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

> [!NOTE]
> <span data-ttu-id="99a25-257">Esse ponto de extensão só tem suporte para Android e iOS com uma assinatura do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="99a25-257">This extension point is only supported on Android and iOS with a Microsoft 365 subscription.</span></span>
>
> <span data-ttu-id="99a25-258">Registrar eventos [de Caixa de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) Correio e [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) não está disponível com esse ponto de extensão.</span><span class="sxs-lookup"><span data-stu-id="99a25-258">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="99a25-259">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="99a25-259">Child elements</span></span>

|  <span data-ttu-id="99a25-260">Elemento</span><span class="sxs-lookup"><span data-stu-id="99a25-260">Element</span></span> |  <span data-ttu-id="99a25-261">Descrição</span><span class="sxs-lookup"><span data-stu-id="99a25-261">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="99a25-262">Control</span><span class="sxs-lookup"><span data-stu-id="99a25-262">Control</span></span>](control.md) |  <span data-ttu-id="99a25-263">Adiciona um botão à superfície de comando.</span><span class="sxs-lookup"><span data-stu-id="99a25-263">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="99a25-264">`ExtensionPoint` elementos desse tipo só podem ter um elemento filho: um `Control` elemento.</span><span class="sxs-lookup"><span data-stu-id="99a25-264">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="99a25-265">O `Control` elemento contido neste ponto de extensão deve ter o atributo definido como `xsi:type` `MobileButton` .</span><span class="sxs-lookup"><span data-stu-id="99a25-265">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="99a25-266">As `Icon` imagens devem estar em escala de cinza usando código hexaxa `#919191` ou seu equivalente em outros [formatos de cor.](https://convertingcolors.com/hex-color-919191.html)</span><span class="sxs-lookup"><span data-stu-id="99a25-266">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="99a25-267">Exemplo</span><span class="sxs-lookup"><span data-stu-id="99a25-267">Example</span></span>

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

### <a name="launchevent-preview"></a><span data-ttu-id="99a25-268">LaunchEvent (visualização)</span><span class="sxs-lookup"><span data-stu-id="99a25-268">LaunchEvent (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="99a25-269">Esse ponto de extensão só é suportado na [visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Outlook na Web e no Windows com uma assinatura do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="99a25-269">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="99a25-270">Esse ponto de extensão permite que um complemento seja ativado com base em eventos suportados no fator de formulário da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="99a25-270">This extension point enables an add-in to activate based on supported events in the desktop form factor.</span></span> <span data-ttu-id="99a25-271">Atualmente, os únicos eventos com suporte `OnNewMessageCompose` são e `OnNewAppointmentOrganizer` .</span><span class="sxs-lookup"><span data-stu-id="99a25-271">Currently, the only supported events are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> <span data-ttu-id="99a25-272">Para saber mais sobre esse cenário, consulte o artigo Configurar seu [complemento do Outlook para ativação baseada em](../../outlook/autolaunch.md) eventos.</span><span class="sxs-lookup"><span data-stu-id="99a25-272">To learn more about this scenario, see the [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md) article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="99a25-273">Registrar eventos [de Caixa de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) Correio e [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) não está disponível com esse ponto de extensão.</span><span class="sxs-lookup"><span data-stu-id="99a25-273">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="99a25-274">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="99a25-274">Child elements</span></span>

|  <span data-ttu-id="99a25-275">Elemento</span><span class="sxs-lookup"><span data-stu-id="99a25-275">Element</span></span> |  <span data-ttu-id="99a25-276">Descrição</span><span class="sxs-lookup"><span data-stu-id="99a25-276">Description</span></span>  |
|:-----|:-----|
| [<span data-ttu-id="99a25-277">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="99a25-277">LaunchEvents</span></span>](launchevents.md) |  <span data-ttu-id="99a25-278">Lista de [LaunchEvent](launchevent.md) para ativação baseada em evento.</span><span class="sxs-lookup"><span data-stu-id="99a25-278">List of [LaunchEvent](launchevent.md) for event-based activation.</span></span>  |
| [<span data-ttu-id="99a25-279">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="99a25-279">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="99a25-280">O local do arquivo JavaScript de origem.</span><span class="sxs-lookup"><span data-stu-id="99a25-280">The location of the source JavaScript file.</span></span>  |

#### <a name="example"></a><span data-ttu-id="99a25-281">Exemplo</span><span class="sxs-lookup"><span data-stu-id="99a25-281">Example</span></span>

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

### <a name="events"></a><span data-ttu-id="99a25-282">Eventos</span><span class="sxs-lookup"><span data-stu-id="99a25-282">Events</span></span>

<span data-ttu-id="99a25-283">Este ponto de extensão adiciona um manipulador de eventos para um evento especificado.</span><span class="sxs-lookup"><span data-stu-id="99a25-283">This extension point adds an event handler for a specified event.</span></span> <span data-ttu-id="99a25-284">Para obter mais informações sobre como usar esse ponto de extensão, consulte Recurso Ao [enviar para os complementos do Outlook](../../outlook/outlook-on-send-addins.md).</span><span class="sxs-lookup"><span data-stu-id="99a25-284">For more information about using this extension point, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="99a25-285">Registrar eventos [de Caixa de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) Correio e [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) não está disponível com esse ponto de extensão.</span><span class="sxs-lookup"><span data-stu-id="99a25-285">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

| <span data-ttu-id="99a25-286">Elemento</span><span class="sxs-lookup"><span data-stu-id="99a25-286">Element</span></span> | <span data-ttu-id="99a25-287">Descrição</span><span class="sxs-lookup"><span data-stu-id="99a25-287">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="99a25-288">Event</span><span class="sxs-lookup"><span data-stu-id="99a25-288">Event</span></span>](event.md) |  <span data-ttu-id="99a25-289">Especifica o evento e a função de manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="99a25-289">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="99a25-290">Exemplo do evento ItemSend</span><span class="sxs-lookup"><span data-stu-id="99a25-290">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="99a25-291">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="99a25-291">DetectedEntity</span></span>

<span data-ttu-id="99a25-292">Este ponto extensão adiciona uma ativação do suplemento contextual em um tipo de entidade especificada.</span><span class="sxs-lookup"><span data-stu-id="99a25-292">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="99a25-293">Registrar eventos [de Caixa de](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) Correio e [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) não está disponível com esse ponto de extensão.</span><span class="sxs-lookup"><span data-stu-id="99a25-293">Registering [Mailbox](../objectmodel/preview-requirement-set/office.context.mailbox.md#events) and [Item](../objectmodel/preview-requirement-set/office.context.mailbox.item.md#events) events is not available with this extension point.</span></span>

<span data-ttu-id="99a25-294">O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="99a25-294">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="99a25-295">Este tipo de elemento está disponível para [ clientes do Outlook que ofereçam suporte a conjuntos de requisitos 1.6 e posteriores](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="99a25-295">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="99a25-296">Elemento</span><span class="sxs-lookup"><span data-stu-id="99a25-296">Element</span></span> |  <span data-ttu-id="99a25-297">Descrição</span><span class="sxs-lookup"><span data-stu-id="99a25-297">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="99a25-298">Label</span><span class="sxs-lookup"><span data-stu-id="99a25-298">Label</span></span>](#label) |  <span data-ttu-id="99a25-299">Especifica o rótulo para o suplemento na janela contextual.</span><span class="sxs-lookup"><span data-stu-id="99a25-299">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="99a25-300">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="99a25-300">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="99a25-301">Especifica a URL para a janela contextual.</span><span class="sxs-lookup"><span data-stu-id="99a25-301">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="99a25-302">Rule</span><span class="sxs-lookup"><span data-stu-id="99a25-302">Rule</span></span>](rule.md) |  <span data-ttu-id="99a25-303">Especifica a regra ou regras que determinam quando um suplemento é ativado.</span><span class="sxs-lookup"><span data-stu-id="99a25-303">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="99a25-304">Label</span><span class="sxs-lookup"><span data-stu-id="99a25-304">Label</span></span>

<span data-ttu-id="99a25-305">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="99a25-305">Required.</span></span> <span data-ttu-id="99a25-306">O rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="99a25-306">The label of the group.</span></span> <span data-ttu-id="99a25-307">O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no [elemento Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="99a25-307">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="99a25-308">Requisitos de realce</span><span class="sxs-lookup"><span data-stu-id="99a25-308">Highlight requirements</span></span>

<span data-ttu-id="99a25-p119">A única maneira que um usuário pode ativar um suplemento contextual é interagir com uma entidade realçada. Os desenvolvedores podem controlar quais entidades são realçadas usando o atributo `Highlight` do elemento `Rule` para os tipos de regra `ItemHasKnownEntity` e `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="99a25-p119">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="99a25-p120">No entanto, há algumas limitações que devem ser consideradas. Essas limitações são para garantir que sempre haverá uma entidade realçada em compromissos ou mensagens aplicáveis para oferecer ao usuário uma maneira de ativar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="99a25-p120">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="99a25-313">Os tipos de entidade `EmailAddress` e `Url` não podem ser realçados e, portanto, não podem ser usados para ativar um suplemento.</span><span class="sxs-lookup"><span data-stu-id="99a25-313">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="99a25-314">Se for usada uma única regra, `Highlight` DEVERÁ ser definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="99a25-314">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="99a25-315">Se usar um tipo de regra `RuleCollection` com `Mode="AND"` para combinar várias regras, pelo menos uma das regras DEVERÁ ter o `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="99a25-315">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="99a25-316">Se usar um tipo de regra `RuleCollection` com `Mode="OR"` para combinar várias regras, todas as regras DEVERÃO ter o `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="99a25-316">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="99a25-317">Exemplo do evento DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="99a25-317">DetectedEntity event example</span></span>

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
