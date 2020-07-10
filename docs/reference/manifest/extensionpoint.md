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
# <a name="extensionpoint-element"></a><span data-ttu-id="85321-103">Elemento ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="85321-103">ExtensionPoint element</span></span>

 <span data-ttu-id="85321-104">Define onde um suplemento expõe a funcionalidade na interface de usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="85321-104">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="85321-105">O elemento **ExtensionPoint** é um elemento filho de [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="85321-105">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="85321-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="85321-106">Attributes</span></span>

|  <span data-ttu-id="85321-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="85321-107">Attribute</span></span>  |  <span data-ttu-id="85321-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="85321-108">Required</span></span>  |  <span data-ttu-id="85321-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="85321-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="85321-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="85321-110">**xsi:type**</span></span>  |  <span data-ttu-id="85321-111">Sim</span><span class="sxs-lookup"><span data-stu-id="85321-111">Yes</span></span>  | <span data-ttu-id="85321-112">O tipo de ponto de extensão que está sendo definido.</span><span class="sxs-lookup"><span data-stu-id="85321-112">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="85321-113">Pontos de extensão somente para Excel</span><span class="sxs-lookup"><span data-stu-id="85321-113">Extension points for Excel only</span></span>

- <span data-ttu-id="85321-114">**CustomFunctions**: uma função personalizada escrita em JavaScript para Excel.</span><span class="sxs-lookup"><span data-stu-id="85321-114">**CustomFunctions** - A custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="85321-115">[Este exemplo de código XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) mostra como usar o elemento **ExtensionPoint** com o valor do atributo **CustomFunctions** e os elementos filhos a serem usados.</span><span class="sxs-lookup"><span data-stu-id="85321-115">[This XML code sample](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/manifest.xml) shows how to use the **ExtensionPoint** element with the **CustomFunctions** attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="85321-116">Pontos de extensão para comandos de suplemento do Word, Excel, PowerPoint e OneNote</span><span class="sxs-lookup"><span data-stu-id="85321-116">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="85321-117">**PrimaryCommandSurface**, que se refere à faixa de opções no Office.</span><span class="sxs-lookup"><span data-stu-id="85321-117">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="85321-118">**ContextMenu**, que é o menu de atalho exibido ao clicar com o botão direito do mouse na interface de usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="85321-118">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="85321-119">Os exemplos a seguir mostram como usar o elemento **ExtensionPoint** com os valores de atributo **PrimaryCommandSurface** e **ContextMenu** e os elementos filho que devem ser usados com cada um.</span><span class="sxs-lookup"><span data-stu-id="85321-119">The following examples show how to use the **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="85321-120">Forneça uma ID exclusiva para os elementos que contêm um atributo ID.</span><span class="sxs-lookup"><span data-stu-id="85321-120">For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="85321-121">É recomendável usar o nome de sua empresa com a ID.</span><span class="sxs-lookup"><span data-stu-id="85321-121">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="85321-122">Por exemplo, use o formato a seguir.</span><span class="sxs-lookup"><span data-stu-id="85321-122">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

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

#### <a name="child-elements"></a><span data-ttu-id="85321-123">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="85321-123">Child elements</span></span>
 
|<span data-ttu-id="85321-124">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="85321-124">**Element**</span></span>|<span data-ttu-id="85321-125">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="85321-125">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="85321-126">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="85321-126">**CustomTab**</span></span>|<span data-ttu-id="85321-127">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="85321-127">Required if you want to add a custom tab to the ribbon (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="85321-128">If you use the **CustomTab** element, you can't use the **OfficeTab** element.</span><span class="sxs-lookup"><span data-stu-id="85321-128">If you use the **CustomTab** element, you can't use the **OfficeTab** element.</span></span> <span data-ttu-id="85321-129">The **id** attribute is required.</span><span class="sxs-lookup"><span data-stu-id="85321-129">The **id** attribute is required.</span></span>|
|<span data-ttu-id="85321-130">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="85321-130">**OfficeTab**</span></span>|<span data-ttu-id="85321-131">Obrigatório se você deseja estender uma guia padrão da faixa de opções do aplicativo do Office (usando **PrimaryCommandSurface**).</span><span class="sxs-lookup"><span data-stu-id="85321-131">Required if you want to extend a default Office app ribbon tab (using **PrimaryCommandSurface**).</span></span> <span data-ttu-id="85321-132">Se você usar o elemento **OfficeTab**, o elemento **CustomTab** não poderá ser usado.</span><span class="sxs-lookup"><span data-stu-id="85321-132">If you use the **OfficeTab** element, you can't use the **CustomTab** element.</span></span> <span data-ttu-id="85321-133">Para saber mais, confira [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="85321-133">For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="85321-134">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="85321-134">**OfficeMenu**</span></span>|<span data-ttu-id="85321-135">Required if you're adding add-in commands to a default context menu (using **ContextMenu**).</span><span class="sxs-lookup"><span data-stu-id="85321-135">Required if you're adding add-in commands to a default context menu (using **ContextMenu**).</span></span> <span data-ttu-id="85321-136">The **id** attribute must be set to:</span><span class="sxs-lookup"><span data-stu-id="85321-136">The **id** attribute must be set to:</span></span> <br/> <span data-ttu-id="85321-137">- **ContextMenuText** for Excel or Word.</span><span class="sxs-lookup"><span data-stu-id="85321-137">- **ContextMenuText** for Excel or Word.</span></span> <span data-ttu-id="85321-138">Displays the item on the context menu when text is selected and then the user right-clicks on the selected text.</span><span class="sxs-lookup"><span data-stu-id="85321-138">Displays the item on the context menu when text is selected and then the user right-clicks on the selected text.</span></span> <br/> <span data-ttu-id="85321-139">- **ContextMenuCell** for Excel.</span><span class="sxs-lookup"><span data-stu-id="85321-139">- **ContextMenuCell** for Excel.</span></span> <span data-ttu-id="85321-140">Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span><span class="sxs-lookup"><span data-stu-id="85321-140">Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="85321-141">**Group**</span><span class="sxs-lookup"><span data-stu-id="85321-141">**Group**</span></span>|<span data-ttu-id="85321-142">A group of user interface extension points on a tab. A group can have up to six controls.</span><span class="sxs-lookup"><span data-stu-id="85321-142">A group of user interface extension points on a tab. A group can have up to six controls.</span></span> <span data-ttu-id="85321-143">The **id** attribute is required.</span><span class="sxs-lookup"><span data-stu-id="85321-143">The **id** attribute is required.</span></span> <span data-ttu-id="85321-144">It's a string with a maximum of 125 characters.</span><span class="sxs-lookup"><span data-stu-id="85321-144">It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="85321-145">**Label**</span><span class="sxs-lookup"><span data-stu-id="85321-145">**Label**</span></span>|<span data-ttu-id="85321-146">Required.</span><span class="sxs-lookup"><span data-stu-id="85321-146">Required.</span></span> <span data-ttu-id="85321-147">The label of the group.</span><span class="sxs-lookup"><span data-stu-id="85321-147">The label of the group.</span></span> <span data-ttu-id="85321-148">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="85321-148">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="85321-149">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="85321-149">The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="85321-150">**Icon**</span><span class="sxs-lookup"><span data-stu-id="85321-150">**Icon**</span></span>|<span data-ttu-id="85321-151">Required.</span><span class="sxs-lookup"><span data-stu-id="85321-151">Required.</span></span> <span data-ttu-id="85321-152">Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed.</span><span class="sxs-lookup"><span data-stu-id="85321-152">Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed.</span></span> <span data-ttu-id="85321-153">The **resid** attribute must be set to the value of the **id** attribute of an **Image** element.</span><span class="sxs-lookup"><span data-stu-id="85321-153">The **resid** attribute must be set to the value of the **id** attribute of an **Image** element.</span></span> <span data-ttu-id="85321-154">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="85321-154">The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element.</span></span> <span data-ttu-id="85321-155">The **size** attribute gives the size, in pixels, of the image.</span><span class="sxs-lookup"><span data-stu-id="85321-155">The **size** attribute gives the size, in pixels, of the image.</span></span> <span data-ttu-id="85321-156">Three image sizes are required: 16, 32, and 80.</span><span class="sxs-lookup"><span data-stu-id="85321-156">Three image sizes are required: 16, 32, and 80.</span></span> <span data-ttu-id="85321-157">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span><span class="sxs-lookup"><span data-stu-id="85321-157">Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="85321-158">**Tooltip**</span><span class="sxs-lookup"><span data-stu-id="85321-158">**Tooltip**</span></span>|<span data-ttu-id="85321-159">Optional.</span><span class="sxs-lookup"><span data-stu-id="85321-159">Optional.</span></span> <span data-ttu-id="85321-160">The tooltip of the group.</span><span class="sxs-lookup"><span data-stu-id="85321-160">The tooltip of the group.</span></span> <span data-ttu-id="85321-161">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span><span class="sxs-lookup"><span data-stu-id="85321-161">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="85321-162">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span><span class="sxs-lookup"><span data-stu-id="85321-162">The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="85321-163">**Control**</span><span class="sxs-lookup"><span data-stu-id="85321-163">**Control**</span></span>|<span data-ttu-id="85321-164">Cada grupo exige pelo menos um controle.</span><span class="sxs-lookup"><span data-stu-id="85321-164">Each group requires at least one control.</span></span> <span data-ttu-id="85321-165">Um elemento **Control** pode ser um **Button** ou um **Menu**.</span><span class="sxs-lookup"><span data-stu-id="85321-165">A **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="85321-166">Use **Menu** para especificar uma lista suspensa de controles de botão.</span><span class="sxs-lookup"><span data-stu-id="85321-166">Use **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="85321-167">Atualmente, há suporte apenas para botões e menus.</span><span class="sxs-lookup"><span data-stu-id="85321-167">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="85321-168">Confira as seções [Controles de botão](control.md#button-control) e [Controles de menu](control.md#menu-dropdown-button-controls) para saber mais.</span><span class="sxs-lookup"><span data-stu-id="85321-168">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="85321-169">**Observação:**  Para facilitar a solução de problemas, recomendamos que um elemento **Control** e os elementos filho de **recursos** relacionados sejam adicionados um de cada vez.</span><span class="sxs-lookup"><span data-stu-id="85321-169">**Note:**  To make troubleshooting easier, we recommend that a **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="85321-170">**Script**</span><span class="sxs-lookup"><span data-stu-id="85321-170">**Script**</span></span>|<span data-ttu-id="85321-171">Links para o arquivo JavaScript com a definição de função personalizada e o código de registro</span><span class="sxs-lookup"><span data-stu-id="85321-171">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="85321-172">Esse elemento não é usado na Visualização do Desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="85321-172">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="85321-173">Em vez disso, a página HTML é responsável por carregar todos os arquivos JavaScript.</span><span class="sxs-lookup"><span data-stu-id="85321-173">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="85321-174">**Page**</span><span class="sxs-lookup"><span data-stu-id="85321-174">**Page**</span></span>|<span data-ttu-id="85321-175">Links para a página HTML de suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="85321-175">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="85321-176">Pontos de extensão para Outlook</span><span class="sxs-lookup"><span data-stu-id="85321-176">Extension points for Outlook</span></span>

- [<span data-ttu-id="85321-177">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="85321-177">MessageReadCommandSurface</span></span>](#messagereadcommandsurface)
- [<span data-ttu-id="85321-178">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="85321-178">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface)
- [<span data-ttu-id="85321-179">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="85321-179">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface)
- [<span data-ttu-id="85321-180">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="85321-180">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="85321-181">[Module](#module) (Só pode ser usado em [DesktopFormFactor](desktopformfactor.md)).</span><span class="sxs-lookup"><span data-stu-id="85321-181">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="85321-182">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="85321-182">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="85321-183">MobileOnlineMeetingCommandSurface</span><span class="sxs-lookup"><span data-stu-id="85321-183">MobileOnlineMeetingCommandSurface</span></span>](#mobileonlinemeetingcommandsurface-preview)
- [<span data-ttu-id="85321-184">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="85321-184">LaunchEvent</span></span>](#launchevent-preview)
- [<span data-ttu-id="85321-185">Eventos</span><span class="sxs-lookup"><span data-stu-id="85321-185">Events</span></span>](#events)
- [<span data-ttu-id="85321-186">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="85321-186">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="85321-187">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="85321-187">MessageReadCommandSurface</span></span>

<span data-ttu-id="85321-188">This extension point puts buttons in the command surface for the mail read view.</span><span class="sxs-lookup"><span data-stu-id="85321-188">This extension point puts buttons in the command surface for the mail read view.</span></span> <span data-ttu-id="85321-189">In Outlook desktop, this appears in the ribbon.</span><span class="sxs-lookup"><span data-stu-id="85321-189">In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="85321-190">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="85321-190">Child elements</span></span>

|  <span data-ttu-id="85321-191">Elemento</span><span class="sxs-lookup"><span data-stu-id="85321-191">Element</span></span> |  <span data-ttu-id="85321-192">Descrição</span><span class="sxs-lookup"><span data-stu-id="85321-192">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="85321-193">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="85321-193">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="85321-194">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="85321-194">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="85321-195">CustomTab</span><span class="sxs-lookup"><span data-stu-id="85321-195">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="85321-196">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="85321-196">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="85321-197">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="85321-197">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="85321-198">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="85321-198">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="85321-199">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="85321-199">MessageComposeCommandSurface</span></span>

<span data-ttu-id="85321-200">Este ponto de extensão coloca botões na faixa de opções para suplementos que usam o formulário de composição de email.</span><span class="sxs-lookup"><span data-stu-id="85321-200">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="85321-201">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="85321-201">Child elements</span></span>

|  <span data-ttu-id="85321-202">Elemento</span><span class="sxs-lookup"><span data-stu-id="85321-202">Element</span></span> |  <span data-ttu-id="85321-203">Descrição</span><span class="sxs-lookup"><span data-stu-id="85321-203">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="85321-204">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="85321-204">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="85321-205">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="85321-205">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="85321-206">CustomTab</span><span class="sxs-lookup"><span data-stu-id="85321-206">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="85321-207">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="85321-207">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="85321-208">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="85321-208">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="85321-209">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="85321-209">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="85321-210">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="85321-210">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="85321-211">Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="85321-211">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="85321-212">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="85321-212">Child elements</span></span>

|  <span data-ttu-id="85321-213">Elemento</span><span class="sxs-lookup"><span data-stu-id="85321-213">Element</span></span> |  <span data-ttu-id="85321-214">Descrição</span><span class="sxs-lookup"><span data-stu-id="85321-214">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="85321-215">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="85321-215">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="85321-216">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="85321-216">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="85321-217">CustomTab</span><span class="sxs-lookup"><span data-stu-id="85321-217">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="85321-218">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="85321-218">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="85321-219">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="85321-219">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="85321-220">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="85321-220">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="85321-221">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="85321-221">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="85321-222">Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao participante da reunião.</span><span class="sxs-lookup"><span data-stu-id="85321-222">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="85321-223">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="85321-223">Child elements</span></span>

|  <span data-ttu-id="85321-224">Elemento</span><span class="sxs-lookup"><span data-stu-id="85321-224">Element</span></span> |  <span data-ttu-id="85321-225">Descrição</span><span class="sxs-lookup"><span data-stu-id="85321-225">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="85321-226">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="85321-226">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="85321-227">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="85321-227">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="85321-228">CustomTab</span><span class="sxs-lookup"><span data-stu-id="85321-228">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="85321-229">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="85321-229">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="85321-230">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="85321-230">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="85321-231">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="85321-231">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="85321-232">Module</span><span class="sxs-lookup"><span data-stu-id="85321-232">Module</span></span>

<span data-ttu-id="85321-233">Este ponto de extensão coloca botões na faixa de opções para a extensão do módulo.</span><span class="sxs-lookup"><span data-stu-id="85321-233">This extension point puts buttons on the ribbon for the module extension.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="85321-234">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="85321-234">Child elements</span></span>

|  <span data-ttu-id="85321-235">Elemento</span><span class="sxs-lookup"><span data-stu-id="85321-235">Element</span></span> |  <span data-ttu-id="85321-236">Descrição</span><span class="sxs-lookup"><span data-stu-id="85321-236">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="85321-237">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="85321-237">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="85321-238">Adiciona os comandos à guia da faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="85321-238">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="85321-239">CustomTab</span><span class="sxs-lookup"><span data-stu-id="85321-239">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="85321-240">Adiciona os comandos à guia da faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="85321-240">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="85321-241">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="85321-241">MobileMessageReadCommandSurface</span></span>

<span data-ttu-id="85321-242">Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de email no fator forma móvel.</span><span class="sxs-lookup"><span data-stu-id="85321-242">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="85321-243">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="85321-243">Child elements</span></span>

|  <span data-ttu-id="85321-244">Elemento</span><span class="sxs-lookup"><span data-stu-id="85321-244">Element</span></span> |  <span data-ttu-id="85321-245">Descrição</span><span class="sxs-lookup"><span data-stu-id="85321-245">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="85321-246">Group</span><span class="sxs-lookup"><span data-stu-id="85321-246">Group</span></span>](group.md) |  <span data-ttu-id="85321-247">Adiciona um grupo de botões à superfície de comando.</span><span class="sxs-lookup"><span data-stu-id="85321-247">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="85321-248">Os elementos **ExtensionPoint** desse tipo só podem ter um elemento filho: um elemento **Group**.</span><span class="sxs-lookup"><span data-stu-id="85321-248">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="85321-249">Os elementos **Control** contidos neste ponto de extensão precisam ter o atributo **xsi:type** definido como `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="85321-249">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="85321-250">Exemplo</span><span class="sxs-lookup"><span data-stu-id="85321-250">Example</span></span>

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

### <a name="mobileonlinemeetingcommandsurface-preview"></a><span data-ttu-id="85321-251">MobileOnlineMeetingCommandSurface (visualização)</span><span class="sxs-lookup"><span data-stu-id="85321-251">MobileOnlineMeetingCommandSurface (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="85321-252">Este ponto de extensão só tem suporte na [Visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Android com uma assinatura do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="85321-252">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="85321-253">Este ponto de extensão coloca uma alternância apropriada de modo na superfície de comando para um compromisso no fator de forma móvel.</span><span class="sxs-lookup"><span data-stu-id="85321-253">This extension point puts a mode-appropriate toggle in the command surface for an appointment in the mobile form factor.</span></span> <span data-ttu-id="85321-254">Um organizador da reunião pode criar uma reunião online.</span><span class="sxs-lookup"><span data-stu-id="85321-254">A meeting organizer can create an online meeting.</span></span> <span data-ttu-id="85321-255">Um participante pode ingressar na reunião online subsequentemente.</span><span class="sxs-lookup"><span data-stu-id="85321-255">An attendee can subsequently join the online meeting.</span></span> <span data-ttu-id="85321-256">Para saber mais sobre esse cenário, confira o artigo [criar um suplemento do Outlook Mobile para um provedor de reunião online](../../outlook/online-meeting.md) .</span><span class="sxs-lookup"><span data-stu-id="85321-256">To learn more about this scenario, see the [Create an Outlook mobile add-in for an online-meeting provider](../../outlook/online-meeting.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="85321-257">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="85321-257">Child elements</span></span>

|  <span data-ttu-id="85321-258">Elemento</span><span class="sxs-lookup"><span data-stu-id="85321-258">Element</span></span> |  <span data-ttu-id="85321-259">Descrição</span><span class="sxs-lookup"><span data-stu-id="85321-259">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="85321-260">Control</span><span class="sxs-lookup"><span data-stu-id="85321-260">Control</span></span>](control.md) |  <span data-ttu-id="85321-261">Adiciona um botão à superfície de comando.</span><span class="sxs-lookup"><span data-stu-id="85321-261">Adds a button to the command surface.</span></span>  |

<span data-ttu-id="85321-262">`ExtensionPoint`elementos desse tipo só podem ter um elemento filho: um `Control` elemento.</span><span class="sxs-lookup"><span data-stu-id="85321-262">`ExtensionPoint` elements of this type can only have one child element: a `Control` element.</span></span>

<span data-ttu-id="85321-263">O `Control` elemento contido neste ponto de extensão deve ter o `xsi:type` atributo definido como `MobileButton` .</span><span class="sxs-lookup"><span data-stu-id="85321-263">The `Control` element contained in this extension point must have the `xsi:type` attribute set to `MobileButton`.</span></span>

<span data-ttu-id="85321-264">As `Icon` imagens devem estar em escala de cinza usando `#919191` o código hex ou seu equivalente em [outros formatos de cor](https://convertingcolors.com/hex-color-919191.html).</span><span class="sxs-lookup"><span data-stu-id="85321-264">The `Icon` images should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>

#### <a name="example"></a><span data-ttu-id="85321-265">Exemplo</span><span class="sxs-lookup"><span data-stu-id="85321-265">Example</span></span>

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

### <a name="launchevent-preview"></a><span data-ttu-id="85321-266">LaunchEvent (visualização)</span><span class="sxs-lookup"><span data-stu-id="85321-266">LaunchEvent (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="85321-267">Este ponto de extensão só tem suporte na [Visualização](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) no Outlook na Web com uma assinatura do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="85321-267">This extension point is only supported in [preview](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="85321-268">Este ponto de extensão permite que um suplemento seja ativado com base nos eventos suportados no fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="85321-268">This extension point enables an add-in to activate based on supported events in the desktop form factor.</span></span> <span data-ttu-id="85321-269">Atualmente, os únicos eventos com suporte são `OnNewMessageCompose` e `OnNewAppointmentOrganizer` .</span><span class="sxs-lookup"><span data-stu-id="85321-269">Currently, the only supported events are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> <span data-ttu-id="85321-270">Para saber mais sobre esse cenário, confira o artigo [Configurar o suplemento do Outlook para ativação baseada em eventos](../../outlook/autolaunch.md) .</span><span class="sxs-lookup"><span data-stu-id="85321-270">To learn more about this scenario, see the [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md) article.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="85321-271">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="85321-271">Child elements</span></span>

|  <span data-ttu-id="85321-272">Elemento</span><span class="sxs-lookup"><span data-stu-id="85321-272">Element</span></span> |  <span data-ttu-id="85321-273">Descrição</span><span class="sxs-lookup"><span data-stu-id="85321-273">Description</span></span>  |
|:-----|:-----|
| [<span data-ttu-id="85321-274">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="85321-274">LaunchEvents</span></span>](launchevents.md) |  <span data-ttu-id="85321-275">Lista de [LaunchEvent](launchevent.md) para a ativação baseada em evento.</span><span class="sxs-lookup"><span data-stu-id="85321-275">List of [LaunchEvent](launchevent.md) for event-based activation.</span></span>  |
| [<span data-ttu-id="85321-276">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="85321-276">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="85321-277">O local do arquivo JavaScript de origem.</span><span class="sxs-lookup"><span data-stu-id="85321-277">The location of the source JavaScript file.</span></span>  |

#### <a name="example"></a><span data-ttu-id="85321-278">Exemplo</span><span class="sxs-lookup"><span data-stu-id="85321-278">Example</span></span>

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

### <a name="events"></a><span data-ttu-id="85321-279">Eventos</span><span class="sxs-lookup"><span data-stu-id="85321-279">Events</span></span>

<span data-ttu-id="85321-280">Este ponto de extensão adiciona um manipulador de eventos para um evento especificado.</span><span class="sxs-lookup"><span data-stu-id="85321-280">This extension point adds an event handler for a specified event.</span></span> <span data-ttu-id="85321-281">Para obter mais informações sobre como usar esse ponto de extensão, consulte o [recurso ao enviar para suplementos do Outlook](../../outlook/outlook-on-send-addins.md).</span><span class="sxs-lookup"><span data-stu-id="85321-281">For more information about using this extension point, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

| <span data-ttu-id="85321-282">Elemento</span><span class="sxs-lookup"><span data-stu-id="85321-282">Element</span></span> | <span data-ttu-id="85321-283">Descrição</span><span class="sxs-lookup"><span data-stu-id="85321-283">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="85321-284">Event</span><span class="sxs-lookup"><span data-stu-id="85321-284">Event</span></span>](event.md) |  <span data-ttu-id="85321-285">Especifica o evento e a função de manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="85321-285">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="85321-286">Exemplo do evento ItemSend</span><span class="sxs-lookup"><span data-stu-id="85321-286">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events">
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
</ExtensionPoint>
```

### <a name="detectedentity"></a><span data-ttu-id="85321-287">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="85321-287">DetectedEntity</span></span>

<span data-ttu-id="85321-288">Este ponto extensão adiciona uma ativação do suplemento contextual em um tipo de entidade especificada.</span><span class="sxs-lookup"><span data-stu-id="85321-288">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="85321-289">O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="85321-289">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="85321-290">Este tipo de elemento está disponível para [ clientes do Outlook que ofereçam suporte a conjuntos de requisitos 1.6 e posteriores](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="85321-290">This element type is available to [Outlook clients that support requirement sets 1.6 and later](../requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span></span>

|  <span data-ttu-id="85321-291">Elemento</span><span class="sxs-lookup"><span data-stu-id="85321-291">Element</span></span> |  <span data-ttu-id="85321-292">Descrição</span><span class="sxs-lookup"><span data-stu-id="85321-292">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="85321-293">Label</span><span class="sxs-lookup"><span data-stu-id="85321-293">Label</span></span>](#label) |  <span data-ttu-id="85321-294">Especifica o rótulo para o suplemento na janela contextual.</span><span class="sxs-lookup"><span data-stu-id="85321-294">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="85321-295">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="85321-295">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="85321-296">Especifica a URL para a janela contextual.</span><span class="sxs-lookup"><span data-stu-id="85321-296">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="85321-297">Rule</span><span class="sxs-lookup"><span data-stu-id="85321-297">Rule</span></span>](rule.md) |  <span data-ttu-id="85321-298">Especifica a regra ou regras que determinam quando um suplemento é ativado.</span><span class="sxs-lookup"><span data-stu-id="85321-298">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="85321-299">Label</span><span class="sxs-lookup"><span data-stu-id="85321-299">Label</span></span>

<span data-ttu-id="85321-300">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="85321-300">Required.</span></span> <span data-ttu-id="85321-301">O rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="85321-301">The label of the group.</span></span> <span data-ttu-id="85321-302">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="85321-302">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="85321-303">Requisitos de realce</span><span class="sxs-lookup"><span data-stu-id="85321-303">Highlight requirements</span></span>

<span data-ttu-id="85321-304">The only way a user can activate a contextual add-in is to interact with a highlighted entity.</span><span class="sxs-lookup"><span data-stu-id="85321-304">The only way a user can activate a contextual add-in is to interact with a highlighted entity.</span></span> <span data-ttu-id="85321-305">Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span><span class="sxs-lookup"><span data-stu-id="85321-305">Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="85321-306">However, there are some limitations to be aware of.</span><span class="sxs-lookup"><span data-stu-id="85321-306">However, there are some limitations to be aware of.</span></span> <span data-ttu-id="85321-307">These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span><span class="sxs-lookup"><span data-stu-id="85321-307">These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="85321-308">Os tipos de entidade `EmailAddress` e `Url` não podem ser realçados e, portanto, não podem ser usados para ativar um suplemento.</span><span class="sxs-lookup"><span data-stu-id="85321-308">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="85321-309">Se for usada uma única regra, `Highlight` DEVERÁ ser definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="85321-309">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="85321-310">Se usar um tipo de regra `RuleCollection` com `Mode="AND"` para combinar várias regras, pelo menos uma das regras DEVERÁ ter o `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="85321-310">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="85321-311">Se usar um tipo de regra `RuleCollection` com `Mode="OR"` para combinar várias regras, todas as regras DEVERÃO ter o `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="85321-311">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="85321-312">Exemplo do evento DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="85321-312">DetectedEntity event example</span></span>

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
