# <a name="extensionpoint-element"></a><span data-ttu-id="ee05d-101">Elemento ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="ee05d-101">ExtensionPoint element</span></span>

 <span data-ttu-id="ee05d-102">Define onde um suplemento expõe a funcionalidade na interface de usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="ee05d-102">Defines where an add-in exposes functionality in the Office UI.</span></span> <span data-ttu-id="ee05d-103">O elemento **ExtensionPoint** é um elemento filho de [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="ee05d-103">The **ExtensionPoint** element is a child element of [AllFormFactors](allformfactors.md), [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> 

## <a name="attributes"></a><span data-ttu-id="ee05d-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="ee05d-104">Attributes</span></span>

|  <span data-ttu-id="ee05d-105">Atributo</span><span class="sxs-lookup"><span data-stu-id="ee05d-105">Attribute</span></span>  |  <span data-ttu-id="ee05d-106">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ee05d-106">Required</span></span>  |  <span data-ttu-id="ee05d-107">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee05d-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ee05d-108">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="ee05d-108">**xsi:type**</span></span>  |  <span data-ttu-id="ee05d-109">Sim</span><span class="sxs-lookup"><span data-stu-id="ee05d-109">Yes</span></span>  | <span data-ttu-id="ee05d-110">O tipo de ponto de extensão que está sendo definido.</span><span class="sxs-lookup"><span data-stu-id="ee05d-110">The type of extension point being defined.</span></span>|

## <a name="extension-points-for-excel-only"></a><span data-ttu-id="ee05d-111">Pontos de extensão somente para Excel</span><span class="sxs-lookup"><span data-stu-id="ee05d-111">Extension points for Excel only</span></span>

- <span data-ttu-id="ee05d-112">**CustomFunctions** - uma função personalizada escrita em JavaScript para Excel.</span><span class="sxs-lookup"><span data-stu-id="ee05d-112">**CustomFunctions** - a custom function written in JavaScript for Excel.</span></span>

<span data-ttu-id="ee05d-113">[Esse exemplo de código XML](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.xml) mostra como usar o elemento **ExtensionPoint** com o valor de atributo **CustomFunctions** e quais elementos filhos devem ser usados.</span><span class="sxs-lookup"><span data-stu-id="ee05d-113">The following examples show how to use the ExtensionPoint element with the CustomFunctions attribute value, and the child elements to be used.</span></span>

## <a name="extension-points-for-word-excel-powerpoint-and-onenote-add-in-commands"></a><span data-ttu-id="ee05d-114">Pontos de extensão para comandos de suplemento do Word, Excel, PowerPoint e OneNote</span><span class="sxs-lookup"><span data-stu-id="ee05d-114">Extension points for Word, Excel, PowerPoint, and OneNote add-in commands</span></span>

- <span data-ttu-id="ee05d-115">**PrimaryCommandSurface** - a faixa de opções no Office.</span><span class="sxs-lookup"><span data-stu-id="ee05d-115">**PrimaryCommandSurface** - The ribbon in Office.</span></span>
- <span data-ttu-id="ee05d-116">**ContextMenu** - o menu de atalho exibido ao clicar com o botão direito do mouse na interface de usuário do Office.</span><span class="sxs-lookup"><span data-stu-id="ee05d-116">**ContextMenu** - The shortcut menu that appears when you right-click in the Office UI.</span></span>

<span data-ttu-id="ee05d-117">Os exemplos a seguir mostram como usar o elemento **ExtensionPoint** com os valores de atributo **PrimaryCommandSurface** e **ContextMenu** e os elementos filhos que devem ser usados com cada um.</span><span class="sxs-lookup"><span data-stu-id="ee05d-117">The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="ee05d-118">Forneça uma ID exclusiva para os elementos que contêm um atributo ID.</span><span class="sxs-lookup"><span data-stu-id="ee05d-118">IMPORTANT  For elements that contain an ID attribute, make sure you provide a unique ID.</span></span> <span data-ttu-id="ee05d-119">É recomendável usar o nome da sua empresa com a ID.</span><span class="sxs-lookup"><span data-stu-id="ee05d-119">We recommend that you use your company's name along with your ID.</span></span> <span data-ttu-id="ee05d-120">Por exemplo, use o formato a seguir.</span><span class="sxs-lookup"><span data-stu-id="ee05d-120">For example, use the following format.</span></span> <CustomTab id="mycompanyname.mygroupname">

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

#### <a name="child-elements"></a><span data-ttu-id="ee05d-121">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ee05d-121">Child elements</span></span>
 
|<span data-ttu-id="ee05d-122">**Elemento**</span><span class="sxs-lookup"><span data-stu-id="ee05d-122">**Element**</span></span>|<span data-ttu-id="ee05d-123">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="ee05d-123">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="ee05d-124">**CustomTab**</span><span class="sxs-lookup"><span data-stu-id="ee05d-124">**CustomTab**</span></span>|<span data-ttu-id="ee05d-p103">Obrigatório se você quiser adicionar uma guia personalizada à faixa de opções (usando **PrimaryCommandSurface**). Se você usar o elemento **CustomTab** , não será possível usar o elemento **OfficeTab**. O atributo **id** é obrigatório.</span><span class="sxs-lookup"><span data-stu-id="ee05d-p103">Required if you want to add a custom tab to the ribbon (using  **PrimaryCommandSurface**). If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.</span></span>|
|<span data-ttu-id="ee05d-128">**OfficeTab**</span><span class="sxs-lookup"><span data-stu-id="ee05d-128">**OfficeTab**</span></span>|<span data-ttu-id="ee05d-p104">Obrigatório se você quiser estender uma guia de faixa de opções padrão do Office (usando **PrimaryCommandSurface**). Se você usar o elemento **OfficeTab** , não poderá usar o elemento **CustomTab**. Para saber mais, confira [OfficeTab](officetab.md).</span><span class="sxs-lookup"><span data-stu-id="ee05d-p104">Required if you want to extend a default Office ribbon tab (using **PrimaryCommandSurface**). If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).</span></span>|
|<span data-ttu-id="ee05d-132">**OfficeMenu**</span><span class="sxs-lookup"><span data-stu-id="ee05d-132">**OfficeMenu**</span></span>|<span data-ttu-id="ee05d-p105">Obrigatório se você estiver adicionando comandos de suplemento a um menu de contexto padrão (usando **ContextMenu**). O atributo **id** deve ser definido como: </span><span class="sxs-lookup"><span data-stu-id="ee05d-p105">Required if you're adding add-in commands to a default context menu (using  **ContextMenu**). The  **id** attribute must be set to: </span></span><br/> <span data-ttu-id="ee05d-p106">- **ContextMenuText** para o Excel ou Word. Exibe o item no menu de contexto quando o texto for selecionado e o usuário clicar com o botão direito do mouse no texto selecionado. </span><span class="sxs-lookup"><span data-stu-id="ee05d-p106">- **ContextMenuText** for Excel or Word. Displays the item on the context menu when text is selected and then the user right-clicks on the selected text. </span></span><br/> <span data-ttu-id="ee05d-p107">- **ContextMenuCell** para Excel. Exibe o item no menu de contexto quando o usuário clica com o botão direito do mouse em uma célula na planilha.</span><span class="sxs-lookup"><span data-stu-id="ee05d-p107">- **ContextMenuCell** for Excel. Displays the  item on the context menu when the user right-clicks on a cell on the spreadsheet.</span></span>|
|<span data-ttu-id="ee05d-139">**Grupo**</span><span class="sxs-lookup"><span data-stu-id="ee05d-139">**Group**</span></span>|<span data-ttu-id="ee05d-p108">Um grupo de pontos de extensão de interface do usuário em uma guia. O grupo pode ter até seis controles. O atributo **id** é obrigatório. É uma sequência de caracteres com, no máximo, 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="ee05d-p108">A group of user interface extension points on a tab. A group can have up to six controls. The  **id** attribute is required. It's a string with a maximum of 125 characters.</span></span>|
|<span data-ttu-id="ee05d-143">**Rótulo**</span><span class="sxs-lookup"><span data-stu-id="ee05d-143">**Label**</span></span>|<span data-ttu-id="ee05d-p109">Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **ShortStrings**, que é elemento filho do elemento **Resources**.</span><span class="sxs-lookup"><span data-stu-id="ee05d-p109">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **ShortStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="ee05d-148">**Ícone**</span><span class="sxs-lookup"><span data-stu-id="ee05d-148">**Icon**</span></span>|<span data-ttu-id="ee05d-p110">Obrigatório. Especifica o ícone do grupo a ser usado em dispositivos de fator forma pequeno, ou quando muitos botões forem exibidos. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **Image**. O elemento **Image** é elemento filho do elemento **Images**, que é elemento filho do elemento **Resources**. O atributo **size** fornece o tamanho da imagem em pixels. Três tamanhos de imagem são obrigatórios: 16, 32 e 80 pixels. Também há suporte para cinco tamanhos opcionais: 20, 24, 40, 48 e 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="ee05d-p110">Required. Specifies the group's icon to be used on small form factor devices, or when too many buttons are displayed. The  **resid** attribute must be set to the value of the **id** attribute of an **Image** element. The **Image** element is a child element of the **Images** element, which is a child element of the **Resources** element. The **size** attribute gives the size, in pixels, of the image. Three image sizes are required: 16, 32, and 80. Five optional sizes are also supported: 20, 24, 40, 48, and 64.</span></span>|
|<span data-ttu-id="ee05d-156">**Dica de ferramenta**</span><span class="sxs-lookup"><span data-stu-id="ee05d-156">**Tooltip**</span></span>|<span data-ttu-id="ee05d-p111">Opcional. A dica de ferramenta do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **LongStrings**, que, por sua vez, é um elemento filho do elemento **Resources**.</span><span class="sxs-lookup"><span data-stu-id="ee05d-p111">Optional. The tooltip of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child element of the **LongStrings** element, which is a child element of the **Resources** element.</span></span>|
|<span data-ttu-id="ee05d-161">**Controle**</span><span class="sxs-lookup"><span data-stu-id="ee05d-161">**Control**</span></span>|<span data-ttu-id="ee05d-162">Cada grupo exige pelo menos um controle.</span><span class="sxs-lookup"><span data-stu-id="ee05d-162">Each group requires at least one control.</span></span> <span data-ttu-id="ee05d-163">Um elemento **Control** pode ser um **Button** ou um **Menu**.</span><span class="sxs-lookup"><span data-stu-id="ee05d-163">A  **Control** element can be either a **Button** or a **Menu**.</span></span> <span data-ttu-id="ee05d-164">Use **Menu** para especificar uma lista suspensa de controles de botão.</span><span class="sxs-lookup"><span data-stu-id="ee05d-164">Use  **Menu** to specify a drop-down list of button controls.</span></span> <span data-ttu-id="ee05d-165">Atualmente, há suporte apenas para botões e menus.</span><span class="sxs-lookup"><span data-stu-id="ee05d-165">Currently, only buttons and menus are supported.</span></span> <span data-ttu-id="ee05d-166">Confira as seções [Controles de botão](control.md#button-control) e [Controles de menu](control.md#menu-dropdown-button-controls) para saber mais.</span><span class="sxs-lookup"><span data-stu-id="ee05d-166">See the [Button controls](control.md#button-control) and [Menu controls](control.md#menu-dropdown-button-controls) sections for more information.</span></span><br/><span data-ttu-id="ee05d-167">**Observação:** para facilitar a solução de problemas, é recomendável que um elemento **Control** e os elementos filho **Resources** associados sejam adicionados um de cada vez.</span><span class="sxs-lookup"><span data-stu-id="ee05d-167">**Note**  To make troubleshooting easier, we recommend that a  **Control** element and the related **Resources** child elements be added one at a time.</span></span>|
|<span data-ttu-id="ee05d-168">**Script**</span><span class="sxs-lookup"><span data-stu-id="ee05d-168">**Script**</span></span>|<span data-ttu-id="ee05d-169">Links para o arquivo JavaScript com a definição de função personalizada e o código de registro</span><span class="sxs-lookup"><span data-stu-id="ee05d-169">Links to the JavaScript file with the custom function definition and registration code.</span></span> <span data-ttu-id="ee05d-170">Esse elemento não é usado na Visualização do Desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="ee05d-170">This element is not used in the Developer Preview.</span></span> <span data-ttu-id="ee05d-171">Em vez disso, a página HTML é responsável por carregar todos os arquivos JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ee05d-171">Instead, the HTML page is responsible for loading all JavaScript files.</span></span>|
|<span data-ttu-id="ee05d-172">**Página**</span><span class="sxs-lookup"><span data-stu-id="ee05d-172">**Page**</span></span>|<span data-ttu-id="ee05d-173">Links para a página HTML de suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ee05d-173">Links to the HTML page for your custom functions.</span></span>|

## <a name="extension-points-for-outlook"></a><span data-ttu-id="ee05d-174">Pontos de extensão do Outlook</span><span class="sxs-lookup"><span data-stu-id="ee05d-174">Extension points for Outlook add-in commands</span></span>

- [<span data-ttu-id="ee05d-175">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ee05d-175">MessageReadCommandSurface</span></span>](#messagereadcommandsurface) 
- [<span data-ttu-id="ee05d-176">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ee05d-176">MessageComposeCommandSurface</span></span>](#messagecomposecommandsurface) 
- [<span data-ttu-id="ee05d-177">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ee05d-177">AppointmentOrganizerCommandSurface</span></span>](#appointmentorganizercommandsurface) 
- [<span data-ttu-id="ee05d-178">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ee05d-178">AppointmentAttendeeCommandSurface</span></span>](#appointmentattendeecommandsurface)
- <span data-ttu-id="ee05d-179">[Module](#module) (só pode ser usado em [DesktopFormFactor](desktopformfactor.md)).</span><span class="sxs-lookup"><span data-stu-id="ee05d-179">[Module](#module) (Can only be used in the [DesktopFormFactor](desktopformfactor.md).)</span></span>
- [<span data-ttu-id="ee05d-180">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ee05d-180">MobileMessageReadCommandSurface</span></span>](#mobilemessagereadcommandsurface)
- [<span data-ttu-id="ee05d-181">Events</span><span class="sxs-lookup"><span data-stu-id="ee05d-181">Events</span></span>](#events)
- [<span data-ttu-id="ee05d-182">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="ee05d-182">DetectedEntity</span></span>](#detectedentity)

### <a name="messagereadcommandsurface"></a><span data-ttu-id="ee05d-183">MessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ee05d-183">MessageReadCommandSurface</span></span>
<span data-ttu-id="ee05d-p114">Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de e-mail. No Outlook para área de trabalho, isso aparece na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="ee05d-p114">This extension point puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="ee05d-186">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ee05d-186">Child elements</span></span>

|  <span data-ttu-id="ee05d-187">Elemento</span><span class="sxs-lookup"><span data-stu-id="ee05d-187">Element</span></span> |  <span data-ttu-id="ee05d-188">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee05d-188">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ee05d-189">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-189">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="ee05d-190">Adiciona os comandos à guia de faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="ee05d-190">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="ee05d-191">CustomTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-191">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="ee05d-192">Adiciona os comandos à guia de faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="ee05d-192">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="ee05d-193">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-193">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="ee05d-194">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-194">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="messagecomposecommandsurface"></a><span data-ttu-id="ee05d-195">MessageComposeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ee05d-195">MessageComposeCommandSurface</span></span>
<span data-ttu-id="ee05d-196">Este ponto de extensão coloca botões na faixa de opções para suplementos que usam o formulário de composição de e-mail.</span><span class="sxs-lookup"><span data-stu-id="ee05d-196">This extension point puts buttons on the ribbon for add-ins using mail compose form.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="ee05d-197">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ee05d-197">Child elements</span></span>

|  <span data-ttu-id="ee05d-198">Elemento</span><span class="sxs-lookup"><span data-stu-id="ee05d-198">Element</span></span> |  <span data-ttu-id="ee05d-199">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee05d-199">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ee05d-200">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-200">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="ee05d-201">Adiciona os comandos à guia de faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="ee05d-201">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="ee05d-202">CustomTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-202">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="ee05d-203">Adiciona os comandos à guia de faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="ee05d-203">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="ee05d-204">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-204">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="ee05d-205">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-205">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentorganizercommandsurface"></a><span data-ttu-id="ee05d-206">AppointmentOrganizerCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ee05d-206">AppointmentOrganizerCommandSurface</span></span>

<span data-ttu-id="ee05d-207">Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao organizador da reunião.</span><span class="sxs-lookup"><span data-stu-id="ee05d-207">This extension point puts buttons on the ribbon for the form that's displayed to the organizer of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="ee05d-208">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ee05d-208">Child elements</span></span>

|  <span data-ttu-id="ee05d-209">Elemento</span><span class="sxs-lookup"><span data-stu-id="ee05d-209">Element</span></span> |  <span data-ttu-id="ee05d-210">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee05d-210">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ee05d-211">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-211">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="ee05d-212">Adiciona os comandos à guia de faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="ee05d-212">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="ee05d-213">CustomTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-213">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="ee05d-214">Adiciona os comandos à guia de faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="ee05d-214">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="ee05d-215">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-215">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="ee05d-216">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-216">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="appointmentattendeecommandsurface"></a><span data-ttu-id="ee05d-217">AppointmentAttendeeCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ee05d-217">AppointmentAttendeeCommandSurface</span></span>

<span data-ttu-id="ee05d-218">Este ponto de extensão coloca botões na faixa de opções para o formulário exibido ao participante da reunião.</span><span class="sxs-lookup"><span data-stu-id="ee05d-218">This extension point puts buttons on the ribbon for the form that's displayed to the attendee of the meeting.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="ee05d-219">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ee05d-219">Child elements</span></span>

|  <span data-ttu-id="ee05d-220">Elemento</span><span class="sxs-lookup"><span data-stu-id="ee05d-220">Element</span></span> |  <span data-ttu-id="ee05d-221">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee05d-221">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ee05d-222">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-222">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="ee05d-223">Adiciona os comandos à guia de faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="ee05d-223">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="ee05d-224">CustomTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-224">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="ee05d-225">Adiciona os comandos à guia de faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="ee05d-225">Adds the command(s) to the custom ribbon tab.</span></span>  |

#### <a name="officetab-example"></a><span data-ttu-id="ee05d-226">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-226">OfficeTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <OfficeTab id="TabDefault">
        <-- OfficeTab Definition -->
  </OfficeTab>
</ExtensionPoint>
```

#### <a name="customtab-example"></a><span data-ttu-id="ee05d-227">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-227">CustomTab example</span></span>
```xml
<ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
  <CustomTab id="TabCustom1">
        <-- CustomTab Definition -->
  </CustomTab>
</ExtensionPoint>
```

### <a name="module"></a><span data-ttu-id="ee05d-228">Módulo</span><span class="sxs-lookup"><span data-stu-id="ee05d-228">Module</span></span>

<span data-ttu-id="ee05d-229">Este ponto de extensão coloca botões na faixa de opções para a extensão do módulo.</span><span class="sxs-lookup"><span data-stu-id="ee05d-229">This extension point puts buttons on the ribbon for the module extension.</span></span> 

#### <a name="child-elements"></a><span data-ttu-id="ee05d-230">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ee05d-230">Child elements</span></span>

|  <span data-ttu-id="ee05d-231">Elemento</span><span class="sxs-lookup"><span data-stu-id="ee05d-231">Element</span></span> |  <span data-ttu-id="ee05d-232">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee05d-232">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ee05d-233">OfficeTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-233">OfficeTab</span></span>](officetab.md) |  <span data-ttu-id="ee05d-234">Adiciona os comandos à guia de faixa de opções padrão.</span><span class="sxs-lookup"><span data-stu-id="ee05d-234">Adds the command(s) to the default ribbon tab.</span></span>  |
|  [<span data-ttu-id="ee05d-235">CustomTab</span><span class="sxs-lookup"><span data-stu-id="ee05d-235">CustomTab</span></span>](customtab.md) |  <span data-ttu-id="ee05d-236">Adiciona os comandos à guia de faixa de opções personalizada.</span><span class="sxs-lookup"><span data-stu-id="ee05d-236">Adds the command(s) to the custom ribbon tab.</span></span>  |

### <a name="mobilemessagereadcommandsurface"></a><span data-ttu-id="ee05d-237">MobileMessageReadCommandSurface</span><span class="sxs-lookup"><span data-stu-id="ee05d-237">MobileMessageReadCommandSurface</span></span>
<span data-ttu-id="ee05d-238">Este ponto de extensão coloca os botões na superfície de comando para o modo de exibição de leitura de e-mail no fator forma móvel.</span><span class="sxs-lookup"><span data-stu-id="ee05d-238">This extension point puts buttons in the command surface for the mail read view in the mobile form factor.</span></span>

#### <a name="child-elements"></a><span data-ttu-id="ee05d-239">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ee05d-239">Child elements</span></span>

|  <span data-ttu-id="ee05d-240">Elemento</span><span class="sxs-lookup"><span data-stu-id="ee05d-240">Element</span></span> |  <span data-ttu-id="ee05d-241">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee05d-241">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ee05d-242">Grupo</span><span class="sxs-lookup"><span data-stu-id="ee05d-242">Group</span></span>](group.md) |  <span data-ttu-id="ee05d-243">Adiciona um grupo de botões à superfície de comando.</span><span class="sxs-lookup"><span data-stu-id="ee05d-243">Adds a group of buttons to the command surface.</span></span>  |

<span data-ttu-id="ee05d-244">Os elementos **ExtensionPoint** desse tipo só podem ter um elemento filho: um elemento **Group**.</span><span class="sxs-lookup"><span data-stu-id="ee05d-244">**ExtensionPoint** elements of this type can only have one child element: a **Group** element.</span></span>

<span data-ttu-id="ee05d-245">Os elementos **Control** contidos neste ponto de extensão precisam ter o atributo **xsi:type** definido como `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="ee05d-245">**Control** elements contained in this extension point must have the **xsi:type** attribute set to `MobileButton`.</span></span>

#### <a name="example"></a><span data-ttu-id="ee05d-246">Exemplo</span><span class="sxs-lookup"><span data-stu-id="ee05d-246">Example</span></span>
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

### <a name="events"></a><span data-ttu-id="ee05d-247">Events</span><span class="sxs-lookup"><span data-stu-id="ee05d-247">Events</span></span>

<span data-ttu-id="ee05d-248">Esse ponto de extensão adiciona um manipulador de eventos para um evento especificado.</span><span class="sxs-lookup"><span data-stu-id="ee05d-248">This extension point adds an event handler for a specified event.</span></span>

> [!NOTE]
> <span data-ttu-id="ee05d-249">Este tipo de elemento só tem suporte pelo Outlook na Web no Office 365.</span><span class="sxs-lookup"><span data-stu-id="ee05d-249">Note: This element type is only supported by Outlook on the web in Office 365.</span></span>

| <span data-ttu-id="ee05d-250">Elemento</span><span class="sxs-lookup"><span data-stu-id="ee05d-250">Element</span></span> | <span data-ttu-id="ee05d-251">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee05d-251">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ee05d-252">Event</span><span class="sxs-lookup"><span data-stu-id="ee05d-252">Event</span></span>](event.md) |  <span data-ttu-id="ee05d-253">Especifica o evento e a função de manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="ee05d-253">Specifies the event and event handler function.</span></span>  |

#### <a name="itemsend-event-example"></a><span data-ttu-id="ee05d-254">Exemplo do evento ItemSend</span><span class="sxs-lookup"><span data-stu-id="ee05d-254">ItemSend event example</span></span>

```xml
<ExtensionPoint xsi:type="Events"> 
  <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
</ExtensionPoint> 
```

### <a name="detectedentity"></a><span data-ttu-id="ee05d-255">DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="ee05d-255">DetectedEntity</span></span>

<span data-ttu-id="ee05d-256">Este ponto de extensão adiciona uma ativação do suplemento contextual em um tipo de entidade especificada.</span><span class="sxs-lookup"><span data-stu-id="ee05d-256">This extension point adds a contextual add-in activation on a specified entity type.</span></span>

<span data-ttu-id="ee05d-257">O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor do atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="ee05d-257">The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

> [!NOTE]
> <span data-ttu-id="ee05d-258">Este tipo de elemento só tem suporte pelo Outlook na Web no Office 365.</span><span class="sxs-lookup"><span data-stu-id="ee05d-258">Note: This element type is only supported by Outlook on the web in Office 365.</span></span>

|  <span data-ttu-id="ee05d-259">Elemento</span><span class="sxs-lookup"><span data-stu-id="ee05d-259">Element</span></span> |  <span data-ttu-id="ee05d-260">Descrição</span><span class="sxs-lookup"><span data-stu-id="ee05d-260">Description</span></span>  |
|:-----|:-----|
|  [<span data-ttu-id="ee05d-261">Rótulo</span><span class="sxs-lookup"><span data-stu-id="ee05d-261">Label</span></span>](#label) |  <span data-ttu-id="ee05d-262">Especifica o rótulo para o suplemento na janela contextual.</span><span class="sxs-lookup"><span data-stu-id="ee05d-262">Specifies the label for the add-in in the contextual window.</span></span>  |
|  [<span data-ttu-id="ee05d-263">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="ee05d-263">SourceLocation</span></span>](sourcelocation.md) |  <span data-ttu-id="ee05d-264">Especifica o URL para a janela contextual.</span><span class="sxs-lookup"><span data-stu-id="ee05d-264">Specifies the URL for the contextual window.</span></span>  |
|  [<span data-ttu-id="ee05d-265">Rule</span><span class="sxs-lookup"><span data-stu-id="ee05d-265">Rule</span></span>](rule.md) |  <span data-ttu-id="ee05d-266">Especifica a regra ou regras que determinam quando um suplemento é ativado.</span><span class="sxs-lookup"><span data-stu-id="ee05d-266">Specifies the rule or rules that determine when an add-in activates.</span></span>  |

#### <a name="label"></a><span data-ttu-id="ee05d-267">Rótulo</span><span class="sxs-lookup"><span data-stu-id="ee05d-267">Label</span></span>

<span data-ttu-id="ee05d-p115">Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="ee05d-p115">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

#### <a name="highlight-requirements"></a><span data-ttu-id="ee05d-271">Requisitos de destaque</span><span class="sxs-lookup"><span data-stu-id="ee05d-271">Highlight requirements</span></span>

<span data-ttu-id="ee05d-p116">A única maneira que um usuário pode ativar um suplemento contextual é interagir com uma entidade destacada. Os desenvolvedores podem controlar quais entidades são destacadas usando o atributo `Highlight` do elemento `Rule` para os tipos de regra `ItemHasKnownEntity` e `ItemHasRegularExpressionMatch`.</span><span class="sxs-lookup"><span data-stu-id="ee05d-p116">The only way a user can activate a contextual add-in is to interact with a highlighted entity. Developers can control which entities are highlighted by using the `Highlight` attribute of the `Rule` element for `ItemHasKnownEntity` and `ItemHasRegularExpressionMatch` rule types.</span></span>

<span data-ttu-id="ee05d-p117">No entanto, há algumas limitações que devem ser consideradas. Essas limitações são para garantir que sempre haverá uma entidade destacada em compromissos ou mensagens aplicáveis para oferecer ao usuário uma maneira de ativar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="ee05d-p117">However, there are some limitations to be aware of. These limitations are in place to ensure that there will always be a highlighted entity in applicable messages or appointments to give the user a way to activate the add-in.</span></span>

- <span data-ttu-id="ee05d-276">Os tipos de entidade `EmailAddress` e `Url` não podem ser destacados e, portanto, não podem ser usados para ativar um suplemento.</span><span class="sxs-lookup"><span data-stu-id="ee05d-276">The `EmailAddress` and `Url` entity types cannot be highlighted, and therefore cannot be used to activate an add-in.</span></span>
- <span data-ttu-id="ee05d-277">Se for usada uma única regra, `Highlight` DEVERÁ ser definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="ee05d-277">If using a single rule, `Highlight` MUST be set to `all`.</span></span>
- <span data-ttu-id="ee05d-278">Se usar um tipo de regra `RuleCollection` com `Mode="AND"` para combinar várias regras, pelo menos uma das regras DEVERÁ ter o `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="ee05d-278">If using a `RuleCollection` rule type with `Mode="AND"` to combine multiple rules, at least one of the rules MUST have `Highlight` set to `all`.</span></span>
- <span data-ttu-id="ee05d-279">Se usar um tipo de regra `RuleCollection` com `Mode="OR"` para combinar várias regras, todas as regras DEVERÃO ter o `Highlight` definido como `all`.</span><span class="sxs-lookup"><span data-stu-id="ee05d-279">If using a `RuleCollection` rule type with `Mode="OR"` to combine multiple rules, all of the rules MUST have `Highlight` set to `all`.</span></span>

#### <a name="detectedentity-event-example"></a><span data-ttu-id="ee05d-280">Exemplo do evento DetectedEntity</span><span class="sxs-lookup"><span data-stu-id="ee05d-280">DetectedEntity event example</span></span>

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