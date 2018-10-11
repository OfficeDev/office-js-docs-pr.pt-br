# <a name="control-element"></a><span data-ttu-id="81d0f-101">Elemento Control</span><span class="sxs-lookup"><span data-stu-id="81d0f-101">Control element</span></span>

<span data-ttu-id="81d0f-p101">Define a função JavaScript que executa e aciona ou inicia um painel de tarefas. Um elemento **Control** pode ser um botão ou um menu. Pelo menos um **Control** deve ser incluído em um elemento [Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="81d0f-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="81d0f-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="81d0f-105">Attributes</span></span>

|  <span data-ttu-id="81d0f-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="81d0f-106">Attribute</span></span>  |  <span data-ttu-id="81d0f-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="81d0f-107">Required</span></span>  |  <span data-ttu-id="81d0f-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="81d0f-108">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="81d0f-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="81d0f-109">**xsi:type**</span></span>|<span data-ttu-id="81d0f-110">Sim</span><span class="sxs-lookup"><span data-stu-id="81d0f-110">Yes</span></span>|<span data-ttu-id="81d0f-p102">O tipo de controle que está sendo definido. Pode ser `Button`, `Menu` ou `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="81d0f-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="81d0f-113">**id**</span><span class="sxs-lookup"><span data-stu-id="81d0f-113">**id**</span></span>|<span data-ttu-id="81d0f-114">Não</span><span class="sxs-lookup"><span data-stu-id="81d0f-114">No</span></span>|<span data-ttu-id="81d0f-p103">A ID do elemento Control. Pode ter no máximo 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="81d0f-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="81d0f-117">O `MobileButton` valor para **tipo xsi** é definido no esquema de VersionOverrides 1.1.</span><span class="sxs-lookup"><span data-stu-id="81d0f-117">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing VersionOverrides element must have an  attribute value of .</span></span> <span data-ttu-id="81d0f-118">Ele se aplica somente aos elementos **Control** contidos em um elemento [MobileFormFactor](mobileformfactor.md) .</span><span class="sxs-lookup"><span data-stu-id="81d0f-118">Note: The  value for xsi:type is defined in VersionOverrides schema 1.1. It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="81d0f-119">Controle de botão</span><span class="sxs-lookup"><span data-stu-id="81d0f-119">Button control</span></span>

<span data-ttu-id="81d0f-p105">Um botão executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função ou a exibição de um painel de tarefas. Cada controle de botão deve ter um único `id` para o manifesto.</span><span class="sxs-lookup"><span data-stu-id="81d0f-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="81d0f-123">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="81d0f-123">Child elements</span></span>
|  <span data-ttu-id="81d0f-124">Elemento</span><span class="sxs-lookup"><span data-stu-id="81d0f-124">Element</span></span> |  <span data-ttu-id="81d0f-125">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="81d0f-125">Required</span></span>  |  <span data-ttu-id="81d0f-126">Descrição</span><span class="sxs-lookup"><span data-stu-id="81d0f-126">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="81d0f-127">**Label**</span><span class="sxs-lookup"><span data-stu-id="81d0f-127">**Label**</span></span>     | <span data-ttu-id="81d0f-128">Sim</span><span class="sxs-lookup"><span data-stu-id="81d0f-128">Yes</span></span> |  <span data-ttu-id="81d0f-p106">O texto do botão. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="81d0f-p106">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="81d0f-131">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="81d0f-131">**ToolTip**</span></span>  |<span data-ttu-id="81d0f-132">Não</span><span class="sxs-lookup"><span data-stu-id="81d0f-132">No</span></span>|<span data-ttu-id="81d0f-p107">A dica de ferramenta do botão. O atributo **resid** deve ser atribuído com o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **LongStrings**, que é filho do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="81d0f-p107">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="81d0f-136">Superdica</span><span class="sxs-lookup"><span data-stu-id="81d0f-136">Supertip</span></span>](supertip.md)  | <span data-ttu-id="81d0f-137">Sim</span><span class="sxs-lookup"><span data-stu-id="81d0f-137">Yes</span></span> |  <span data-ttu-id="81d0f-138">A dica detalhada do botão.</span><span class="sxs-lookup"><span data-stu-id="81d0f-138">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="81d0f-139">Icon</span><span class="sxs-lookup"><span data-stu-id="81d0f-139">Icon</span></span>](icon.md)      | <span data-ttu-id="81d0f-140">Sim</span><span class="sxs-lookup"><span data-stu-id="81d0f-140">Yes</span></span> |  <span data-ttu-id="81d0f-141">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="81d0f-141">An image for the button.</span></span>         |
|  [<span data-ttu-id="81d0f-142">Ação</span><span class="sxs-lookup"><span data-stu-id="81d0f-142">Action</span></span>](action.md)    | <span data-ttu-id="81d0f-143">Sim</span><span class="sxs-lookup"><span data-stu-id="81d0f-143">Yes</span></span> |  <span data-ttu-id="81d0f-144">Especifica a ação a realizar.</span><span class="sxs-lookup"><span data-stu-id="81d0f-144">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="81d0f-145">Exemplo do botão ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="81d0f-145">ExecuteFunction button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="81d0f-146">Exemplo do botão ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="81d0f-146">ShowTaskpane button example</span></span>

```xml
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="81d0f-147">Controles de menu (botão suspenso)</span><span class="sxs-lookup"><span data-stu-id="81d0f-147">Menu (dropdown button) controls</span></span>

<span data-ttu-id="81d0f-p108">Um menu define uma lista estática de opções. Cada item de menu executa uma função ou mostra um painel de tarefas. Não há suporte para submenus.</span><span class="sxs-lookup"><span data-stu-id="81d0f-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="81d0f-151">Quando usado com um **ponto de extensão** **PrimaryCommandSurface** ou [ContextMenu](extensionpoint.md) , o controle de menu define:</span><span class="sxs-lookup"><span data-stu-id="81d0f-151">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="81d0f-152">Um item de menu no nível raiz.</span><span class="sxs-lookup"><span data-stu-id="81d0f-152">A root-level menu item.</span></span>

- <span data-ttu-id="81d0f-153">Uma lista de itens de submenu.</span><span class="sxs-lookup"><span data-stu-id="81d0f-153">A list of submenu items.</span></span>

<span data-ttu-id="81d0f-p109">Quando usado com **PrimaryCommandSurface**, o item de menu raiz é exibido como um botão na faixa de opções. Quando o botão é selecionado, o submenu é exibido como uma lista suspensa. Quando usado com **ContextMenu**, um item de menu com um submenu é inserido no menu de contexto. Em ambos os casos, cada item de submenu pode executar uma função JavaScript ou mostrar um painel de tarefas. Somente há suporte para um nível de submenus no momento.</span><span class="sxs-lookup"><span data-stu-id="81d0f-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="81d0f-p110">O exemplo a seguir mostra como definir um item de menu com dois itens de submenu. O primeiro item do submenu mostra um painel de tarefas e o segundo item executa uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="81d0f-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="81d0f-161">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="81d0f-161">Child elements</span></span>

|  <span data-ttu-id="81d0f-162">Elemento</span><span class="sxs-lookup"><span data-stu-id="81d0f-162">Element</span></span> |  <span data-ttu-id="81d0f-163">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="81d0f-163">Required</span></span>  |  <span data-ttu-id="81d0f-164">Descrição</span><span class="sxs-lookup"><span data-stu-id="81d0f-164">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="81d0f-165">**Label**</span><span class="sxs-lookup"><span data-stu-id="81d0f-165">**Label**</span></span>     | <span data-ttu-id="81d0f-166">Sim</span><span class="sxs-lookup"><span data-stu-id="81d0f-166">Yes</span></span> |  <span data-ttu-id="81d0f-p111">O texto do botão. O atributo **resid** deve ser atribuído com o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="81d0f-p111">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="81d0f-169">**Dica de ferramenta**</span><span class="sxs-lookup"><span data-stu-id="81d0f-169">**ToolTip**</span></span>  |<span data-ttu-id="81d0f-170">Não</span><span class="sxs-lookup"><span data-stu-id="81d0f-170">No</span></span>|<span data-ttu-id="81d0f-p112">A dica de ferramenta do botão. O atributo **resid** deve ser atribuído com o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **LongStrings**, que é filho do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="81d0f-p112">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="81d0f-174">Superdica</span><span class="sxs-lookup"><span data-stu-id="81d0f-174">Supertip</span></span>](supertip.md)  | <span data-ttu-id="81d0f-175">Sim</span><span class="sxs-lookup"><span data-stu-id="81d0f-175">Yes</span></span> |  <span data-ttu-id="81d0f-176">A dica detalhada do botão.</span><span class="sxs-lookup"><span data-stu-id="81d0f-176">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="81d0f-177">Ícone</span><span class="sxs-lookup"><span data-stu-id="81d0f-177">Icon</span></span>](icon.md)      | <span data-ttu-id="81d0f-178">Sim</span><span class="sxs-lookup"><span data-stu-id="81d0f-178">Yes</span></span> |  <span data-ttu-id="81d0f-179">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="81d0f-179">An image for the button.</span></span>         |
|  <span data-ttu-id="81d0f-180">**Items**</span><span class="sxs-lookup"><span data-stu-id="81d0f-180">**Items**</span></span>     | <span data-ttu-id="81d0f-181">Sim</span><span class="sxs-lookup"><span data-stu-id="81d0f-181">Yes</span></span> |  <span data-ttu-id="81d0f-p113">Um conjunto de botões a exibir dentro do menu. Contém os elementos **Item** para cada item do submenu. Cada elemento **Item** contém os elementos filhos do [Button control](#button-control).</span><span class="sxs-lookup"><span data-stu-id="81d0f-p113">A collection of Buttons to display within the menu. Contains the  **Item** elements for each submenu item. Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="81d0f-185">Exemplo de controle de menu</span><span class="sxs-lookup"><span data-stu-id="81d0f-185">Menu control examples</span></span>

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

```xml
<Control xsi:type="Menu" id="msgReadMenuButton">
  <Label resid="menuReadButtonLabel" />
  <Supertip>
    <Title resid="menuReadSuperTipTitle" />
    <Description resid="menuReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="red-icon-16" />
    <bt:Image size="32" resid="red-icon-32" />
    <bt:Image size="80" resid="red-icon-80" />
  </Icon>
  <Items>
    <Item id="msgReadMenuItem1">
      <Label resid="menuItem1ReadLabel" />
      <Supertip>
        <Title resid="menuItem1ReadLabel" />
        <Description resid="menuItem1ReadTip" />
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="red-icon-16" />
        <bt:Image size="32" resid="red-icon-32" />
        <bt:Image size="80" resid="red-icon-80" />
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>getItemClass</FunctionName>
      </Action>
    </Item>
  </Items>
</Control>
```

## <a name="mobilebutton-control"></a><span data-ttu-id="81d0f-186">Controle MobileButton</span><span class="sxs-lookup"><span data-stu-id="81d0f-186">MobileButton control</span></span>

<span data-ttu-id="81d0f-p114">Um botão móvel executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função ou a exibição de um painel de tarefas. Cada controle de botão móvel deve ter um único `id` para o manifesto.</span><span class="sxs-lookup"><span data-stu-id="81d0f-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="81d0f-p115">O valor `MobileButton` para **xsi:type** é definido no esquema VersionOverrides 1.1. O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um  `xsi:type` valor do atributo  de `VersionOverridesV1_1` .</span><span class="sxs-lookup"><span data-stu-id="81d0f-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="81d0f-192">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="81d0f-192">Child elements</span></span>
|  <span data-ttu-id="81d0f-193">Elemento</span><span class="sxs-lookup"><span data-stu-id="81d0f-193">Element</span></span> |  <span data-ttu-id="81d0f-194">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="81d0f-194">Required</span></span>  |  <span data-ttu-id="81d0f-195">Descrição</span><span class="sxs-lookup"><span data-stu-id="81d0f-195">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="81d0f-196">**Label**</span><span class="sxs-lookup"><span data-stu-id="81d0f-196">**Label**</span></span>     | <span data-ttu-id="81d0f-197">Sim</span><span class="sxs-lookup"><span data-stu-id="81d0f-197">Yes</span></span> |  <span data-ttu-id="81d0f-p116">O texto do botão. O atributo **resid** deve ser atribuído com o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="81d0f-p116">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="81d0f-200">Ícone</span><span class="sxs-lookup"><span data-stu-id="81d0f-200">Icon</span></span>](icon.md)      | <span data-ttu-id="81d0f-201">Sim</span><span class="sxs-lookup"><span data-stu-id="81d0f-201">Yes</span></span> |  <span data-ttu-id="81d0f-202">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="81d0f-202">An image for the button.</span></span>         |
|  [<span data-ttu-id="81d0f-203">Action</span><span class="sxs-lookup"><span data-stu-id="81d0f-203">Action</span></span>](action.md)    | <span data-ttu-id="81d0f-204">Sim</span><span class="sxs-lookup"><span data-stu-id="81d0f-204">Yes</span></span> |  <span data-ttu-id="81d0f-205">Especifica a ação a realizar.</span><span class="sxs-lookup"><span data-stu-id="81d0f-205">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="81d0f-206">Exemplo de botão móvel ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="81d0f-206">ExecuteFunction mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="81d0f-207">Exemplo de botão móvel ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="81d0f-207">ShowTaskpane mobile button example</span></span>

```xml
<Control xsi:type="MobileButton" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```