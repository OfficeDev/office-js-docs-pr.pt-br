---
title: Elemento Control no arquivo de manifesto
description: Define a função JavaScript que executa e aciona ou inicia um painel de tarefas.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 0d13fc4fd30f32b4998d385058cff941d3f72353
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720670"
---
# <a name="control-element"></a><span data-ttu-id="f37eb-103">Elemento Control</span><span class="sxs-lookup"><span data-stu-id="f37eb-103">Control element</span></span>

<span data-ttu-id="f37eb-p101">Define a função JavaScript que executa e aciona ou inicia um painel de tarefas. Um elemento **Control** pode ser um botão ou um menu. Pelo menos um **Control** deve ser incluído em um elemento [Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="f37eb-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="f37eb-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="f37eb-107">Attributes</span></span>

|  <span data-ttu-id="f37eb-108">Atributo</span><span class="sxs-lookup"><span data-stu-id="f37eb-108">Attribute</span></span>  |  <span data-ttu-id="f37eb-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f37eb-109">Required</span></span>  |  <span data-ttu-id="f37eb-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="f37eb-110">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="f37eb-111">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="f37eb-111">**xsi:type**</span></span>|<span data-ttu-id="f37eb-112">Sim</span><span class="sxs-lookup"><span data-stu-id="f37eb-112">Yes</span></span>|<span data-ttu-id="f37eb-p102">O tipo de controle que está sendo definido. Pode ser `Button`, `Menu` ou `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="f37eb-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="f37eb-115">**id**</span><span class="sxs-lookup"><span data-stu-id="f37eb-115">**id**</span></span>|<span data-ttu-id="f37eb-116">Não</span><span class="sxs-lookup"><span data-stu-id="f37eb-116">No</span></span>|<span data-ttu-id="f37eb-p103">A ID do elemento Control. Pode ter no máximo 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="f37eb-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="f37eb-119">O valor `MobileButton` de **xsi:type** é definido no esquema VersionOverrides 1.1.</span><span class="sxs-lookup"><span data-stu-id="f37eb-119">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="f37eb-120">Ele só se aplica aos elementos **Control** contidos em um elemento [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="f37eb-120">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="f37eb-121">Button control</span><span class="sxs-lookup"><span data-stu-id="f37eb-121">Button control</span></span>

<span data-ttu-id="f37eb-p105">Um botão executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função ou a exibição de um painel de tarefas. Cada controle de botão deve ter um único `id` para o manifesto.</span><span class="sxs-lookup"><span data-stu-id="f37eb-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="f37eb-125">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="f37eb-125">Child elements</span></span>
|  <span data-ttu-id="f37eb-126">Elemento</span><span class="sxs-lookup"><span data-stu-id="f37eb-126">Element</span></span> |  <span data-ttu-id="f37eb-127">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f37eb-127">Required</span></span>  |  <span data-ttu-id="f37eb-128">Descrição</span><span class="sxs-lookup"><span data-stu-id="f37eb-128">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f37eb-129">**Label**</span><span class="sxs-lookup"><span data-stu-id="f37eb-129">**Label**</span></span>     | <span data-ttu-id="f37eb-130">Sim</span><span class="sxs-lookup"><span data-stu-id="f37eb-130">Yes</span></span> |  <span data-ttu-id="f37eb-131">O texto do botão.</span><span class="sxs-lookup"><span data-stu-id="f37eb-131">The text for the button.</span></span> <span data-ttu-id="f37eb-132">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="f37eb-132">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="f37eb-133">**Emergente**</span><span class="sxs-lookup"><span data-stu-id="f37eb-133">**ToolTip**</span></span>    |<span data-ttu-id="f37eb-134">Não</span><span class="sxs-lookup"><span data-stu-id="f37eb-134">No</span></span>|<span data-ttu-id="f37eb-135">A dica de ferramenta do botão.</span><span class="sxs-lookup"><span data-stu-id="f37eb-135">The tooltip for the button.</span></span> <span data-ttu-id="f37eb-136">O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**.</span><span class="sxs-lookup"><span data-stu-id="f37eb-136">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="f37eb-137">O elemento **String** é um elemento filho do elemento **LongStrings**, que é filho do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="f37eb-137">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="f37eb-138">Supertip</span><span class="sxs-lookup"><span data-stu-id="f37eb-138">Supertip</span></span>](supertip.md)  | <span data-ttu-id="f37eb-139">Sim</span><span class="sxs-lookup"><span data-stu-id="f37eb-139">Yes</span></span> |  <span data-ttu-id="f37eb-140">A dica detalhada do botão.</span><span class="sxs-lookup"><span data-stu-id="f37eb-140">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="f37eb-141">Icon</span><span class="sxs-lookup"><span data-stu-id="f37eb-141">Icon</span></span>](icon.md)      | <span data-ttu-id="f37eb-142">Sim</span><span class="sxs-lookup"><span data-stu-id="f37eb-142">Yes</span></span> |  <span data-ttu-id="f37eb-143">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="f37eb-143">An image for the button.</span></span>         |
|  [<span data-ttu-id="f37eb-144">Action</span><span class="sxs-lookup"><span data-stu-id="f37eb-144">Action</span></span>](action.md)    | <span data-ttu-id="f37eb-145">Sim</span><span class="sxs-lookup"><span data-stu-id="f37eb-145">Yes</span></span> |  <span data-ttu-id="f37eb-146">Especifica a ação a ser executada.</span><span class="sxs-lookup"><span data-stu-id="f37eb-146">Specifies the action to perform.</span></span>  |
|  [<span data-ttu-id="f37eb-147">Enabled</span><span class="sxs-lookup"><span data-stu-id="f37eb-147">Enabled</span></span>](enabled.md)    | <span data-ttu-id="f37eb-148">Não</span><span class="sxs-lookup"><span data-stu-id="f37eb-148">No</span></span> |  <span data-ttu-id="f37eb-149">Especifica se o controle está habilitado quando o suplemento é iniciado.</span><span class="sxs-lookup"><span data-stu-id="f37eb-149">Specifies whether the control is enabled when the add-in launches.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="f37eb-150">Exemplo do botão ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="f37eb-150">ExecuteFunction button example</span></span>

<span data-ttu-id="f37eb-151">No exemplo a seguir, o botão é desabilitado quando o suplemento é iniciado.</span><span class="sxs-lookup"><span data-stu-id="f37eb-151">In the following example, the button is disabled when the add-in launches.</span></span> <span data-ttu-id="f37eb-152">Ele pode ser habilitado programaticamente.</span><span class="sxs-lookup"><span data-stu-id="f37eb-152">It can be programmatically enabled.</span></span> <span data-ttu-id="f37eb-153">Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="f37eb-153">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

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
  <Enabled>false</Enabled>
</Control>
```

### <a name="showtaskpane-button-example"></a><span data-ttu-id="f37eb-154">Exemplo do botão ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="f37eb-154">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="f37eb-155">Controles de menu (botão suspenso)</span><span class="sxs-lookup"><span data-stu-id="f37eb-155">Menu (dropdown button) controls</span></span>

<span data-ttu-id="f37eb-p109">Um menu define uma lista estática de opções. Cada item de menu executa uma função ou mostra um painel de tarefas. Não há suporte para submenus.</span><span class="sxs-lookup"><span data-stu-id="f37eb-p109">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="f37eb-159">Quando usado com um **ponto de extensão\*\*\*\*PrimaryCommandSurface** ou [ContextMenu](extensionpoint.md), o controle de menu define:</span><span class="sxs-lookup"><span data-stu-id="f37eb-159">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="f37eb-160">Um item de menu no nível raiz.</span><span class="sxs-lookup"><span data-stu-id="f37eb-160">A root-level menu item.</span></span>

- <span data-ttu-id="f37eb-161">Uma lista de itens de submenu.</span><span class="sxs-lookup"><span data-stu-id="f37eb-161">A list of submenu items.</span></span>

<span data-ttu-id="f37eb-p110">Quando usado com **PrimaryCommandSurface**, o item de menu raiz é exibido como um botão na faixa de opções. Quando o botão é selecionado, o submenu é exibido como uma lista suspensa. Quando usado com **ContextMenu**, um item de menu com um submenu é inserido no menu de contexto. Em ambos os casos, cada item de submenu pode executar uma função JavaScript ou mostrar um painel de tarefas. Somente um nível de submenus é compatível no momento.</span><span class="sxs-lookup"><span data-stu-id="f37eb-p110">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="f37eb-p111">O exemplo a seguir mostra como definir um item de menu com dois itens de submenu. O primeiro item do submenu mostra um painel de tarefas e o segundo item executa uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f37eb-p111">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="f37eb-169">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="f37eb-169">Child elements</span></span>

|  <span data-ttu-id="f37eb-170">Elemento</span><span class="sxs-lookup"><span data-stu-id="f37eb-170">Element</span></span> |  <span data-ttu-id="f37eb-171">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f37eb-171">Required</span></span>  |  <span data-ttu-id="f37eb-172">Descrição</span><span class="sxs-lookup"><span data-stu-id="f37eb-172">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f37eb-173">**Label**</span><span class="sxs-lookup"><span data-stu-id="f37eb-173">**Label**</span></span>     | <span data-ttu-id="f37eb-174">Sim</span><span class="sxs-lookup"><span data-stu-id="f37eb-174">Yes</span></span> |  <span data-ttu-id="f37eb-175">O texto do botão.</span><span class="sxs-lookup"><span data-stu-id="f37eb-175">The text for the button.</span></span> <span data-ttu-id="f37eb-176">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="f37eb-176">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="f37eb-177">**Emergente**</span><span class="sxs-lookup"><span data-stu-id="f37eb-177">**ToolTip**</span></span>    |<span data-ttu-id="f37eb-178">Não</span><span class="sxs-lookup"><span data-stu-id="f37eb-178">No</span></span>|<span data-ttu-id="f37eb-179">A dica de ferramenta do botão.</span><span class="sxs-lookup"><span data-stu-id="f37eb-179">The tooltip for the button.</span></span> <span data-ttu-id="f37eb-180">O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**.</span><span class="sxs-lookup"><span data-stu-id="f37eb-180">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="f37eb-181">O elemento **String** é um elemento filho do elemento **LongStrings**, que é filho do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="f37eb-181">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="f37eb-182">Supertip</span><span class="sxs-lookup"><span data-stu-id="f37eb-182">Supertip</span></span>](supertip.md)  | <span data-ttu-id="f37eb-183">Sim</span><span class="sxs-lookup"><span data-stu-id="f37eb-183">Yes</span></span> |  <span data-ttu-id="f37eb-184">A dica detalhada do botão.</span><span class="sxs-lookup"><span data-stu-id="f37eb-184">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="f37eb-185">Icon</span><span class="sxs-lookup"><span data-stu-id="f37eb-185">Icon</span></span>](icon.md)      | <span data-ttu-id="f37eb-186">Sim</span><span class="sxs-lookup"><span data-stu-id="f37eb-186">Yes</span></span> |  <span data-ttu-id="f37eb-187">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="f37eb-187">An image for the button.</span></span>         |
|  <span data-ttu-id="f37eb-188">**Items**</span><span class="sxs-lookup"><span data-stu-id="f37eb-188">**Items**</span></span>     | <span data-ttu-id="f37eb-189">Sim</span><span class="sxs-lookup"><span data-stu-id="f37eb-189">Yes</span></span> |  <span data-ttu-id="f37eb-190">Um conjunto de botões a exibir dentro do menu.</span><span class="sxs-lookup"><span data-stu-id="f37eb-190">A collection of Buttons to display within the menu.</span></span> <span data-ttu-id="f37eb-191">Contém os elementos **Item** para cada item do submenu.</span><span class="sxs-lookup"><span data-stu-id="f37eb-191">Contains the **Item** elements for each submenu item.</span></span> <span data-ttu-id="f37eb-192">Cada elemento **Item** contém os mesmos elementos filhos que [Button control](#button-control).</span><span class="sxs-lookup"><span data-stu-id="f37eb-192">Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="f37eb-193">Exemplo de controle de menu</span><span class="sxs-lookup"><span data-stu-id="f37eb-193">Menu control examples</span></span>

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

## <a name="mobilebutton-control"></a><span data-ttu-id="f37eb-194">Controle MobileButton</span><span class="sxs-lookup"><span data-stu-id="f37eb-194">MobileButton control</span></span>

<span data-ttu-id="f37eb-p115">Um botão móvel executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função ou a exibição de um painel de tarefas. Cada controle de botão móvel deve ter um único `id` para o manifesto.</span><span class="sxs-lookup"><span data-stu-id="f37eb-p115">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="f37eb-p116">O valor `MobileButton` de **xsi:type** é definido no esquema VersionOverrides 1.1. O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="f37eb-p116">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="f37eb-200">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="f37eb-200">Child elements</span></span>
|  <span data-ttu-id="f37eb-201">Elemento</span><span class="sxs-lookup"><span data-stu-id="f37eb-201">Element</span></span> |  <span data-ttu-id="f37eb-202">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="f37eb-202">Required</span></span>  |  <span data-ttu-id="f37eb-203">Descrição</span><span class="sxs-lookup"><span data-stu-id="f37eb-203">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="f37eb-204">**Label**</span><span class="sxs-lookup"><span data-stu-id="f37eb-204">**Label**</span></span>     | <span data-ttu-id="f37eb-205">Sim</span><span class="sxs-lookup"><span data-stu-id="f37eb-205">Yes</span></span> |  <span data-ttu-id="f37eb-206">O texto do botão.</span><span class="sxs-lookup"><span data-stu-id="f37eb-206">The text for the button.</span></span> <span data-ttu-id="f37eb-207">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="f37eb-207">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="f37eb-208">Icon</span><span class="sxs-lookup"><span data-stu-id="f37eb-208">Icon</span></span>](icon.md)      | <span data-ttu-id="f37eb-209">Sim</span><span class="sxs-lookup"><span data-stu-id="f37eb-209">Yes</span></span> |  <span data-ttu-id="f37eb-210">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="f37eb-210">An image for the button.</span></span>         |
|  [<span data-ttu-id="f37eb-211">Action</span><span class="sxs-lookup"><span data-stu-id="f37eb-211">Action</span></span>](action.md)    | <span data-ttu-id="f37eb-212">Sim</span><span class="sxs-lookup"><span data-stu-id="f37eb-212">Yes</span></span> |  <span data-ttu-id="f37eb-213">Especifica a ação a ser executada.</span><span class="sxs-lookup"><span data-stu-id="f37eb-213">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="f37eb-214">Exemplo de botão móvel ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="f37eb-214">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="f37eb-215">Exemplo de botão móvel ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="f37eb-215">ShowTaskpane mobile button example</span></span>

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
