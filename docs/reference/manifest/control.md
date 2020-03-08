---
title: Elemento Control no arquivo de manifesto
description: ''
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: ccf7c3065db13a311825498292713b619f1cd745
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/07/2020
ms.locfileid: "42562004"
---
# <a name="control-element"></a><span data-ttu-id="3e405-102">Elemento Control</span><span class="sxs-lookup"><span data-stu-id="3e405-102">Control element</span></span>

<span data-ttu-id="3e405-p101">Define a função JavaScript que executa e aciona ou inicia um painel de tarefas. Um elemento **Control** pode ser um botão ou um menu. Pelo menos um **Control** deve ser incluído em um elemento [Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="3e405-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="3e405-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="3e405-106">Attributes</span></span>

|  <span data-ttu-id="3e405-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="3e405-107">Attribute</span></span>  |  <span data-ttu-id="3e405-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3e405-108">Required</span></span>  |  <span data-ttu-id="3e405-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="3e405-109">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="3e405-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="3e405-110">**xsi:type**</span></span>|<span data-ttu-id="3e405-111">Sim</span><span class="sxs-lookup"><span data-stu-id="3e405-111">Yes</span></span>|<span data-ttu-id="3e405-p102">O tipo de controle que está sendo definido. Pode ser `Button`, `Menu` ou `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="3e405-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="3e405-114">**id**</span><span class="sxs-lookup"><span data-stu-id="3e405-114">**id**</span></span>|<span data-ttu-id="3e405-115">Não</span><span class="sxs-lookup"><span data-stu-id="3e405-115">No</span></span>|<span data-ttu-id="3e405-p103">A ID do elemento Control. Pode ter no máximo 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="3e405-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="3e405-118">O valor `MobileButton` de **xsi:type** é definido no esquema VersionOverrides 1.1.</span><span class="sxs-lookup"><span data-stu-id="3e405-118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="3e405-119">Ele só se aplica aos elementos **Control** contidos em um elemento [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="3e405-119">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="3e405-120">Button control</span><span class="sxs-lookup"><span data-stu-id="3e405-120">Button control</span></span>

<span data-ttu-id="3e405-p105">Um botão executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função ou a exibição de um painel de tarefas. Cada controle de botão deve ter um único `id` para o manifesto.</span><span class="sxs-lookup"><span data-stu-id="3e405-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="3e405-124">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3e405-124">Child elements</span></span>
|  <span data-ttu-id="3e405-125">Elemento</span><span class="sxs-lookup"><span data-stu-id="3e405-125">Element</span></span> |  <span data-ttu-id="3e405-126">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3e405-126">Required</span></span>  |  <span data-ttu-id="3e405-127">Descrição</span><span class="sxs-lookup"><span data-stu-id="3e405-127">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3e405-128">**Label**</span><span class="sxs-lookup"><span data-stu-id="3e405-128">**Label**</span></span>     | <span data-ttu-id="3e405-129">Sim</span><span class="sxs-lookup"><span data-stu-id="3e405-129">Yes</span></span> |  <span data-ttu-id="3e405-130">O texto do botão.</span><span class="sxs-lookup"><span data-stu-id="3e405-130">The text for the button.</span></span> <span data-ttu-id="3e405-131">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="3e405-131">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="3e405-132">**Emergente**</span><span class="sxs-lookup"><span data-stu-id="3e405-132">**ToolTip**</span></span>  |<span data-ttu-id="3e405-133">Não</span><span class="sxs-lookup"><span data-stu-id="3e405-133">No</span></span>|<span data-ttu-id="3e405-134">A dica de ferramenta do botão.</span><span class="sxs-lookup"><span data-stu-id="3e405-134">The tooltip for the button.</span></span> <span data-ttu-id="3e405-135">O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**.</span><span class="sxs-lookup"><span data-stu-id="3e405-135">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="3e405-136">O elemento **String** é um elemento filho do elemento **LongStrings**, que é filho do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="3e405-136">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|     
|  [<span data-ttu-id="3e405-137">Supertip</span><span class="sxs-lookup"><span data-stu-id="3e405-137">Supertip</span></span>](supertip.md)  | <span data-ttu-id="3e405-138">Sim</span><span class="sxs-lookup"><span data-stu-id="3e405-138">Yes</span></span> |  <span data-ttu-id="3e405-139">A dica detalhada do botão.</span><span class="sxs-lookup"><span data-stu-id="3e405-139">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="3e405-140">Icon</span><span class="sxs-lookup"><span data-stu-id="3e405-140">Icon</span></span>](icon.md)      | <span data-ttu-id="3e405-141">Sim</span><span class="sxs-lookup"><span data-stu-id="3e405-141">Yes</span></span> |  <span data-ttu-id="3e405-142">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="3e405-142">An image for the button.</span></span>         |
|  [<span data-ttu-id="3e405-143">Action</span><span class="sxs-lookup"><span data-stu-id="3e405-143">Action</span></span>](action.md)    | <span data-ttu-id="3e405-144">Sim</span><span class="sxs-lookup"><span data-stu-id="3e405-144">Yes</span></span> |  <span data-ttu-id="3e405-145">Especifica a ação a ser executada.</span><span class="sxs-lookup"><span data-stu-id="3e405-145">Specifies the action to perform.</span></span>  |
|  [<span data-ttu-id="3e405-146">Enabled</span><span class="sxs-lookup"><span data-stu-id="3e405-146">Enabled</span></span>](enabled.md)    | <span data-ttu-id="3e405-147">Não</span><span class="sxs-lookup"><span data-stu-id="3e405-147">No</span></span> |  <span data-ttu-id="3e405-148">Especifica se o controle está habilitado quando o suplemento é iniciado.</span><span class="sxs-lookup"><span data-stu-id="3e405-148">Specifies whether the control is enabled when the add-in launches.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="3e405-149">Exemplo do botão ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="3e405-149">ExecuteFunction button example</span></span>

<span data-ttu-id="3e405-150">No exemplo a seguir, o botão é desabilitado quando o suplemento é iniciado.</span><span class="sxs-lookup"><span data-stu-id="3e405-150">In the following example, the button is disabled when the add-in launches.</span></span> <span data-ttu-id="3e405-151">Ele pode ser habilitado programaticamente.</span><span class="sxs-lookup"><span data-stu-id="3e405-151">It can be programmatically enabled.</span></span> <span data-ttu-id="3e405-152">Para obter mais informações, consulte [habilitar e desabilitar comandos de suplemento](/office/dev/add-ins/design/disable-add-in-commands).</span><span class="sxs-lookup"><span data-stu-id="3e405-152">For more information, see [Enable and Disable Add-in Commands](/office/dev/add-ins/design/disable-add-in-commands).</span></span>

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

### <a name="showtaskpane-button-example"></a><span data-ttu-id="3e405-153">Exemplo do botão ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="3e405-153">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="3e405-154">Controles de menu (botão suspenso)</span><span class="sxs-lookup"><span data-stu-id="3e405-154">Menu (dropdown button) controls</span></span>

<span data-ttu-id="3e405-p109">Um menu define uma lista estática de opções. Cada item de menu executa uma função ou mostra um painel de tarefas. Não há suporte para submenus.</span><span class="sxs-lookup"><span data-stu-id="3e405-p109">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="3e405-158">Quando usado com um **ponto de extensão\*\*\*\*PrimaryCommandSurface** ou [ContextMenu](extensionpoint.md), o controle de menu define:</span><span class="sxs-lookup"><span data-stu-id="3e405-158">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="3e405-159">Um item de menu no nível raiz.</span><span class="sxs-lookup"><span data-stu-id="3e405-159">A root-level menu item.</span></span>

- <span data-ttu-id="3e405-160">Uma lista de itens de submenu.</span><span class="sxs-lookup"><span data-stu-id="3e405-160">A list of submenu items.</span></span>

<span data-ttu-id="3e405-p110">Quando usado com **PrimaryCommandSurface**, o item de menu raiz é exibido como um botão na faixa de opções. Quando o botão é selecionado, o submenu é exibido como uma lista suspensa. Quando usado com **ContextMenu**, um item de menu com um submenu é inserido no menu de contexto. Em ambos os casos, cada item de submenu pode executar uma função JavaScript ou mostrar um painel de tarefas. Somente um nível de submenus é compatível no momento.</span><span class="sxs-lookup"><span data-stu-id="3e405-p110">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="3e405-p111">O exemplo a seguir mostra como definir um item de menu com dois itens de submenu. O primeiro item do submenu mostra um painel de tarefas e o segundo item executa uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3e405-p111">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="3e405-168">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3e405-168">Child elements</span></span>

|  <span data-ttu-id="3e405-169">Elemento</span><span class="sxs-lookup"><span data-stu-id="3e405-169">Element</span></span> |  <span data-ttu-id="3e405-170">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3e405-170">Required</span></span>  |  <span data-ttu-id="3e405-171">Descrição</span><span class="sxs-lookup"><span data-stu-id="3e405-171">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3e405-172">**Label**</span><span class="sxs-lookup"><span data-stu-id="3e405-172">**Label**</span></span>     | <span data-ttu-id="3e405-173">Sim</span><span class="sxs-lookup"><span data-stu-id="3e405-173">Yes</span></span> |  <span data-ttu-id="3e405-174">O texto do botão.</span><span class="sxs-lookup"><span data-stu-id="3e405-174">The text for the button.</span></span> <span data-ttu-id="3e405-175">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="3e405-175">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="3e405-176">**Emergente**</span><span class="sxs-lookup"><span data-stu-id="3e405-176">**ToolTip**</span></span>  |<span data-ttu-id="3e405-177">Não</span><span class="sxs-lookup"><span data-stu-id="3e405-177">No</span></span>|<span data-ttu-id="3e405-178">A dica de ferramenta do botão.</span><span class="sxs-lookup"><span data-stu-id="3e405-178">The tooltip for the button.</span></span> <span data-ttu-id="3e405-179">O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**.</span><span class="sxs-lookup"><span data-stu-id="3e405-179">The **resid** attribute must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="3e405-180">O elemento **String** é um elemento filho do elemento **LongStrings**, que é filho do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="3e405-180">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|     
|  [<span data-ttu-id="3e405-181">Supertip</span><span class="sxs-lookup"><span data-stu-id="3e405-181">Supertip</span></span>](supertip.md)  | <span data-ttu-id="3e405-182">Sim</span><span class="sxs-lookup"><span data-stu-id="3e405-182">Yes</span></span> |  <span data-ttu-id="3e405-183">A dica detalhada do botão.</span><span class="sxs-lookup"><span data-stu-id="3e405-183">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="3e405-184">Icon</span><span class="sxs-lookup"><span data-stu-id="3e405-184">Icon</span></span>](icon.md)      | <span data-ttu-id="3e405-185">Sim</span><span class="sxs-lookup"><span data-stu-id="3e405-185">Yes</span></span> |  <span data-ttu-id="3e405-186">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="3e405-186">An image for the button.</span></span>         |
|  <span data-ttu-id="3e405-187">**Items**</span><span class="sxs-lookup"><span data-stu-id="3e405-187">**Items**</span></span>     | <span data-ttu-id="3e405-188">Sim</span><span class="sxs-lookup"><span data-stu-id="3e405-188">Yes</span></span> |  <span data-ttu-id="3e405-189">Um conjunto de botões a exibir dentro do menu.</span><span class="sxs-lookup"><span data-stu-id="3e405-189">A collection of Buttons to display within the menu.</span></span> <span data-ttu-id="3e405-190">Contém os elementos **Item** para cada item do submenu.</span><span class="sxs-lookup"><span data-stu-id="3e405-190">Contains the **Item** elements for each submenu item.</span></span> <span data-ttu-id="3e405-191">Cada elemento **Item** contém os mesmos elementos filhos que [Button control](#button-control).</span><span class="sxs-lookup"><span data-stu-id="3e405-191">Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="3e405-192">Exemplo de controle de menu</span><span class="sxs-lookup"><span data-stu-id="3e405-192">Menu control examples</span></span>

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

## <a name="mobilebutton-control"></a><span data-ttu-id="3e405-193">Controle MobileButton</span><span class="sxs-lookup"><span data-stu-id="3e405-193">MobileButton control</span></span>

<span data-ttu-id="3e405-p115">Um botão móvel executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função ou a exibição de um painel de tarefas. Cada controle de botão móvel deve ter um único `id` para o manifesto.</span><span class="sxs-lookup"><span data-stu-id="3e405-p115">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="3e405-p116">O valor `MobileButton` de **xsi:type** é definido no esquema VersionOverrides 1.1. O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="3e405-p116">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="3e405-199">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3e405-199">Child elements</span></span>
|  <span data-ttu-id="3e405-200">Elemento</span><span class="sxs-lookup"><span data-stu-id="3e405-200">Element</span></span> |  <span data-ttu-id="3e405-201">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3e405-201">Required</span></span>  |  <span data-ttu-id="3e405-202">Descrição</span><span class="sxs-lookup"><span data-stu-id="3e405-202">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="3e405-203">**Label**</span><span class="sxs-lookup"><span data-stu-id="3e405-203">**Label**</span></span>     | <span data-ttu-id="3e405-204">Sim</span><span class="sxs-lookup"><span data-stu-id="3e405-204">Yes</span></span> |  <span data-ttu-id="3e405-205">O texto do botão.</span><span class="sxs-lookup"><span data-stu-id="3e405-205">The text for the button.</span></span> <span data-ttu-id="3e405-206">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="3e405-206">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="3e405-207">Icon</span><span class="sxs-lookup"><span data-stu-id="3e405-207">Icon</span></span>](icon.md)      | <span data-ttu-id="3e405-208">Sim</span><span class="sxs-lookup"><span data-stu-id="3e405-208">Yes</span></span> |  <span data-ttu-id="3e405-209">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="3e405-209">An image for the button.</span></span>         |
|  [<span data-ttu-id="3e405-210">Action</span><span class="sxs-lookup"><span data-stu-id="3e405-210">Action</span></span>](action.md)    | <span data-ttu-id="3e405-211">Sim</span><span class="sxs-lookup"><span data-stu-id="3e405-211">Yes</span></span> |  <span data-ttu-id="3e405-212">Especifica a ação a ser executada.</span><span class="sxs-lookup"><span data-stu-id="3e405-212">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="3e405-213">Exemplo de botão móvel ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="3e405-213">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="3e405-214">Exemplo de botão móvel ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="3e405-214">ShowTaskpane mobile button example</span></span>

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
