---
title: Elemento Control no arquivo de manifesto
description: Define a função JavaScript que executa e aciona ou inicia um painel de tarefas.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 737902bef52edeb70e2c5760df5bb589b624271b
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173980"
---
# <a name="control-element"></a><span data-ttu-id="33642-103">Elemento Control</span><span class="sxs-lookup"><span data-stu-id="33642-103">Control element</span></span>

<span data-ttu-id="33642-p101">Define a função JavaScript que executa e aciona ou inicia um painel de tarefas. Um elemento **Control** pode ser um botão ou um menu. Pelo menos um **Control** deve ser incluído em um elemento [Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="33642-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="33642-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="33642-107">Attributes</span></span>

|  <span data-ttu-id="33642-108">Atributo</span><span class="sxs-lookup"><span data-stu-id="33642-108">Attribute</span></span>  |  <span data-ttu-id="33642-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="33642-109">Required</span></span>  |  <span data-ttu-id="33642-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="33642-110">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="33642-111">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="33642-111">**xsi:type**</span></span>|<span data-ttu-id="33642-112">Sim</span><span class="sxs-lookup"><span data-stu-id="33642-112">Yes</span></span>|<span data-ttu-id="33642-p102">O tipo de controle que está sendo definido. Pode ser `Button`, `Menu` ou `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="33642-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="33642-115">**id**</span><span class="sxs-lookup"><span data-stu-id="33642-115">**id**</span></span>|<span data-ttu-id="33642-116">Não</span><span class="sxs-lookup"><span data-stu-id="33642-116">No</span></span>|<span data-ttu-id="33642-p103">A ID do elemento Control. Pode ter no máximo 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="33642-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="33642-119">O valor `MobileButton` de **xsi:type** é definido no esquema VersionOverrides 1.1.</span><span class="sxs-lookup"><span data-stu-id="33642-119">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1.</span></span> <span data-ttu-id="33642-120">Ele só se aplica aos elementos **Control** contidos em um elemento [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="33642-120">It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="33642-121">Button control</span><span class="sxs-lookup"><span data-stu-id="33642-121">Button control</span></span>

<span data-ttu-id="33642-p105">Um botão executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função ou a exibição de um painel de tarefas. Cada controle de botão deve ter um único `id` para o manifesto.</span><span class="sxs-lookup"><span data-stu-id="33642-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="33642-125">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="33642-125">Child elements</span></span>
|  <span data-ttu-id="33642-126">Elemento</span><span class="sxs-lookup"><span data-stu-id="33642-126">Element</span></span> |  <span data-ttu-id="33642-127">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="33642-127">Required</span></span>  |  <span data-ttu-id="33642-128">Descrição</span><span class="sxs-lookup"><span data-stu-id="33642-128">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="33642-129">**Label**</span><span class="sxs-lookup"><span data-stu-id="33642-129">**Label**</span></span>     | <span data-ttu-id="33642-130">Sim</span><span class="sxs-lookup"><span data-stu-id="33642-130">Yes</span></span> |  <span data-ttu-id="33642-131">O texto do botão.</span><span class="sxs-lookup"><span data-stu-id="33642-131">The text for the button.</span></span> <span data-ttu-id="33642-132">O **atributo resid** pode ter no máximo 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no elemento [Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="33642-132">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="33642-133">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="33642-133">**ToolTip**</span></span>    |<span data-ttu-id="33642-134">Não</span><span class="sxs-lookup"><span data-stu-id="33642-134">No</span></span>|<span data-ttu-id="33642-135">A dica de ferramenta do botão.</span><span class="sxs-lookup"><span data-stu-id="33642-135">The tooltip for the button.</span></span> <span data-ttu-id="33642-136">O **atributo resid** pode ter no máximo 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String.**</span><span class="sxs-lookup"><span data-stu-id="33642-136">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="33642-137">O elemento **String** é um elemento filho do elemento **LongStrings**, que é filho do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="33642-137">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="33642-138">Supertip</span><span class="sxs-lookup"><span data-stu-id="33642-138">Supertip</span></span>](supertip.md)  | <span data-ttu-id="33642-139">Sim</span><span class="sxs-lookup"><span data-stu-id="33642-139">Yes</span></span> |  <span data-ttu-id="33642-140">A dica detalhada do botão.</span><span class="sxs-lookup"><span data-stu-id="33642-140">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="33642-141">Icon</span><span class="sxs-lookup"><span data-stu-id="33642-141">Icon</span></span>](icon.md)      | <span data-ttu-id="33642-142">Sim</span><span class="sxs-lookup"><span data-stu-id="33642-142">Yes</span></span> |  <span data-ttu-id="33642-143">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="33642-143">An image for the button.</span></span>         |
|  [<span data-ttu-id="33642-144">Action</span><span class="sxs-lookup"><span data-stu-id="33642-144">Action</span></span>](action.md)    | <span data-ttu-id="33642-145">Sim</span><span class="sxs-lookup"><span data-stu-id="33642-145">Yes</span></span> |  <span data-ttu-id="33642-146">Especifica a ação a ser executada.</span><span class="sxs-lookup"><span data-stu-id="33642-146">Specifies the action to perform.</span></span>  |
|  [<span data-ttu-id="33642-147">Enabled</span><span class="sxs-lookup"><span data-stu-id="33642-147">Enabled</span></span>](enabled.md)    | <span data-ttu-id="33642-148">Não</span><span class="sxs-lookup"><span data-stu-id="33642-148">No</span></span> |  <span data-ttu-id="33642-149">Especifica se o controle está habilitado quando o complemento é lançado.</span><span class="sxs-lookup"><span data-stu-id="33642-149">Specifies whether the control is enabled when the add-in launches.</span></span>  |
|  [<span data-ttu-id="33642-150">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="33642-150">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="33642-151">Não</span><span class="sxs-lookup"><span data-stu-id="33642-151">No</span></span> |  <span data-ttu-id="33642-152">Especifica se o botão deve aparecer em combinações de aplicativos e plataformas que suportam guias contextuais personalizadas.</span><span class="sxs-lookup"><span data-stu-id="33642-152">Specifies whether the button should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="33642-153">Se usado, ele deve ser o *primeiro* elemento filho.</span><span class="sxs-lookup"><span data-stu-id="33642-153">If used, it must be the *first* child element.</span></span> |

### <a name="executefunction-button-example"></a><span data-ttu-id="33642-154">Exemplo do botão ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="33642-154">ExecuteFunction button example</span></span>

<span data-ttu-id="33642-155">No exemplo a seguir, o botão é desabilitado quando o complemento é lançado.</span><span class="sxs-lookup"><span data-stu-id="33642-155">In the following example, the button is disabled when the add-in launches.</span></span> <span data-ttu-id="33642-156">Ela pode ser habilitada programaticamente.</span><span class="sxs-lookup"><span data-stu-id="33642-156">It can be programmatically enabled.</span></span> <span data-ttu-id="33642-157">Para obter mais informações, consulte [Ativar e Desativar Comandos de Suplemento](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="33642-157">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

```xml
<Control xsi:type="Button" id="msgReadFunctionButton">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
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

### <a name="showtaskpane-button-example"></a><span data-ttu-id="33642-158">Exemplo do botão ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="33642-158">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="33642-159">Controles de menu (botão suspenso)</span><span class="sxs-lookup"><span data-stu-id="33642-159">Menu (dropdown button) controls</span></span>

<span data-ttu-id="33642-p110">Um menu define uma lista estática de opções. Cada item de menu executa uma função ou mostra um painel de tarefas. Não há suporte para submenus.</span><span class="sxs-lookup"><span data-stu-id="33642-p110">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="33642-163">Quando usado com um **ponto de extensão\*\*\*\*PrimaryCommandSurface** ou [ContextMenu](extensionpoint.md), o controle de menu define:</span><span class="sxs-lookup"><span data-stu-id="33642-163">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="33642-164">Um item de menu no nível raiz.</span><span class="sxs-lookup"><span data-stu-id="33642-164">A root-level menu item.</span></span>

- <span data-ttu-id="33642-165">Uma lista de itens de submenu.</span><span class="sxs-lookup"><span data-stu-id="33642-165">A list of submenu items.</span></span>

<span data-ttu-id="33642-p111">Quando usado com **PrimaryCommandSurface**, o item de menu raiz é exibido como um botão na faixa de opções. Quando o botão é selecionado, o submenu é exibido como uma lista suspensa. Quando usado com **ContextMenu**, um item de menu com um submenu é inserido no menu de contexto. Em ambos os casos, cada item de submenu pode executar uma função JavaScript ou mostrar um painel de tarefas. Somente um nível de submenus é compatível no momento.</span><span class="sxs-lookup"><span data-stu-id="33642-p111">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="33642-p112">O exemplo a seguir mostra como definir um item de menu com dois itens de submenu. O primeiro item do submenu mostra um painel de tarefas e o segundo item executa uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="33642-p112">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="33642-173">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="33642-173">Child elements</span></span>

|  <span data-ttu-id="33642-174">Elemento</span><span class="sxs-lookup"><span data-stu-id="33642-174">Element</span></span> |  <span data-ttu-id="33642-175">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="33642-175">Required</span></span>  |  <span data-ttu-id="33642-176">Descrição</span><span class="sxs-lookup"><span data-stu-id="33642-176">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="33642-177">**Label**</span><span class="sxs-lookup"><span data-stu-id="33642-177">**Label**</span></span>     | <span data-ttu-id="33642-178">Sim</span><span class="sxs-lookup"><span data-stu-id="33642-178">Yes</span></span> |  <span data-ttu-id="33642-179">O texto do botão.</span><span class="sxs-lookup"><span data-stu-id="33642-179">The text for the button.</span></span> <span data-ttu-id="33642-180">O **atributo resid** pode ter no máximo 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no elemento [Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="33642-180">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="33642-181">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="33642-181">**ToolTip**</span></span>    |<span data-ttu-id="33642-182">Não</span><span class="sxs-lookup"><span data-stu-id="33642-182">No</span></span>|<span data-ttu-id="33642-183">A dica de ferramenta do botão.</span><span class="sxs-lookup"><span data-stu-id="33642-183">The tooltip for the button.</span></span> <span data-ttu-id="33642-184">O **atributo resid** pode ter no máximo 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String.**</span><span class="sxs-lookup"><span data-stu-id="33642-184">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element.</span></span> <span data-ttu-id="33642-185">O elemento **String** é um elemento filho do elemento **LongStrings**, que é filho do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="33642-185">The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="33642-186">Supertip</span><span class="sxs-lookup"><span data-stu-id="33642-186">Supertip</span></span>](supertip.md)  | <span data-ttu-id="33642-187">Sim</span><span class="sxs-lookup"><span data-stu-id="33642-187">Yes</span></span> |  <span data-ttu-id="33642-188">A dica detalhada do botão.</span><span class="sxs-lookup"><span data-stu-id="33642-188">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="33642-189">Icon</span><span class="sxs-lookup"><span data-stu-id="33642-189">Icon</span></span>](icon.md)      | <span data-ttu-id="33642-190">Sim</span><span class="sxs-lookup"><span data-stu-id="33642-190">Yes</span></span> |  <span data-ttu-id="33642-191">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="33642-191">An image for the button.</span></span>         |
|  <span data-ttu-id="33642-192">**Items**</span><span class="sxs-lookup"><span data-stu-id="33642-192">**Items**</span></span>     | <span data-ttu-id="33642-193">Sim</span><span class="sxs-lookup"><span data-stu-id="33642-193">Yes</span></span> |  <span data-ttu-id="33642-194">Um conjunto de botões a exibir dentro do menu.</span><span class="sxs-lookup"><span data-stu-id="33642-194">A collection of Buttons to display within the menu.</span></span> <span data-ttu-id="33642-195">Contém os elementos **Item** para cada item do submenu.</span><span class="sxs-lookup"><span data-stu-id="33642-195">Contains the **Item** elements for each submenu item.</span></span> <span data-ttu-id="33642-196">Cada elemento **Item** contém os mesmos elementos filhos do [Controle de botão](#button-control).</span><span class="sxs-lookup"><span data-stu-id="33642-196">Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|
|  [<span data-ttu-id="33642-197">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="33642-197">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="33642-198">Não</span><span class="sxs-lookup"><span data-stu-id="33642-198">No</span></span> |  <span data-ttu-id="33642-199">Especifica se o menu deve aparecer em combinações de aplicativos e plataformas que suportam guias contextuais personalizadas.</span><span class="sxs-lookup"><span data-stu-id="33642-199">Specifies whether the menu should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="33642-200">Se usado, ele deve ser o *primeiro* elemento filho.</span><span class="sxs-lookup"><span data-stu-id="33642-200">If used, it must be the *first* child element.</span></span> |

### <a name="menu-control-examples"></a><span data-ttu-id="33642-201">Exemplo de controle de menu</span><span class="sxs-lookup"><span data-stu-id="33642-201">Menu control examples</span></span>

```xml
<Control xsi:type="Menu" id="TestMenu2">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
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
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
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

## <a name="mobilebutton-control"></a><span data-ttu-id="33642-202">Controle MobileButton</span><span class="sxs-lookup"><span data-stu-id="33642-202">MobileButton control</span></span>

<span data-ttu-id="33642-p117">Um botão móvel executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função ou a exibição de um painel de tarefas. Cada controle de botão móvel deve ter um único `id` para o manifesto.</span><span class="sxs-lookup"><span data-stu-id="33642-p117">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="33642-p118">O valor `MobileButton` de **xsi:type** é definido no esquema VersionOverrides 1.1. O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="33642-p118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="33642-208">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="33642-208">Child elements</span></span>
|  <span data-ttu-id="33642-209">Elemento</span><span class="sxs-lookup"><span data-stu-id="33642-209">Element</span></span> |  <span data-ttu-id="33642-210">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="33642-210">Required</span></span>  |  <span data-ttu-id="33642-211">Descrição</span><span class="sxs-lookup"><span data-stu-id="33642-211">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="33642-212">**Label**</span><span class="sxs-lookup"><span data-stu-id="33642-212">**Label**</span></span>     | <span data-ttu-id="33642-213">Sim</span><span class="sxs-lookup"><span data-stu-id="33642-213">Yes</span></span> |  <span data-ttu-id="33642-214">O texto do botão.</span><span class="sxs-lookup"><span data-stu-id="33642-214">The text for the button.</span></span> <span data-ttu-id="33642-215">O **atributo resid** pode ter no máximo 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no elemento [Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="33642-215">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="33642-216">Icon</span><span class="sxs-lookup"><span data-stu-id="33642-216">Icon</span></span>](icon.md)      | <span data-ttu-id="33642-217">Sim</span><span class="sxs-lookup"><span data-stu-id="33642-217">Yes</span></span> |  <span data-ttu-id="33642-218">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="33642-218">An image for the button.</span></span>         |
|  [<span data-ttu-id="33642-219">Action</span><span class="sxs-lookup"><span data-stu-id="33642-219">Action</span></span>](action.md)    | <span data-ttu-id="33642-220">Sim</span><span class="sxs-lookup"><span data-stu-id="33642-220">Yes</span></span> |  <span data-ttu-id="33642-221">Especifica a ação a ser executada.</span><span class="sxs-lookup"><span data-stu-id="33642-221">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="33642-222">Exemplo de botão móvel ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="33642-222">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="33642-223">Exemplo de botão móvel ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="33642-223">ShowTaskpane mobile button example</span></span>

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
