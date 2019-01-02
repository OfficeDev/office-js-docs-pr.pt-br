---
title: Elemento Control no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: e5d8574e322c21e768fb9f66fe9bbb0c12a34ed4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433933"
---
# <a name="control-element"></a><span data-ttu-id="14b38-102">Elemento Control</span><span class="sxs-lookup"><span data-stu-id="14b38-102">Control element</span></span>

<span data-ttu-id="14b38-p101">Define a função JavaScript que executa e aciona ou inicia um painel de tarefas. Um elemento **Control** pode ser um botão ou um menu. Pelo menos um **Control** deve ser incluído em um elemento [Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="14b38-p101">Defines a JavaScript function that executes an action or launches a task pane. A **Control** element can be either a button or a menu option. At least one **Control** must be included in a [Group](group.md) element.</span></span>

## <a name="attributes"></a><span data-ttu-id="14b38-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="14b38-106">Attributes</span></span>

|  <span data-ttu-id="14b38-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="14b38-107">Attribute</span></span>  |  <span data-ttu-id="14b38-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="14b38-108">Required</span></span>  |  <span data-ttu-id="14b38-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="14b38-109">Description</span></span>  |
|:-----|:-----|:-----|
|<span data-ttu-id="14b38-110">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="14b38-110">**xsi:type**</span></span>|<span data-ttu-id="14b38-111">Sim</span><span class="sxs-lookup"><span data-stu-id="14b38-111">Yes</span></span>|<span data-ttu-id="14b38-p102">O tipo de controle que está sendo definido. Pode ser `Button`, `Menu` ou `MobileButton`.</span><span class="sxs-lookup"><span data-stu-id="14b38-p102">The type of control being defined. Can be either `Button`, `Menu`, or `MobileButton`.</span></span> |
|<span data-ttu-id="14b38-114">**id**</span><span class="sxs-lookup"><span data-stu-id="14b38-114">**id**</span></span>|<span data-ttu-id="14b38-115">Não</span><span class="sxs-lookup"><span data-stu-id="14b38-115">No</span></span>|<span data-ttu-id="14b38-p103">A ID do elemento Control. Pode ter no máximo 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="14b38-p103">The ID of the control element. Can be a maximum of 125 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="14b38-118">O valor `MobileButton` de **xsi:type** é definido no esquema VersionOverrides 1.1.</span><span class="sxs-lookup"><span data-stu-id="14b38-118">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing VersionOverrides element must have an  attribute value of .</span></span> <span data-ttu-id="14b38-119">Ele só se aplica aos elementos **Control** contidos em um elemento [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="14b38-119">Note: The  value for xsi:type is defined in VersionOverrides schema 1.1. It only applies to the **Control** elements contained within a [MobileFormFactor](mobileformfactor.md) element.</span></span>

## <a name="button-control"></a><span data-ttu-id="14b38-120">Controle de botão</span><span class="sxs-lookup"><span data-stu-id="14b38-120">Button control</span></span>

<span data-ttu-id="14b38-p105">Um botão executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função ou a exibição de um painel de tarefas. Cada controle de botão deve ter um único `id` para o manifesto.</span><span class="sxs-lookup"><span data-stu-id="14b38-p105">A button performs a single action when the user selects it. It can either execute a function or show a task pane. Each button control must have an `id` unique to the manifest.</span></span> 

### <a name="child-elements"></a><span data-ttu-id="14b38-124">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="14b38-124">Child elements</span></span>
|  <span data-ttu-id="14b38-125">Elemento</span><span class="sxs-lookup"><span data-stu-id="14b38-125">Element</span></span> |  <span data-ttu-id="14b38-126">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="14b38-126">Required</span></span>  |  <span data-ttu-id="14b38-127">Descrição</span><span class="sxs-lookup"><span data-stu-id="14b38-127">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="14b38-128">**Rótulo**</span><span class="sxs-lookup"><span data-stu-id="14b38-128">**Label**</span></span>     | <span data-ttu-id="14b38-129">Sim</span><span class="sxs-lookup"><span data-stu-id="14b38-129">Yes</span></span> |  <span data-ttu-id="14b38-p106">O texto do botão. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="14b38-p106">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  <span data-ttu-id="14b38-132">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="14b38-132">**ToolTip**</span></span>  |<span data-ttu-id="14b38-133">Não</span><span class="sxs-lookup"><span data-stu-id="14b38-133">No</span></span>|<span data-ttu-id="14b38-p107">A dica de ferramenta do botão. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **LongStrings**, que é filho do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="14b38-p107">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="14b38-137">Supertip</span><span class="sxs-lookup"><span data-stu-id="14b38-137">Supertip</span></span>](supertip.md)  | <span data-ttu-id="14b38-138">Sim</span><span class="sxs-lookup"><span data-stu-id="14b38-138">Yes</span></span> |  <span data-ttu-id="14b38-139">A dica detalhada do botão.</span><span class="sxs-lookup"><span data-stu-id="14b38-139">The supertip for the button.</span></span>    |
|  [<span data-ttu-id="14b38-140">Icon</span><span class="sxs-lookup"><span data-stu-id="14b38-140">Icon</span></span>](icon.md)      | <span data-ttu-id="14b38-141">Sim</span><span class="sxs-lookup"><span data-stu-id="14b38-141">Yes</span></span> |  <span data-ttu-id="14b38-142">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="14b38-142">An image for the button.</span></span>         |
|  [<span data-ttu-id="14b38-143">Action</span><span class="sxs-lookup"><span data-stu-id="14b38-143">Action</span></span>](action.md)    | <span data-ttu-id="14b38-144">Sim</span><span class="sxs-lookup"><span data-stu-id="14b38-144">Yes</span></span> |  <span data-ttu-id="14b38-145">Especifica a ação a realizar.</span><span class="sxs-lookup"><span data-stu-id="14b38-145">Specifies the action to perform.</span></span>  |

### <a name="executefunction-button-example"></a><span data-ttu-id="14b38-146">Exemplo do botão ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="14b38-146">ExecuteFunction button example</span></span>

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

### <a name="showtaskpane-button-example"></a><span data-ttu-id="14b38-147">Exemplo do botão ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="14b38-147">ShowTaskpane button example</span></span>

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

## <a name="menu-dropdown-button-controls"></a><span data-ttu-id="14b38-148">Controles de menu (botão suspenso)</span><span class="sxs-lookup"><span data-stu-id="14b38-148">Menu (dropdown button) controls</span></span>

<span data-ttu-id="14b38-p108">Um menu define uma lista estática de opções. Cada item de menu executa uma função ou mostra um painel de tarefas. Não há suporte para submenus.</span><span class="sxs-lookup"><span data-stu-id="14b38-p108">A menu defines a static list of options. Each menu item either executes a function or shows a task pane. Submenus are not supported.</span></span> 

<span data-ttu-id="14b38-152">Quando usado com um [ponto de extensão](extensionpoint.md)**PrimaryCommandSurface** ou **ContextMenu**, o controle de menu define:</span><span class="sxs-lookup"><span data-stu-id="14b38-152">When used with a **PrimaryCommandSurface** or **ContextMenu** [extension point](extensionpoint.md), the menu control defines:</span></span>

- <span data-ttu-id="14b38-153">Um item de menu no nível raiz.</span><span class="sxs-lookup"><span data-stu-id="14b38-153">A root-level menu item.</span></span>

- <span data-ttu-id="14b38-154">Uma lista de itens de submenu.</span><span class="sxs-lookup"><span data-stu-id="14b38-154">A list of submenu items.</span></span>

<span data-ttu-id="14b38-p109">Quando usado com **PrimaryCommandSurface**, o item de menu raiz é exibido como um botão na faixa de opções. Quando o botão é selecionado, o submenu é exibido como uma lista suspensa. Quando usado com **ContextMenu**, um item de menu com um submenu é inserido no menu de contexto. Em ambos os casos, cada item de submenu pode executar uma função JavaScript ou mostrar um painel de tarefas. Somente há suporte para um nível de submenus no momento.</span><span class="sxs-lookup"><span data-stu-id="14b38-p109">When used with **PrimaryCommandSurface**, the root menu item displays as a button on the ribbon. When the button is selected, the submenu displays as a drop-down list. When used with  **ContextMenu**, a menu item with a submenu is inserted on the context menu. In both cases, individual submenu items can either execute a JavaScript function or show a task pane. Only one level of submenus is supported at this time.</span></span>

<span data-ttu-id="14b38-p110">O exemplo a seguir mostra como definir um item de menu com dois itens de submenu. O primeiro item do submenu mostra um painel de tarefas e o segundo item executa uma função JavaScript.</span><span class="sxs-lookup"><span data-stu-id="14b38-p110">The following example shows how to define a menu item with two submenu items. The first submenu item shows a task pane, and the second submenu item runs a JavaScript function.</span></span>

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

### <a name="child-elements"></a><span data-ttu-id="14b38-162">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="14b38-162">Child elements</span></span>

|  <span data-ttu-id="14b38-163">Elemento</span><span class="sxs-lookup"><span data-stu-id="14b38-163">Element</span></span> |  <span data-ttu-id="14b38-164">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="14b38-164">Required</span></span>  |  <span data-ttu-id="14b38-165">Descrição</span><span class="sxs-lookup"><span data-stu-id="14b38-165">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="14b38-166">**Rótulo**</span><span class="sxs-lookup"><span data-stu-id="14b38-166">**Label**</span></span>     | <span data-ttu-id="14b38-167">Sim</span><span class="sxs-lookup"><span data-stu-id="14b38-167">Yes</span></span> |  <span data-ttu-id="14b38-p111">O texto do botão. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="14b38-p111">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>      |
|  <span data-ttu-id="14b38-170">**ToolTip**</span><span class="sxs-lookup"><span data-stu-id="14b38-170">**ToolTip**</span></span>  |<span data-ttu-id="14b38-171">Não</span><span class="sxs-lookup"><span data-stu-id="14b38-171">No</span></span>|<span data-ttu-id="14b38-p112">A dica de ferramenta do botão. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String**. O elemento **String** é um elemento filho do elemento **LongStrings**, que é filho do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="14b38-p112">The tooltip for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element. The **String** element is a child of the **LongStrings** element, which is a child of the [Resources](resources.md) element.</span></span>|        
|  [<span data-ttu-id="14b38-175">Supertip</span><span class="sxs-lookup"><span data-stu-id="14b38-175">Supertip</span></span>](supertip.md)  | <span data-ttu-id="14b38-176">Sim</span><span class="sxs-lookup"><span data-stu-id="14b38-176">Yes</span></span> |  <span data-ttu-id="14b38-177">A dica detalhada do botão.</span><span class="sxs-lookup"><span data-stu-id="14b38-177">The supertip for this button.</span></span>    |
|  [<span data-ttu-id="14b38-178">Icon</span><span class="sxs-lookup"><span data-stu-id="14b38-178">Icon</span></span>](icon.md)      | <span data-ttu-id="14b38-179">Sim</span><span class="sxs-lookup"><span data-stu-id="14b38-179">Yes</span></span> |  <span data-ttu-id="14b38-180">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="14b38-180">An image for the button.</span></span>         |
|  <span data-ttu-id="14b38-181">**Itens**</span><span class="sxs-lookup"><span data-stu-id="14b38-181">**Items**</span></span>     | <span data-ttu-id="14b38-182">Sim</span><span class="sxs-lookup"><span data-stu-id="14b38-182">Yes</span></span> |  <span data-ttu-id="14b38-p113">Um conjunto de botões a exibir dentro do menu. Contém os elementos **Item** para cada item do submenu. Cada elemento **Item** contém os mesmos elementos filhos do [Controle de botão](#button-control).</span><span class="sxs-lookup"><span data-stu-id="14b38-p113">A collection of Buttons to display within the menu. Contains the  **Item** elements for each submenu item. Each **Item** element contains the  child elements of the [Button control](#button-control).</span></span>|

### <a name="menu-control-examples"></a><span data-ttu-id="14b38-186">Exemplo de controle de menu</span><span class="sxs-lookup"><span data-stu-id="14b38-186">Menu control examples</span></span>

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

## <a name="mobilebutton-control"></a><span data-ttu-id="14b38-187">Controle MobileButton</span><span class="sxs-lookup"><span data-stu-id="14b38-187">MobileButton control</span></span>

<span data-ttu-id="14b38-p114">Um botão móvel executa uma única ação quando o usuário o seleciona. Pode ser a execução de uma função ou a exibição de um painel de tarefas. Cada controle de botão móvel deve ter um único `id` para o manifesto.</span><span class="sxs-lookup"><span data-stu-id="14b38-p114">A mobile button performs a single action when the user selects it. It can either execute a function or show a task pane. Each mobile button control must have an `id` unique to the manifest.</span></span>

<span data-ttu-id="14b38-p115">O valor `MobileButton` de **xsi:type** é definido no esquema VersionOverrides 1.1. O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="14b38-p115">The `MobileButton` value for **xsi:type** is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

### <a name="child-elements"></a><span data-ttu-id="14b38-193">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="14b38-193">Child elements</span></span>
|  <span data-ttu-id="14b38-194">Elemento</span><span class="sxs-lookup"><span data-stu-id="14b38-194">Element</span></span> |  <span data-ttu-id="14b38-195">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="14b38-195">Required</span></span>  |  <span data-ttu-id="14b38-196">Descrição</span><span class="sxs-lookup"><span data-stu-id="14b38-196">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="14b38-197">**Rótulo**</span><span class="sxs-lookup"><span data-stu-id="14b38-197">**Label**</span></span>     | <span data-ttu-id="14b38-198">Sim</span><span class="sxs-lookup"><span data-stu-id="14b38-198">Yes</span></span> |  <span data-ttu-id="14b38-p116">O texto do botão. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="14b38-p116">The text for the button. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md)  element.</span></span>        |
|  [<span data-ttu-id="14b38-201">Icon</span><span class="sxs-lookup"><span data-stu-id="14b38-201">Icon</span></span>](icon.md)      | <span data-ttu-id="14b38-202">Sim</span><span class="sxs-lookup"><span data-stu-id="14b38-202">Yes</span></span> |  <span data-ttu-id="14b38-203">Uma imagem para o botão.</span><span class="sxs-lookup"><span data-stu-id="14b38-203">An image for the button.</span></span>         |
|  [<span data-ttu-id="14b38-204">Action</span><span class="sxs-lookup"><span data-stu-id="14b38-204">Action</span></span>](action.md)    | <span data-ttu-id="14b38-205">Sim</span><span class="sxs-lookup"><span data-stu-id="14b38-205">Yes</span></span> |  <span data-ttu-id="14b38-206">Especifica a ação a realizar.</span><span class="sxs-lookup"><span data-stu-id="14b38-206">Specifies the action to perform.</span></span>  |

### <a name="executefunction-mobile-button-example"></a><span data-ttu-id="14b38-207">Exemplo de botão móvel ExecuteFunction</span><span class="sxs-lookup"><span data-stu-id="14b38-207">ExecuteFunction mobile button example</span></span>

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

### <a name="showtaskpane-mobile-button-example"></a><span data-ttu-id="14b38-208">Exemplo de botão móvel ShowTaskpane</span><span class="sxs-lookup"><span data-stu-id="14b38-208">ShowTaskpane mobile button example</span></span>

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