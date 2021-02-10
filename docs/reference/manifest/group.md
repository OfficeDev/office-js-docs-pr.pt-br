---
title: Elemento Group no arquivo de manifesto
description: Define um grupo de controles de interface do usuário em uma guia.
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 1bb3a4d65e954a54acb6e93f7c4d52e6b0845315
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173959"
---
# <a name="group-element"></a><span data-ttu-id="85834-103">Elemento Group</span><span class="sxs-lookup"><span data-stu-id="85834-103">Group element</span></span>

<span data-ttu-id="85834-104">Define um grupo de controles de interface do usuário em uma guia. Em guias personalizadas, o complemento pode criar vários grupos.</span><span class="sxs-lookup"><span data-stu-id="85834-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="85834-105">Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="85834-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="85834-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="85834-106">Attributes</span></span>

|  <span data-ttu-id="85834-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="85834-107">Attribute</span></span>  |  <span data-ttu-id="85834-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="85834-108">Required</span></span>  |  <span data-ttu-id="85834-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="85834-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="85834-110">id</span><span class="sxs-lookup"><span data-stu-id="85834-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="85834-111">Sim</span><span class="sxs-lookup"><span data-stu-id="85834-111">Yes</span></span>  | <span data-ttu-id="85834-112">Identificação exclusiva do grupo.</span><span class="sxs-lookup"><span data-stu-id="85834-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="85834-113">id attribute</span><span class="sxs-lookup"><span data-stu-id="85834-113">id attribute</span></span>

<span data-ttu-id="85834-p102">Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Esse valor deve ser exclusivo dentro o manifesto, ou o grupo não será processado.</span><span class="sxs-lookup"><span data-stu-id="85834-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="85834-118">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="85834-118">Child elements</span></span>

|  <span data-ttu-id="85834-119">Elemento</span><span class="sxs-lookup"><span data-stu-id="85834-119">Element</span></span> |  <span data-ttu-id="85834-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="85834-120">Required</span></span>  |  <span data-ttu-id="85834-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="85834-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="85834-122">Label</span><span class="sxs-lookup"><span data-stu-id="85834-122">Label</span></span>](#label)      | <span data-ttu-id="85834-123">Sim</span><span class="sxs-lookup"><span data-stu-id="85834-123">Yes</span></span> |  <span data-ttu-id="85834-124">O rótulo para a CustomTab ou um grupo.</span><span class="sxs-lookup"><span data-stu-id="85834-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="85834-125">Icon</span><span class="sxs-lookup"><span data-stu-id="85834-125">Icon</span></span>](icon.md)      | <span data-ttu-id="85834-126">Sim</span><span class="sxs-lookup"><span data-stu-id="85834-126">Yes</span></span> |  <span data-ttu-id="85834-127">A imagem de um grupo.</span><span class="sxs-lookup"><span data-stu-id="85834-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="85834-128">Control</span><span class="sxs-lookup"><span data-stu-id="85834-128">Control</span></span>](#control)    | <span data-ttu-id="85834-129">Não</span><span class="sxs-lookup"><span data-stu-id="85834-129">No</span></span> |  <span data-ttu-id="85834-130">Representa um objeto Control .</span><span class="sxs-lookup"><span data-stu-id="85834-130">Represents a Control object.</span></span> <span data-ttu-id="85834-131">Pode ser zero ou mais.</span><span class="sxs-lookup"><span data-stu-id="85834-131">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="85834-132">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="85834-132">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="85834-133">Não</span><span class="sxs-lookup"><span data-stu-id="85834-133">No</span></span> | <span data-ttu-id="85834-134">Representa um dos controles internos do Office.</span><span class="sxs-lookup"><span data-stu-id="85834-134">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="85834-135">Pode ser zero ou mais.</span><span class="sxs-lookup"><span data-stu-id="85834-135">Can be zero or more.</span></span> |
|  [<span data-ttu-id="85834-136">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="85834-136">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="85834-137">Não</span><span class="sxs-lookup"><span data-stu-id="85834-137">No</span></span> |  <span data-ttu-id="85834-138">Especifica se o grupo deve aparecer em combinações de aplicativos e plataformas que suportam guias contextuais personalizadas.</span><span class="sxs-lookup"><span data-stu-id="85834-138">Specifies whether the group should appear on application and platform combinations that support custom contextual tabs.</span></span>  |

### <a name="label"></a><span data-ttu-id="85834-139">Rótulo</span><span class="sxs-lookup"><span data-stu-id="85834-139">Label</span></span>

<span data-ttu-id="85834-140">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="85834-140">Required.</span></span> <span data-ttu-id="85834-141">O rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="85834-141">The label of the group.</span></span> <span data-ttu-id="85834-142">O **atributo resid** pode ter no máximo 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no elemento [Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="85834-142">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="85834-143">Ícone</span><span class="sxs-lookup"><span data-stu-id="85834-143">Icon</span></span>

<span data-ttu-id="85834-144">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="85834-144">Required.</span></span> <span data-ttu-id="85834-145">Se uma guia contiver muitos grupos e a janela do programa for resizedida, a imagem especificada poderá ser exibida em vez disso.</span><span class="sxs-lookup"><span data-stu-id="85834-145">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="85834-146">Controle</span><span class="sxs-lookup"><span data-stu-id="85834-146">Control</span></span>

<span data-ttu-id="85834-147">Opcional, mas se não houver deve haver pelo menos um **OfficeControl**.</span><span class="sxs-lookup"><span data-stu-id="85834-147">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="85834-148">Para obter detalhes sobre os tipos de controles com suporte, consulte o [elemento Control.](control.md)</span><span class="sxs-lookup"><span data-stu-id="85834-148">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="85834-149">A ordem de **Controle** e **OfficeControl** no manifesto é intercambiável e eles podem ser intercambiáveis se houver vários elementos, mas todos devem estar abaixo do **elemento Icon.**</span><span class="sxs-lookup"><span data-stu-id="85834-149">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```

### <a name="officecontrol"></a><span data-ttu-id="85834-150">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="85834-150">OfficeControl</span></span>

<span data-ttu-id="85834-151">Opcional, mas se não houver deve haver pelo menos um **controle**.</span><span class="sxs-lookup"><span data-stu-id="85834-151">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="85834-152">Inclua um ou mais controles internos do Office no grupo com `<OfficeControl>` elementos.</span><span class="sxs-lookup"><span data-stu-id="85834-152">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="85834-153">O `id` atributo especifica a ID do controle office integrado.</span><span class="sxs-lookup"><span data-stu-id="85834-153">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="85834-154">Para encontrar a ID de um controle, consulte [Encontrar as IDs de controles e grupos de controles.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="85834-154">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="85834-155">A ordem de **Controle** e **OfficeControl** no manifesto é intercambiável e eles podem ser intercambiáveis se houver vários elementos, mas todos devem estar abaixo do **elemento Icon.**</span><span class="sxs-lookup"><span data-stu-id="85834-155">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="contosoCustomTab.grp1">
    <Label resid="CustomTabGroupLabel"/>
    <Icon>
        <bt:Image size="16" resid="blue-icon-16" />
        <bt:Image size="32" resid="blue-icon-32" />
        <bt:Image size="80" resid="blue-icon-80" />
    </Icon>
    <Control xsi:type="Button" id="Button2">
        <!-- information on the control -->
    </Control>
    <OfficeControl id="Superscript" />
    <!-- other controls, as needed -->
</Group>
```

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="85834-156">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="85834-156">OverriddenByRibbonApi</span></span>

<span data-ttu-id="85834-157">Opcional (booliana).</span><span class="sxs-lookup"><span data-stu-id="85834-157">Optional (boolean).</span></span> <span data-ttu-id="85834-158">Especifica se  o grupo ficará oculto em combinações de aplicativos e plataformas que suportam uma API que instala uma guia contextual personalizada na faixa de opções em tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="85834-158">Specifies whether the **Group** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="85834-159">O valor padrão, se não estiver presente, é `false` .</span><span class="sxs-lookup"><span data-stu-id="85834-159">The default value, if not present, is `false`.</span></span> <span data-ttu-id="85834-160">Se usado, **OverriddenByRibbonApi** deve ser o *primeiro* filho de **Group**.</span><span class="sxs-lookup"><span data-stu-id="85834-160">If used, **OverriddenByRibbonApi** must be the *first* child of **Group**.</span></span> <span data-ttu-id="85834-161">Para obter mais informações, [consulte OverriddenByRibbonApi](overriddenbyribbonapi.md).</span><span class="sxs-lookup"><span data-stu-id="85834-161">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="ContosoCustomTab.grp1">
      <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
      <!-- other child elements of the group -->
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
