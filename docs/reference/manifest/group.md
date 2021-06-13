---
title: Elemento Group no arquivo de manifesto
description: Define um grupo de controles de interface do usuário em uma guia.
ms.date: 06/08/2021
localization_priority: Normal
ms.openlocfilehash: 89ed16f7996ab06bd21e1ebaa71c959b11af2029
ms.sourcegitcommit: ab3d38f2829e83f624bf43c49c0d267166552eec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/11/2021
ms.locfileid: "52893509"
---
# <a name="group-element"></a><span data-ttu-id="0be6e-103">Elemento Group</span><span class="sxs-lookup"><span data-stu-id="0be6e-103">Group element</span></span>

<span data-ttu-id="0be6e-104">Define um grupo de controles de interface do usuário em uma guia. Em guias personalizadas, o complemento pode criar vários grupos.</span><span class="sxs-lookup"><span data-stu-id="0be6e-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="0be6e-105">Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="0be6e-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="0be6e-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="0be6e-106">Attributes</span></span>

|  <span data-ttu-id="0be6e-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="0be6e-107">Attribute</span></span>  |  <span data-ttu-id="0be6e-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0be6e-108">Required</span></span>  |  <span data-ttu-id="0be6e-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="0be6e-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0be6e-110">id</span><span class="sxs-lookup"><span data-stu-id="0be6e-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="0be6e-111">Sim</span><span class="sxs-lookup"><span data-stu-id="0be6e-111">Yes</span></span>  | <span data-ttu-id="0be6e-112">Identificação exclusiva do grupo.</span><span class="sxs-lookup"><span data-stu-id="0be6e-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="0be6e-113">id attribute</span><span class="sxs-lookup"><span data-stu-id="0be6e-113">id attribute</span></span>

<span data-ttu-id="0be6e-p102">Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Esse valor deve ser exclusivo dentro o manifesto, ou o grupo não será processado.</span><span class="sxs-lookup"><span data-stu-id="0be6e-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="0be6e-118">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="0be6e-118">Child elements</span></span>

|  <span data-ttu-id="0be6e-119">Elemento</span><span class="sxs-lookup"><span data-stu-id="0be6e-119">Element</span></span> |  <span data-ttu-id="0be6e-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="0be6e-120">Required</span></span>  |  <span data-ttu-id="0be6e-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="0be6e-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0be6e-122">Label</span><span class="sxs-lookup"><span data-stu-id="0be6e-122">Label</span></span>](#label)      | <span data-ttu-id="0be6e-123">Sim</span><span class="sxs-lookup"><span data-stu-id="0be6e-123">Yes</span></span> |  <span data-ttu-id="0be6e-124">O rótulo para a CustomTab ou um grupo.</span><span class="sxs-lookup"><span data-stu-id="0be6e-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="0be6e-125">Icon</span><span class="sxs-lookup"><span data-stu-id="0be6e-125">Icon</span></span>](icon.md)      | <span data-ttu-id="0be6e-126">Sim</span><span class="sxs-lookup"><span data-stu-id="0be6e-126">Yes</span></span> |  <span data-ttu-id="0be6e-127">A imagem de um grupo.</span><span class="sxs-lookup"><span data-stu-id="0be6e-127">The image for a group.</span></span> <span data-ttu-id="0be6e-128">Não há suporte em Outlook de complementos.</span><span class="sxs-lookup"><span data-stu-id="0be6e-128">Not supported in Outlook add-ins.</span></span> |
|  [<span data-ttu-id="0be6e-129">Control</span><span class="sxs-lookup"><span data-stu-id="0be6e-129">Control</span></span>](#control)    | <span data-ttu-id="0be6e-130">Não</span><span class="sxs-lookup"><span data-stu-id="0be6e-130">No</span></span> |  <span data-ttu-id="0be6e-131">Representa um objeto Control.</span><span class="sxs-lookup"><span data-stu-id="0be6e-131">Represents a Control object.</span></span> <span data-ttu-id="0be6e-132">Pode ser zero ou mais.</span><span class="sxs-lookup"><span data-stu-id="0be6e-132">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="0be6e-133">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="0be6e-133">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="0be6e-134">Não</span><span class="sxs-lookup"><span data-stu-id="0be6e-134">No</span></span> | <span data-ttu-id="0be6e-135">Representa um dos controles internos Office internos.</span><span class="sxs-lookup"><span data-stu-id="0be6e-135">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="0be6e-136">Pode ser zero ou mais.</span><span class="sxs-lookup"><span data-stu-id="0be6e-136">Can be zero or more.</span></span> <span data-ttu-id="0be6e-137">Não há suporte em Outlook de complementos.</span><span class="sxs-lookup"><span data-stu-id="0be6e-137">Not supported in Outlook add-ins.</span></span>|
|  [<span data-ttu-id="0be6e-138">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="0be6e-138">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="0be6e-139">Não</span><span class="sxs-lookup"><span data-stu-id="0be6e-139">No</span></span> |  <span data-ttu-id="0be6e-140">Especifica se o grupo deve aparecer em combinações de aplicativos e plataformas que suportam guias contextuais personalizadas.</span><span class="sxs-lookup"><span data-stu-id="0be6e-140">Specifies whether the group should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="0be6e-141">Não há suporte em Outlook de complementos.</span><span class="sxs-lookup"><span data-stu-id="0be6e-141">Not supported in Outlook add-ins.</span></span> |

### <a name="label"></a><span data-ttu-id="0be6e-142">Rótulo</span><span class="sxs-lookup"><span data-stu-id="0be6e-142">Label</span></span>

<span data-ttu-id="0be6e-143">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="0be6e-143">Required.</span></span> <span data-ttu-id="0be6e-144">O rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="0be6e-144">The label of the group.</span></span> <span data-ttu-id="0be6e-145">O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no [elemento Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="0be6e-145">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="0be6e-146">Ícone</span><span class="sxs-lookup"><span data-stu-id="0be6e-146">Icon</span></span>

<span data-ttu-id="0be6e-147">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="0be6e-147">Required.</span></span> <span data-ttu-id="0be6e-148">Se uma guia contiver muitos grupos e a janela do programa for resized, a imagem especificada poderá ser exibida.</span><span class="sxs-lookup"><span data-stu-id="0be6e-148">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

> [!NOTE]
> <span data-ttu-id="0be6e-149">Esse elemento filho não é suportado em Outlook de complementos.</span><span class="sxs-lookup"><span data-stu-id="0be6e-149">This child element is not supported in Outlook add-ins.</span></span>

### <a name="control"></a><span data-ttu-id="0be6e-150">Controle</span><span class="sxs-lookup"><span data-stu-id="0be6e-150">Control</span></span>

<span data-ttu-id="0be6e-151">Opcional, mas se não estiver presente, deve haver pelo menos um **OfficeControl**.</span><span class="sxs-lookup"><span data-stu-id="0be6e-151">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="0be6e-152">Para obter detalhes sobre os tipos de controles com suporte, consulte o [elemento Control.](control.md)</span><span class="sxs-lookup"><span data-stu-id="0be6e-152">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="0be6e-153">A ordem de **Control** e **OfficeControl** no manifesto é intercambiável e eles podem ser intercambiáveis se houver vários elementos, mas todos devem estar abaixo do **elemento Icon.**</span><span class="sxs-lookup"><span data-stu-id="0be6e-153">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

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

### <a name="officecontrol"></a><span data-ttu-id="0be6e-154">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="0be6e-154">OfficeControl</span></span>

<span data-ttu-id="0be6e-155">Opcional, mas se não estiver presente, deve haver pelo menos um **Control**.</span><span class="sxs-lookup"><span data-stu-id="0be6e-155">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="0be6e-156">Inclua um ou mais controles internos Office no grupo com `<OfficeControl>` elementos.</span><span class="sxs-lookup"><span data-stu-id="0be6e-156">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="0be6e-157">O `id` atributo especifica a ID do controle Office integrado.</span><span class="sxs-lookup"><span data-stu-id="0be6e-157">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="0be6e-158">Para encontrar a ID de um controle, consulte [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span><span class="sxs-lookup"><span data-stu-id="0be6e-158">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="0be6e-159">A ordem de **Control** e **OfficeControl** no manifesto é intercambiável e eles podem ser intercambiáveis se houver vários elementos, mas todos devem estar abaixo do **elemento Icon.**</span><span class="sxs-lookup"><span data-stu-id="0be6e-159">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

> [!NOTE]
> <span data-ttu-id="0be6e-160">Esse elemento filho não é suportado em Outlook de complementos.</span><span class="sxs-lookup"><span data-stu-id="0be6e-160">This child element is not supported in Outlook add-ins.</span></span>

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

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="0be6e-161">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="0be6e-161">OverriddenByRibbonApi</span></span>

<span data-ttu-id="0be6e-162">Opcional (booleano).</span><span class="sxs-lookup"><span data-stu-id="0be6e-162">Optional (boolean).</span></span> <span data-ttu-id="0be6e-163">Especifica se o **Grupo** ficará oculto em combinações de aplicativos e plataformas que suportam uma API que instala uma guia contextual personalizada na faixa de opções no tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="0be6e-163">Specifies whether the **Group** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="0be6e-164">O valor padrão, se não estiver presente, é `false` .</span><span class="sxs-lookup"><span data-stu-id="0be6e-164">The default value, if not present, is `false`.</span></span> <span data-ttu-id="0be6e-165">Se usado, **OverriddenByRibbonApi** deve ser o *primeiro* filho de **Group**.</span><span class="sxs-lookup"><span data-stu-id="0be6e-165">If used, **OverriddenByRibbonApi** must be the *first* child of **Group**.</span></span> <span data-ttu-id="0be6e-166">Para obter mais informações, [consulte OverriddenByRibbonApi](overriddenbyribbonapi.md).</span><span class="sxs-lookup"><span data-stu-id="0be6e-166">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

> [!NOTE]
> <span data-ttu-id="0be6e-167">Esse elemento filho não é suportado em Outlook de complementos.</span><span class="sxs-lookup"><span data-stu-id="0be6e-167">This child element is not supported in Outlook add-ins.</span></span>

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
