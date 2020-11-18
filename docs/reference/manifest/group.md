---
title: Elemento Group no arquivo de manifesto
description: Define um grupo de controles da interface do usuário em uma guia.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 6ee8d499767eccb95b4fdf9ceb91dd2cd12bce95
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087942"
---
# <a name="group-element"></a><span data-ttu-id="1d1b4-103">Elemento Group</span><span class="sxs-lookup"><span data-stu-id="1d1b4-103">Group element</span></span>

<span data-ttu-id="1d1b4-104">Define um grupo de controles da interface do usuário em uma guia. Nas guias personalizadas, o suplemento pode criar vários grupos.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-104">Defines a group of UI controls in a tab. On custom tabs, the add-in can create multiple groups.</span></span> <span data-ttu-id="1d1b4-105">Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-105">Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="1d1b4-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="1d1b4-106">Attributes</span></span>

|  <span data-ttu-id="1d1b4-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="1d1b4-107">Attribute</span></span>  |  <span data-ttu-id="1d1b4-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1d1b4-108">Required</span></span>  |  <span data-ttu-id="1d1b4-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d1b4-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1d1b4-110">id</span><span class="sxs-lookup"><span data-stu-id="1d1b4-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="1d1b4-111">Sim</span><span class="sxs-lookup"><span data-stu-id="1d1b4-111">Yes</span></span>  | <span data-ttu-id="1d1b4-112">Identificação exclusiva do grupo.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="1d1b4-113">id attribute</span><span class="sxs-lookup"><span data-stu-id="1d1b4-113">id attribute</span></span>

<span data-ttu-id="1d1b4-p102">Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Esse valor deve ser exclusivo dentro o manifesto, ou o grupo não será processado.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1d1b4-118">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="1d1b4-118">Child elements</span></span>

|  <span data-ttu-id="1d1b4-119">Elemento</span><span class="sxs-lookup"><span data-stu-id="1d1b4-119">Element</span></span> |  <span data-ttu-id="1d1b4-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1d1b4-120">Required</span></span>  |  <span data-ttu-id="1d1b4-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="1d1b4-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1d1b4-122">Label</span><span class="sxs-lookup"><span data-stu-id="1d1b4-122">Label</span></span>](#label)      | <span data-ttu-id="1d1b4-123">Sim</span><span class="sxs-lookup"><span data-stu-id="1d1b4-123">Yes</span></span> |  <span data-ttu-id="1d1b4-124">O rótulo para a CustomTab ou um grupo.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="1d1b4-125">Icon</span><span class="sxs-lookup"><span data-stu-id="1d1b4-125">Icon</span></span>](icon.md)      | <span data-ttu-id="1d1b4-126">Sim</span><span class="sxs-lookup"><span data-stu-id="1d1b4-126">Yes</span></span> |  <span data-ttu-id="1d1b4-127">A imagem de um grupo.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="1d1b4-128">Control</span><span class="sxs-lookup"><span data-stu-id="1d1b4-128">Control</span></span>](#control)    | <span data-ttu-id="1d1b4-129">Não</span><span class="sxs-lookup"><span data-stu-id="1d1b4-129">No</span></span> |  <span data-ttu-id="1d1b4-130">Representa um objeto Control.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-130">Represents a Control object.</span></span> <span data-ttu-id="1d1b4-131">Pode ser zero ou mais.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-131">Can be zero or more.</span></span>  |
|  [<span data-ttu-id="1d1b4-132">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="1d1b4-132">OfficeControl</span></span>](#officecontrol)  | <span data-ttu-id="1d1b4-133">Não</span><span class="sxs-lookup"><span data-stu-id="1d1b4-133">No</span></span> | <span data-ttu-id="1d1b4-134">Representa um dos controles internos do Office.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-134">Represents one of the built-in Office controls.</span></span> <span data-ttu-id="1d1b4-135">Pode ser zero ou mais.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-135">Can be zero or more.</span></span> |

### <a name="label"></a><span data-ttu-id="1d1b4-136">Rótulo</span><span class="sxs-lookup"><span data-stu-id="1d1b4-136">Label</span></span>

<span data-ttu-id="1d1b4-137">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-137">Required.</span></span> <span data-ttu-id="1d1b4-138">O rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-138">The label of the group.</span></span> <span data-ttu-id="1d1b4-139">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="1d1b4-139">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="1d1b4-140">Ícone</span><span class="sxs-lookup"><span data-stu-id="1d1b4-140">Icon</span></span>

<span data-ttu-id="1d1b4-141">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-141">Required.</span></span> <span data-ttu-id="1d1b4-142">Se uma guia contiver muitos grupos e a janela do programa for redimensionada, a imagem especificada poderá ser exibida.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-142">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="1d1b4-143">Controle</span><span class="sxs-lookup"><span data-stu-id="1d1b4-143">Control</span></span>

<span data-ttu-id="1d1b4-144">Opcional, mas, se não estiver presente, deve haver pelo menos um **OfficeControl**.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-144">Optional, but if not present there must be at least one **OfficeControl**.</span></span> <span data-ttu-id="1d1b4-145">Para obter detalhes sobre os tipos de controles suportados, consulte o elemento [Control](control.md) .</span><span class="sxs-lookup"><span data-stu-id="1d1b4-145">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span> <span data-ttu-id="1d1b4-146">A ordem de **controle** e **OfficeControl** no manifesto é intercambiável e podem ser mescladas se houver vários elementos, mas todos devem estar abaixo do elemento **Icon** .</span><span class="sxs-lookup"><span data-stu-id="1d1b4-146">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
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

### <a name="officecontrol"></a><span data-ttu-id="1d1b4-147">OfficeControl</span><span class="sxs-lookup"><span data-stu-id="1d1b4-147">OfficeControl</span></span>

<span data-ttu-id="1d1b4-148">Opcional, mas, se não estiver presente, deve haver pelo menos um **controle**.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-148">Optional, but if not present there must be at least one **Control**.</span></span> <span data-ttu-id="1d1b4-149">Inclua um ou mais controles internos do Office no grupo com `<OfficeControl>` elementos.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-149">Include one or more built-in Office controls in the group with `<OfficeControl>` elements.</span></span> <span data-ttu-id="1d1b4-150">O `id` atributo especifica a ID do controle interno do Office.</span><span class="sxs-lookup"><span data-stu-id="1d1b4-150">The `id` attribute specifies the ID of the built-in Office control.</span></span> <span data-ttu-id="1d1b4-151">Para localizar a ID de um controle, confira [localizar as IDs de controles e grupos de controle](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span><span class="sxs-lookup"><span data-stu-id="1d1b4-151">To find the ID of a control, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="1d1b4-152">A ordem de **controle** e **OfficeControl** no manifesto é intercambiável e podem ser mescladas se houver vários elementos, mas todos devem estar abaixo do elemento **Icon** .</span><span class="sxs-lookup"><span data-stu-id="1d1b4-152">The order of **Control** and **OfficeControl** in the manifest is interchangeable and they can be intermingled if there are multiple elements, but all must be below the **Icon** element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
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
