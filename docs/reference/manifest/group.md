---
title: Elemento Group no arquivo de manifesto
description: Define um grupo de controles da interface do usuário em uma guia.
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: a598232f230a120dccd58024e760c2172a769727
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611824"
---
# <a name="group-element"></a><span data-ttu-id="051f7-103">Elemento Group</span><span class="sxs-lookup"><span data-stu-id="051f7-103">Group element</span></span>

<span data-ttu-id="051f7-p101">Define um grupo de controles de interface do usuário em uma guia.  Em guias personalizadas, o suplemento pode criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="051f7-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="051f7-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="051f7-107">Attributes</span></span>

|  <span data-ttu-id="051f7-108">Atributo</span><span class="sxs-lookup"><span data-stu-id="051f7-108">Attribute</span></span>  |  <span data-ttu-id="051f7-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="051f7-109">Required</span></span>  |  <span data-ttu-id="051f7-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="051f7-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="051f7-111">id</span><span class="sxs-lookup"><span data-stu-id="051f7-111">id</span></span>](#id-attribute)  |  <span data-ttu-id="051f7-112">Sim</span><span class="sxs-lookup"><span data-stu-id="051f7-112">Yes</span></span>  | <span data-ttu-id="051f7-113">Identificação exclusiva do grupo.</span><span class="sxs-lookup"><span data-stu-id="051f7-113">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="051f7-114">id attribute</span><span class="sxs-lookup"><span data-stu-id="051f7-114">id attribute</span></span>

<span data-ttu-id="051f7-p102">Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Esse valor deve ser exclusivo dentro o manifesto, ou o grupo não será processado.</span><span class="sxs-lookup"><span data-stu-id="051f7-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="051f7-119">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="051f7-119">Child elements</span></span>
|  <span data-ttu-id="051f7-120">Elemento</span><span class="sxs-lookup"><span data-stu-id="051f7-120">Element</span></span> |  <span data-ttu-id="051f7-121">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="051f7-121">Required</span></span>  |  <span data-ttu-id="051f7-122">Descrição</span><span class="sxs-lookup"><span data-stu-id="051f7-122">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="051f7-123">Label</span><span class="sxs-lookup"><span data-stu-id="051f7-123">Label</span></span>](#label)      | <span data-ttu-id="051f7-124">Sim</span><span class="sxs-lookup"><span data-stu-id="051f7-124">Yes</span></span> |  <span data-ttu-id="051f7-125">O rótulo para a CustomTab ou um grupo.</span><span class="sxs-lookup"><span data-stu-id="051f7-125">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="051f7-126">Icon</span><span class="sxs-lookup"><span data-stu-id="051f7-126">Icon</span></span>](icon.md)      | <span data-ttu-id="051f7-127">Sim</span><span class="sxs-lookup"><span data-stu-id="051f7-127">Yes</span></span> |  <span data-ttu-id="051f7-128">A imagem de um grupo.</span><span class="sxs-lookup"><span data-stu-id="051f7-128">The image for a group.</span></span>  |
|  [<span data-ttu-id="051f7-129">Control</span><span class="sxs-lookup"><span data-stu-id="051f7-129">Control</span></span>](#control)    | <span data-ttu-id="051f7-130">Sim</span><span class="sxs-lookup"><span data-stu-id="051f7-130">Yes</span></span> |  <span data-ttu-id="051f7-131">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="051f7-131">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="051f7-132">Rótulo</span><span class="sxs-lookup"><span data-stu-id="051f7-132">Label</span></span> 

<span data-ttu-id="051f7-133">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="051f7-133">Required.</span></span> <span data-ttu-id="051f7-134">O rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="051f7-134">The label of the group.</span></span> <span data-ttu-id="051f7-135">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="051f7-135">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="051f7-136">Ícone</span><span class="sxs-lookup"><span data-stu-id="051f7-136">Icon</span></span>

<span data-ttu-id="051f7-137">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="051f7-137">Required.</span></span> <span data-ttu-id="051f7-138">Se uma guia contiver muitos grupos e a janela do programa for redimensionada, a imagem especificada poderá ser exibida.</span><span class="sxs-lookup"><span data-stu-id="051f7-138">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="051f7-139">Control</span><span class="sxs-lookup"><span data-stu-id="051f7-139">Control</span></span>
<span data-ttu-id="051f7-140">Um grupo exige pelo menos um controle.</span><span class="sxs-lookup"><span data-stu-id="051f7-140">A group requires at least one control.</span></span> <span data-ttu-id="051f7-141">Para obter detalhes sobre os tipos de controles suportados, consulte o elemento [Control](control.md) .</span><span class="sxs-lookup"><span data-stu-id="051f7-141">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

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
