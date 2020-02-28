---
title: Elemento Group no arquivo de manifesto
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 27a168ea17352482e955e7a0d1f8267c7d6b17d8
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324859"
---
# <a name="group-element"></a><span data-ttu-id="1b35e-102">Elemento Group</span><span class="sxs-lookup"><span data-stu-id="1b35e-102">Group element</span></span>

<span data-ttu-id="1b35e-p101">Define um grupo de controles de interface do usuário em uma guia.  Em guias personalizadas, o suplemento pode criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="1b35e-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="1b35e-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="1b35e-106">Attributes</span></span>

|  <span data-ttu-id="1b35e-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="1b35e-107">Attribute</span></span>  |  <span data-ttu-id="1b35e-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1b35e-108">Required</span></span>  |  <span data-ttu-id="1b35e-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="1b35e-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1b35e-110">id</span><span class="sxs-lookup"><span data-stu-id="1b35e-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="1b35e-111">Sim</span><span class="sxs-lookup"><span data-stu-id="1b35e-111">Yes</span></span>  | <span data-ttu-id="1b35e-112">Identificação exclusiva do grupo.</span><span class="sxs-lookup"><span data-stu-id="1b35e-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="1b35e-113">id attribute</span><span class="sxs-lookup"><span data-stu-id="1b35e-113">id attribute</span></span>

<span data-ttu-id="1b35e-p102">Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Esse valor deve ser exclusivo dentro o manifesto, ou o grupo não será processado.</span><span class="sxs-lookup"><span data-stu-id="1b35e-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1b35e-118">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="1b35e-118">Child elements</span></span>
|  <span data-ttu-id="1b35e-119">Elemento</span><span class="sxs-lookup"><span data-stu-id="1b35e-119">Element</span></span> |  <span data-ttu-id="1b35e-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1b35e-120">Required</span></span>  |  <span data-ttu-id="1b35e-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="1b35e-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1b35e-122">Label</span><span class="sxs-lookup"><span data-stu-id="1b35e-122">Label</span></span>](#label)      | <span data-ttu-id="1b35e-123">Sim</span><span class="sxs-lookup"><span data-stu-id="1b35e-123">Yes</span></span> |  <span data-ttu-id="1b35e-124">O rótulo para a CustomTab ou um grupo.</span><span class="sxs-lookup"><span data-stu-id="1b35e-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="1b35e-125">Icon</span><span class="sxs-lookup"><span data-stu-id="1b35e-125">Icon</span></span>](icon.md)      | <span data-ttu-id="1b35e-126">Sim</span><span class="sxs-lookup"><span data-stu-id="1b35e-126">Yes</span></span> |  <span data-ttu-id="1b35e-127">A imagem de um grupo.</span><span class="sxs-lookup"><span data-stu-id="1b35e-127">The image for a group.</span></span>  |
|  [<span data-ttu-id="1b35e-128">Control</span><span class="sxs-lookup"><span data-stu-id="1b35e-128">Control</span></span>](#control)    | <span data-ttu-id="1b35e-129">Sim</span><span class="sxs-lookup"><span data-stu-id="1b35e-129">Yes</span></span> |  <span data-ttu-id="1b35e-130">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="1b35e-130">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="1b35e-131">Rótulo</span><span class="sxs-lookup"><span data-stu-id="1b35e-131">Label</span></span> 

<span data-ttu-id="1b35e-132">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="1b35e-132">Required.</span></span> <span data-ttu-id="1b35e-133">O rótulo do grupo.</span><span class="sxs-lookup"><span data-stu-id="1b35e-133">The label of the group.</span></span> <span data-ttu-id="1b35e-134">O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="1b35e-134">The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="icon"></a><span data-ttu-id="1b35e-135">Ícone</span><span class="sxs-lookup"><span data-stu-id="1b35e-135">Icon</span></span>

<span data-ttu-id="1b35e-136">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="1b35e-136">Required.</span></span> <span data-ttu-id="1b35e-137">Se uma guia contiver muitos grupos e a janela do programa for redimensionada, a imagem especificada poderá ser exibida.</span><span class="sxs-lookup"><span data-stu-id="1b35e-137">If a tab contains a lot of groups and the program window is resized, the specified image may display instead.</span></span>

### <a name="control"></a><span data-ttu-id="1b35e-138">Control</span><span class="sxs-lookup"><span data-stu-id="1b35e-138">Control</span></span>
<span data-ttu-id="1b35e-139">Um grupo exige pelo menos um controle.</span><span class="sxs-lookup"><span data-stu-id="1b35e-139">A group requires at least one control.</span></span> <span data-ttu-id="1b35e-140">Para obter detalhes sobre os tipos de controles suportados, consulte o elemento [Control](control.md) .</span><span class="sxs-lookup"><span data-stu-id="1b35e-140">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

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
