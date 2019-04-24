---
title: Elemento Group no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7cc1f4c398eeb013eb6033b207b395466f7d72ca
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450706"
---
# <a name="group-element"></a><span data-ttu-id="832e7-102">Elemento Group</span><span class="sxs-lookup"><span data-stu-id="832e7-102">Group element</span></span>

<span data-ttu-id="832e7-p101">Define um grupo de controles de interface do usuário em uma guia.  Em guias personalizadas, o suplemento pode criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="832e7-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="832e7-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="832e7-106">Attributes</span></span>

|  <span data-ttu-id="832e7-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="832e7-107">Attribute</span></span>  |  <span data-ttu-id="832e7-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="832e7-108">Required</span></span>  |  <span data-ttu-id="832e7-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="832e7-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="832e7-110">id</span><span class="sxs-lookup"><span data-stu-id="832e7-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="832e7-111">Sim</span><span class="sxs-lookup"><span data-stu-id="832e7-111">Yes</span></span>  | <span data-ttu-id="832e7-112">Identificação exclusiva do grupo.</span><span class="sxs-lookup"><span data-stu-id="832e7-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="832e7-113">id attribute</span><span class="sxs-lookup"><span data-stu-id="832e7-113">id attribute</span></span>

<span data-ttu-id="832e7-p102">Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Esse valor deve ser exclusivo dentro o manifesto, ou o grupo não será processado.</span><span class="sxs-lookup"><span data-stu-id="832e7-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="832e7-118">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="832e7-118">Child elements</span></span>
|  <span data-ttu-id="832e7-119">Elemento</span><span class="sxs-lookup"><span data-stu-id="832e7-119">Element</span></span> |  <span data-ttu-id="832e7-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="832e7-120">Required</span></span>  |  <span data-ttu-id="832e7-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="832e7-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="832e7-122">Label</span><span class="sxs-lookup"><span data-stu-id="832e7-122">Label</span></span>](#label)      | <span data-ttu-id="832e7-123">Sim</span><span class="sxs-lookup"><span data-stu-id="832e7-123">Yes</span></span> |  <span data-ttu-id="832e7-124">O rótulo para a CustomTab ou um grupo.</span><span class="sxs-lookup"><span data-stu-id="832e7-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="832e7-125">Control</span><span class="sxs-lookup"><span data-stu-id="832e7-125">Control</span></span>](#control)    | <span data-ttu-id="832e7-126">Sim</span><span class="sxs-lookup"><span data-stu-id="832e7-126">Yes</span></span> |  <span data-ttu-id="832e7-127">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="832e7-127">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="832e7-128">Label</span><span class="sxs-lookup"><span data-stu-id="832e7-128">Label</span></span> 

<span data-ttu-id="832e7-p103">Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="832e7-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="832e7-132">Control</span><span class="sxs-lookup"><span data-stu-id="832e7-132">Control</span></span>
<span data-ttu-id="832e7-133">Um grupo exige pelo menos um controle.</span><span class="sxs-lookup"><span data-stu-id="832e7-133">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```
