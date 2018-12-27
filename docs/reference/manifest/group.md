---
title: Elemento Group no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 13cd9bbe6f602fd1779caea487e34177c3e9d483
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433689"
---
# <a name="group-element"></a><span data-ttu-id="ed2fe-102">Elemento Group</span><span class="sxs-lookup"><span data-stu-id="ed2fe-102">Group element</span></span>

<span data-ttu-id="ed2fe-p101">Define um grupo de controles de interface do usuário em uma guia.  Em guias personalizadas, o suplemento pode criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="ed2fe-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="ed2fe-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="ed2fe-106">Attributes</span></span>

|  <span data-ttu-id="ed2fe-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="ed2fe-107">Attribute</span></span>  |  <span data-ttu-id="ed2fe-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ed2fe-108">Required</span></span>  |  <span data-ttu-id="ed2fe-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="ed2fe-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ed2fe-110">id</span><span class="sxs-lookup"><span data-stu-id="ed2fe-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="ed2fe-111">Sim</span><span class="sxs-lookup"><span data-stu-id="ed2fe-111">Yes</span></span>  | <span data-ttu-id="ed2fe-112">Identificação exclusiva do grupo.</span><span class="sxs-lookup"><span data-stu-id="ed2fe-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="ed2fe-113">atributo id</span><span class="sxs-lookup"><span data-stu-id="ed2fe-113">id attribute</span></span>

<span data-ttu-id="ed2fe-p102">Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Esse valor deve ser exclusivo dentro o manifesto, ou o grupo não será processado.</span><span class="sxs-lookup"><span data-stu-id="ed2fe-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ed2fe-118">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ed2fe-118">Child elements</span></span>
|  <span data-ttu-id="ed2fe-119">Elemento</span><span class="sxs-lookup"><span data-stu-id="ed2fe-119">Element</span></span> |  <span data-ttu-id="ed2fe-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ed2fe-120">Required</span></span>  |  <span data-ttu-id="ed2fe-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="ed2fe-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ed2fe-122">Rótulo</span><span class="sxs-lookup"><span data-stu-id="ed2fe-122">Label</span></span>](#label)      | <span data-ttu-id="ed2fe-123">Sim</span><span class="sxs-lookup"><span data-stu-id="ed2fe-123">Yes</span></span> |  <span data-ttu-id="ed2fe-124">O rótulo para a CustomTab ou um grupo.</span><span class="sxs-lookup"><span data-stu-id="ed2fe-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="ed2fe-125">Control</span><span class="sxs-lookup"><span data-stu-id="ed2fe-125">Control</span></span>](#control)    | <span data-ttu-id="ed2fe-126">Sim</span><span class="sxs-lookup"><span data-stu-id="ed2fe-126">Yes</span></span> |  <span data-ttu-id="ed2fe-127">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="ed2fe-127">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="ed2fe-128">Rótulo</span><span class="sxs-lookup"><span data-stu-id="ed2fe-128">Label</span></span> 

<span data-ttu-id="ed2fe-p103">Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="ed2fe-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="ed2fe-132">Control</span><span class="sxs-lookup"><span data-stu-id="ed2fe-132">Control</span></span>
<span data-ttu-id="ed2fe-133">Um grupo exige pelo menos um controle.</span><span class="sxs-lookup"><span data-stu-id="ed2fe-133">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```