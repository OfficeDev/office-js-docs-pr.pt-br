---
title: Elemento Group no arquivo de manifesto
description: ''
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: ad1a566e259188ed20032bc5a3004736474e1f01
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670129"
---
# <a name="group-element"></a><span data-ttu-id="86214-102">Elemento Group</span><span class="sxs-lookup"><span data-stu-id="86214-102">Group element</span></span>

<span data-ttu-id="86214-p101">Define um grupo de controles de interface do usuário em uma guia.  Em guias personalizadas, o suplemento pode criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="86214-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="86214-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="86214-106">Attributes</span></span>

|  <span data-ttu-id="86214-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="86214-107">Attribute</span></span>  |  <span data-ttu-id="86214-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="86214-108">Required</span></span>  |  <span data-ttu-id="86214-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="86214-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="86214-110">id</span><span class="sxs-lookup"><span data-stu-id="86214-110">id</span></span>](#id-attribute)  |  <span data-ttu-id="86214-111">Sim</span><span class="sxs-lookup"><span data-stu-id="86214-111">Yes</span></span>  | <span data-ttu-id="86214-112">Identificação exclusiva do grupo.</span><span class="sxs-lookup"><span data-stu-id="86214-112">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="86214-113">id attribute</span><span class="sxs-lookup"><span data-stu-id="86214-113">id attribute</span></span>

<span data-ttu-id="86214-p102">Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Esse valor deve ser exclusivo dentro o manifesto, ou o grupo não será processado.</span><span class="sxs-lookup"><span data-stu-id="86214-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="86214-118">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="86214-118">Child elements</span></span>
|  <span data-ttu-id="86214-119">Elemento</span><span class="sxs-lookup"><span data-stu-id="86214-119">Element</span></span> |  <span data-ttu-id="86214-120">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="86214-120">Required</span></span>  |  <span data-ttu-id="86214-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="86214-121">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="86214-122">Label</span><span class="sxs-lookup"><span data-stu-id="86214-122">Label</span></span>](#label)      | <span data-ttu-id="86214-123">Sim</span><span class="sxs-lookup"><span data-stu-id="86214-123">Yes</span></span> |  <span data-ttu-id="86214-124">O rótulo para a CustomTab ou um grupo.</span><span class="sxs-lookup"><span data-stu-id="86214-124">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="86214-125">Control</span><span class="sxs-lookup"><span data-stu-id="86214-125">Control</span></span>](#control)    | <span data-ttu-id="86214-126">Sim</span><span class="sxs-lookup"><span data-stu-id="86214-126">Yes</span></span> |  <span data-ttu-id="86214-127">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="86214-127">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="86214-128">Label</span><span class="sxs-lookup"><span data-stu-id="86214-128">Label</span></span> 

<span data-ttu-id="86214-p103">Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="86214-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="86214-132">Control</span><span class="sxs-lookup"><span data-stu-id="86214-132">Control</span></span>
<span data-ttu-id="86214-133">Um grupo exige pelo menos um controle.</span><span class="sxs-lookup"><span data-stu-id="86214-133">A group requires at least one control.</span></span> <span data-ttu-id="86214-134">Para obter detalhes sobre os tipos de controles suportados, consulte o elemento [Control](control.md) .</span><span class="sxs-lookup"><span data-stu-id="86214-134">For details about the types of controls that are supported, see the [Control](control.md) element.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```
