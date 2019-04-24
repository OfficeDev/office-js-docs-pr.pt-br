---
title: Elemento CustomTab no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c1c3c6883a1feb94299feb35c078431e6e2e322c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450629"
---
# <a name="customtab-element"></a><span data-ttu-id="3b3d9-102">Elemento CustomTab</span><span class="sxs-lookup"><span data-stu-id="3b3d9-102">CustomTab element</span></span>

<span data-ttu-id="3b3d9-p101">Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento. Isso pode ser realizado na guia padrão (**Início**, **Mensagem** ou **Reunião**) ou em uma guia personalizada definida pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="3b3d9-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="3b3d9-p102">Nas guias personalizadas, o suplemento poderá criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="3b3d9-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="3b3d9-108">O atributo **id** deve ser exclusivo dentro do manifesto.</span><span class="sxs-lookup"><span data-stu-id="3b3d9-108">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="3b3d9-109">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3b3d9-109">Child elements</span></span>

|  <span data-ttu-id="3b3d9-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="3b3d9-110">Element</span></span> |  <span data-ttu-id="3b3d9-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3b3d9-111">Required</span></span>  |  <span data-ttu-id="3b3d9-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="3b3d9-112">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3b3d9-113">Group</span><span class="sxs-lookup"><span data-stu-id="3b3d9-113">Group</span></span>](group.md)      | <span data-ttu-id="3b3d9-114">Sim</span><span class="sxs-lookup"><span data-stu-id="3b3d9-114">Yes</span></span> |  <span data-ttu-id="3b3d9-115">Define um grupo de comandos</span><span class="sxs-lookup"><span data-stu-id="3b3d9-115">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="3b3d9-116">Label</span><span class="sxs-lookup"><span data-stu-id="3b3d9-116">Label</span></span>](#label-tab)      | <span data-ttu-id="3b3d9-117">Sim</span><span class="sxs-lookup"><span data-stu-id="3b3d9-117">Yes</span></span> |  <span data-ttu-id="3b3d9-118">O rótulo para CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="3b3d9-118">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="3b3d9-119">Control</span><span class="sxs-lookup"><span data-stu-id="3b3d9-119">Control</span></span>](control.md)    | <span data-ttu-id="3b3d9-120">Sim</span><span class="sxs-lookup"><span data-stu-id="3b3d9-120">Yes</span></span> |  <span data-ttu-id="3b3d9-121">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="3b3d9-121">A collection of one or more Control objects.</span></span>  |

### <a name="group"></a><span data-ttu-id="3b3d9-122">Group</span><span class="sxs-lookup"><span data-stu-id="3b3d9-122">Group</span></span>

<span data-ttu-id="3b3d9-p103">Obrigatório. Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="3b3d9-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="3b3d9-125">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="3b3d9-125">Label (Tab)</span></span>

<span data-ttu-id="3b3d9-p104">Obrigatório. O rótulo da guia personalizada. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="3b3d9-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="3b3d9-128">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="3b3d9-128">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
