---
title: Elemento CustomTab no arquivo de manifesto
description: Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento.
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: a81b64a17eeeb463d55024e189b09048b2eb96ac
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612301"
---
# <a name="customtab-element"></a><span data-ttu-id="15c67-103">Elemento CustomTab</span><span class="sxs-lookup"><span data-stu-id="15c67-103">CustomTab element</span></span>

<span data-ttu-id="15c67-104">Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="15c67-104">On the ribbon, you specify which tab and group for their add-in commands.</span></span> <span data-ttu-id="15c67-105">Isso pode ser a guia padrão ( **página inicial**, de **mensagem**ou **reunião**) ou em uma guia personalizada definida pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="15c67-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="15c67-p102">Nas guias personalizadas, o suplemento poderá criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="15c67-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="15c67-109">O atributo **ID** deve ser exclusivo dentro do manifesto.</span><span class="sxs-lookup"><span data-stu-id="15c67-109">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="15c67-110">No Outlook no Mac, o `CustomTab` elemento não está disponível, portanto, você terá que usar o [OfficeTab](officetab.md) .</span><span class="sxs-lookup"><span data-stu-id="15c67-110">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="15c67-111">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="15c67-111">Child elements</span></span>

|  <span data-ttu-id="15c67-112">Elemento</span><span class="sxs-lookup"><span data-stu-id="15c67-112">Element</span></span> |  <span data-ttu-id="15c67-113">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="15c67-113">Required</span></span>  |  <span data-ttu-id="15c67-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="15c67-114">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="15c67-115">Group</span><span class="sxs-lookup"><span data-stu-id="15c67-115">Group</span></span>](group.md)      | <span data-ttu-id="15c67-116">Sim</span><span class="sxs-lookup"><span data-stu-id="15c67-116">Yes</span></span> |  <span data-ttu-id="15c67-117">Define um grupo de comandos</span><span class="sxs-lookup"><span data-stu-id="15c67-117">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="15c67-118">Label</span><span class="sxs-lookup"><span data-stu-id="15c67-118">Label</span></span>](#label-tab)      | <span data-ttu-id="15c67-119">Sim</span><span class="sxs-lookup"><span data-stu-id="15c67-119">Yes</span></span> |  <span data-ttu-id="15c67-120">O rótulo para CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="15c67-120">The label for the CustomTab or a Group.</span></span>  |

### <a name="group"></a><span data-ttu-id="15c67-121">Group</span><span class="sxs-lookup"><span data-stu-id="15c67-121">Group</span></span>

<span data-ttu-id="15c67-p103">Obrigatório. Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="15c67-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="15c67-124">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="15c67-124">Label (Tab)</span></span>

<span data-ttu-id="15c67-125">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="15c67-125">Required.</span></span> <span data-ttu-id="15c67-126">O rótulo da guia personalizado. O atributo **Resid** deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="15c67-126">The label of the custom tab. The **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="15c67-127">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="15c67-127">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
