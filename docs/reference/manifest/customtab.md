---
title: Elemento CustomTab no arquivo de manifesto
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: c48e526534a3c1295e9c3f0c6fc626df94a874d3
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554010"
---
# <a name="customtab-element"></a><span data-ttu-id="604ad-102">Elemento CustomTab</span><span class="sxs-lookup"><span data-stu-id="604ad-102">CustomTab element</span></span>

<span data-ttu-id="604ad-p101">Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento. Isso pode ser realizado na guia padrão (**Início**, **Mensagem** ou **Reunião**) ou em uma guia personalizada definida pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="604ad-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="604ad-p102">Nas guias personalizadas, o suplemento poderá criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="604ad-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="604ad-108">O atributo **id** deve ser exclusivo dentro do manifesto.</span><span class="sxs-lookup"><span data-stu-id="604ad-108">The  **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="604ad-109">No Outlook no Mac, o `CustomTab` elemento não está disponível, portanto, você terá que usar o [OfficeTab](officetab.md) .</span><span class="sxs-lookup"><span data-stu-id="604ad-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="604ad-110">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="604ad-110">Child elements</span></span>

|  <span data-ttu-id="604ad-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="604ad-111">Element</span></span> |  <span data-ttu-id="604ad-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="604ad-112">Required</span></span>  |  <span data-ttu-id="604ad-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="604ad-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="604ad-114">Group</span><span class="sxs-lookup"><span data-stu-id="604ad-114">Group</span></span>](group.md)      | <span data-ttu-id="604ad-115">Sim</span><span class="sxs-lookup"><span data-stu-id="604ad-115">Yes</span></span> |  <span data-ttu-id="604ad-116">Define um grupo de comandos</span><span class="sxs-lookup"><span data-stu-id="604ad-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="604ad-117">Label</span><span class="sxs-lookup"><span data-stu-id="604ad-117">Label</span></span>](#label-tab)      | <span data-ttu-id="604ad-118">Sim</span><span class="sxs-lookup"><span data-stu-id="604ad-118">Yes</span></span> |  <span data-ttu-id="604ad-119">O rótulo para CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="604ad-119">The label for the CustomTab or a Group.</span></span>  |

### <a name="group"></a><span data-ttu-id="604ad-120">Group</span><span class="sxs-lookup"><span data-stu-id="604ad-120">Group</span></span>

<span data-ttu-id="604ad-p103">Obrigatório. Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="604ad-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="604ad-123">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="604ad-123">Label (Tab)</span></span>

<span data-ttu-id="604ad-p104">Obrigatório. O rótulo da guia personalizada. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="604ad-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="604ad-126">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="604ad-126">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
