---
title: Elemento CustomTab no arquivo de manifesto
description: ''
ms.date: 04/29/2019
localization_priority: Normal
ms.openlocfilehash: 4fa7dd86736b5ab421be5653f2e256a6b84fb480
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/21/2019
ms.locfileid: "33517391"
---
# <a name="customtab-element"></a><span data-ttu-id="1398c-102">Elemento CustomTab</span><span class="sxs-lookup"><span data-stu-id="1398c-102">CustomTab element</span></span>

<span data-ttu-id="1398c-p101">Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento. Isso pode ser realizado na guia padrão (**Início**, **Mensagem** ou **Reunião**) ou em uma guia personalizada definida pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="1398c-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="1398c-p102">Nas guias personalizadas, o suplemento poderá criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="1398c-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="1398c-108">O atributo **id** deve ser exclusivo dentro do manifesto.</span><span class="sxs-lookup"><span data-stu-id="1398c-108">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1398c-109">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="1398c-109">Child elements</span></span>

|  <span data-ttu-id="1398c-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="1398c-110">Element</span></span> |  <span data-ttu-id="1398c-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1398c-111">Required</span></span>  |  <span data-ttu-id="1398c-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="1398c-112">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1398c-113">Group</span><span class="sxs-lookup"><span data-stu-id="1398c-113">Group</span></span>](group.md)      | <span data-ttu-id="1398c-114">Sim</span><span class="sxs-lookup"><span data-stu-id="1398c-114">Yes</span></span> |  <span data-ttu-id="1398c-115">Define um grupo de comandos</span><span class="sxs-lookup"><span data-stu-id="1398c-115">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="1398c-116">Label</span><span class="sxs-lookup"><span data-stu-id="1398c-116">Label</span></span>](#label-tab)      | <span data-ttu-id="1398c-117">Sim</span><span class="sxs-lookup"><span data-stu-id="1398c-117">Yes</span></span> |  <span data-ttu-id="1398c-118">O rótulo para CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="1398c-118">The label for the CustomTab or a Group.</span></span>  |

### <a name="group"></a><span data-ttu-id="1398c-119">Group</span><span class="sxs-lookup"><span data-stu-id="1398c-119">Group</span></span>

<span data-ttu-id="1398c-p103">Obrigatório. Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="1398c-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="1398c-122">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="1398c-122">Label (Tab)</span></span>

<span data-ttu-id="1398c-p104">Obrigatório. O rótulo da guia personalizada. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="1398c-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="1398c-125">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="1398c-125">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```
