---
title: Elemento CustomTab no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 7d609ad216ba5e8e7358bbc741f7b6c992bc97e2
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433603"
---
# <a name="customtab-element"></a><span data-ttu-id="c2e25-102">Elemento CustomTab</span><span class="sxs-lookup"><span data-stu-id="c2e25-102">CustomTab element</span></span>

<span data-ttu-id="c2e25-p101">Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento. Isso pode ser realizado na guia padrão (**Início**, **Mensagem** ou **Reunião**) ou em uma guia personalizada definida pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="c2e25-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="c2e25-p102">Nas guias personalizadas, o suplemento poderá criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual ele aparece. Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="c2e25-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="c2e25-108">O atributo **id** deve ser exclusivo dentro do manifesto.</span><span class="sxs-lookup"><span data-stu-id="c2e25-108">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c2e25-109">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="c2e25-109">Child elements</span></span>

|  <span data-ttu-id="c2e25-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="c2e25-110">Element</span></span> |  <span data-ttu-id="c2e25-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="c2e25-111">Required</span></span>  |  <span data-ttu-id="c2e25-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="c2e25-112">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c2e25-113">Grupo</span><span class="sxs-lookup"><span data-stu-id="c2e25-113">Group</span></span>](group.md)      | <span data-ttu-id="c2e25-114">Sim</span><span class="sxs-lookup"><span data-stu-id="c2e25-114">Yes</span></span> |  <span data-ttu-id="c2e25-115">Define um grupo de comandos</span><span class="sxs-lookup"><span data-stu-id="c2e25-115">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="c2e25-116">Rótulo</span><span class="sxs-lookup"><span data-stu-id="c2e25-116">Label</span></span>](#label-tab)      | <span data-ttu-id="c2e25-117">Sim</span><span class="sxs-lookup"><span data-stu-id="c2e25-117">Yes</span></span> |  <span data-ttu-id="c2e25-118">O rótulo para CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="c2e25-118">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="c2e25-119">Control</span><span class="sxs-lookup"><span data-stu-id="c2e25-119">Control</span></span>](control.md)    | <span data-ttu-id="c2e25-120">Sim</span><span class="sxs-lookup"><span data-stu-id="c2e25-120">Yes</span></span> |  <span data-ttu-id="c2e25-121">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="c2e25-121">A collection of one or more Control objects.</span></span>  |

### <a name="group"></a><span data-ttu-id="c2e25-122">Group</span><span class="sxs-lookup"><span data-stu-id="c2e25-122">Group</span></span>

<span data-ttu-id="c2e25-p103">Obrigatório. Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="c2e25-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="c2e25-125">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="c2e25-125">Label (Tab)</span></span>

<span data-ttu-id="c2e25-p104">Obrigatório. O rótulo da guia personalizada. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="c2e25-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="c2e25-128">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="c2e25-128">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```