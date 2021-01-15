---
title: Elemento CustomTab no arquivo de manifesto
description: Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento.
ms.date: 11/01/2020
localization_priority: Normal
ms.openlocfilehash: 642222af02431814e4e64141504911c67ca829fa
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771323"
---
# <a name="customtab-element"></a><span data-ttu-id="b3e45-103">Elemento CustomTab</span><span class="sxs-lookup"><span data-stu-id="b3e45-103">CustomTab element</span></span>

<span data-ttu-id="b3e45-104">Na faixa de opções, especifique a guia e o grupo para os comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="b3e45-104">On the ribbon, specify the tab and group for your add-in commands.</span></span> <span data-ttu-id="b3e45-105">Isso pode ser a guia padrão ( **página inicial**, de **mensagem** ou **reunião**) ou em uma guia personalizada definida pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="b3e45-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="b3e45-106">Nas guias personalizadas, o suplemento pode ter grupos internos ou personalizados.</span><span class="sxs-lookup"><span data-stu-id="b3e45-106">On custom tabs, the add-in can have custom or built-in groups.</span></span> <span data-ttu-id="b3e45-107">Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="b3e45-107">Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="b3e45-108">O atributo **ID** deve ser exclusivo dentro do manifesto.</span><span class="sxs-lookup"><span data-stu-id="b3e45-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b3e45-109">No Outlook no Mac, o `CustomTab` elemento não está disponível, portanto, você terá que usar o [OfficeTab](officetab.md) .</span><span class="sxs-lookup"><span data-stu-id="b3e45-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b3e45-110">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="b3e45-110">Child elements</span></span>

|  <span data-ttu-id="b3e45-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="b3e45-111">Element</span></span> |  <span data-ttu-id="b3e45-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="b3e45-112">Required</span></span>  |  <span data-ttu-id="b3e45-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="b3e45-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b3e45-114">Group</span><span class="sxs-lookup"><span data-stu-id="b3e45-114">Group</span></span>](group.md)      | <span data-ttu-id="b3e45-115">Não</span><span class="sxs-lookup"><span data-stu-id="b3e45-115">No</span></span> |  <span data-ttu-id="b3e45-116">Define um grupo de comandos</span><span class="sxs-lookup"><span data-stu-id="b3e45-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="b3e45-117">O Microsoft Office</span><span class="sxs-lookup"><span data-stu-id="b3e45-117">OfficeGroup</span></span>](#officegroup)      | <span data-ttu-id="b3e45-118">Não</span><span class="sxs-lookup"><span data-stu-id="b3e45-118">No</span></span> |  <span data-ttu-id="b3e45-119">Representa um grupo de controle interno do Office.</span><span class="sxs-lookup"><span data-stu-id="b3e45-119">Represents a built-in Office control group.</span></span>  |
|  [<span data-ttu-id="b3e45-120">Label</span><span class="sxs-lookup"><span data-stu-id="b3e45-120">Label</span></span>](#label-tab)      | <span data-ttu-id="b3e45-121">Sim</span><span class="sxs-lookup"><span data-stu-id="b3e45-121">Yes</span></span> |  <span data-ttu-id="b3e45-122">O rótulo para CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="b3e45-122">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="b3e45-123">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="b3e45-123">InsertAfter</span></span>](#insertafter)      | <span data-ttu-id="b3e45-124">Não</span><span class="sxs-lookup"><span data-stu-id="b3e45-124">No</span></span> |  <span data-ttu-id="b3e45-125">Especifica que a guia personalizada deve ser imediatamente após uma guia interna especificada do Office.</span><span class="sxs-lookup"><span data-stu-id="b3e45-125">Specifies that the custom tab should be immediately after a specified built-in Office tab.</span></span>  |
|  [<span data-ttu-id="b3e45-126">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="b3e45-126">InsertBefore</span></span>](#insertbefore)      | <span data-ttu-id="b3e45-127">Não</span><span class="sxs-lookup"><span data-stu-id="b3e45-127">No</span></span> |  <span data-ttu-id="b3e45-128">Especifica que a guia personalizada deve ser imediatamente anterior à guia interna especificada do Office.</span><span class="sxs-lookup"><span data-stu-id="b3e45-128">Specifies that the custom tab should be immediately before a specified built-in Office tab.</span></span>  |

### <a name="group"></a><span data-ttu-id="b3e45-129">Grupo</span><span class="sxs-lookup"><span data-stu-id="b3e45-129">Group</span></span>

<span data-ttu-id="b3e45-130">Opcional, mas, se não estiver presente, deve haver pelo **menos um elemento** de um.</span><span class="sxs-lookup"><span data-stu-id="b3e45-130">Optional, but if not present there must be at least one **OfficeGroup** element.</span></span> <span data-ttu-id="b3e45-131">Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="b3e45-131">See [Group element](group.md).</span></span> <span data-ttu-id="b3e45-132">A ordem do **grupo** e do grupo do **Office** no manifesto deve ser a ordem que você deseja que eles apareçam na guia Personalizar. Eles podem ser mesclados se houver vários elementos, mas todos devem estar acima do elemento **rótulo** .</span><span class="sxs-lookup"><span data-stu-id="b3e45-132">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="officegroup"></a><span data-ttu-id="b3e45-133">O Microsoft Office</span><span class="sxs-lookup"><span data-stu-id="b3e45-133">OfficeGroup</span></span>

<span data-ttu-id="b3e45-134">Opcional, mas se não houver, deve haver pelo menos um elemento de **grupo** .</span><span class="sxs-lookup"><span data-stu-id="b3e45-134">Optional, but if not present there must be at least one **Group** element.</span></span> <span data-ttu-id="b3e45-135">Representa um grupo de controle interno do Office.</span><span class="sxs-lookup"><span data-stu-id="b3e45-135">Represents a built-in Office control group.</span></span> <span data-ttu-id="b3e45-136">O atributo **ID** especifica a ID do grupo interno do Office.</span><span class="sxs-lookup"><span data-stu-id="b3e45-136">The **id** attribute specifies the ID of the built-in Office group.</span></span> <span data-ttu-id="b3e45-137">Para localizar a ID de um grupo interno, confira [localizar as IDs de controles e grupos de controle](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span><span class="sxs-lookup"><span data-stu-id="b3e45-137">To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="b3e45-138">A ordem do **grupo** e do grupo do **Office** no manifesto deve ser a ordem que você deseja que eles apareçam na guia Personalizar. Eles podem ser mesclados se houver vários elementos, mas todos devem estar acima do elemento **rótulo** .</span><span class="sxs-lookup"><span data-stu-id="b3e45-138">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="label-tab"></a><span data-ttu-id="b3e45-139">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="b3e45-139">Label (Tab)</span></span>

<span data-ttu-id="b3e45-140">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="b3e45-140">Required.</span></span> <span data-ttu-id="b3e45-141">O rótulo da guia personalizado. O atributo **Resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **ID** de um elemento **String** no elemento **ShortStrings** no elemento [Resources](resources.md) .</span><span class="sxs-lookup"><span data-stu-id="b3e45-141">The label of the custom tab. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="insertafter"></a><span data-ttu-id="b3e45-142">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="b3e45-142">InsertAfter</span></span>

<span data-ttu-id="b3e45-143">Opcional.</span><span class="sxs-lookup"><span data-stu-id="b3e45-143">Optional.</span></span> <span data-ttu-id="b3e45-144">Especifica que a guia personalizada deve ser imediatamente após uma guia interna especificada do Office. O valor do elemento é a ID da guia interna, como "TabHome" ou "TabReview".</span><span class="sxs-lookup"><span data-stu-id="b3e45-144">Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="b3e45-145">(Consulte [localizar as IDs de controles e grupos de controle](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) Se presente, deve ser após o elemento **Label** .</span><span class="sxs-lookup"><span data-stu-id="b3e45-145">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element.</span></span> <span data-ttu-id="b3e45-146">Você não pode ter **InsertAfter** e **InsertBefore**.</span><span class="sxs-lookup"><span data-stu-id="b3e45-146">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

### <a name="insertbefore"></a><span data-ttu-id="b3e45-147">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="b3e45-147">InsertBefore</span></span>

<span data-ttu-id="b3e45-148">Opcional.</span><span class="sxs-lookup"><span data-stu-id="b3e45-148">Optional.</span></span> <span data-ttu-id="b3e45-149">Especifica que a guia personalizada deve ser imediatamente anterior à guia interna especificada do Office. O valor do elemento é a ID da guia interna, como "TabHome" ou "TabReview".</span><span class="sxs-lookup"><span data-stu-id="b3e45-149">Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="b3e45-150">(Consulte [localizar as IDs de controles e grupos de controle](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  Se presente, deve ser após o elemento **Label** .</span><span class="sxs-lookup"><span data-stu-id="b3e45-150">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element.</span></span> <span data-ttu-id="b3e45-151">Você não pode ter **InsertAfter** e **InsertBefore**.</span><span class="sxs-lookup"><span data-stu-id="b3e45-151">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

## <a name="customtab-example"></a><span data-ttu-id="b3e45-152">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="b3e45-152">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
