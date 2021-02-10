---
title: Elemento CustomTab no arquivo de manifesto
description: Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: d74859d1326d29517b5a8226a86f901322957933
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173924"
---
# <a name="customtab-element"></a><span data-ttu-id="5248d-103">Elemento CustomTab</span><span class="sxs-lookup"><span data-stu-id="5248d-103">CustomTab element</span></span>

<span data-ttu-id="5248d-104">Na faixa de opções, especifique a guia e o grupo para os comandos do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="5248d-104">On the ribbon, specify the tab and group for your add-in commands.</span></span> <span data-ttu-id="5248d-105">Isso pode estar na guia padrão (Página **Início,** Mensagem ou **Reunião)** ou em uma guia personalizada definida pelo complemento.</span><span class="sxs-lookup"><span data-stu-id="5248d-105">This can either be on the default tab (either **Home**, **Message**, or **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="5248d-106">Em guias personalizadas, o complemento pode ter grupos personalizados ou integrados.</span><span class="sxs-lookup"><span data-stu-id="5248d-106">On custom tabs, the add-in can have custom or built-in groups.</span></span> <span data-ttu-id="5248d-107">Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="5248d-107">Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="5248d-108">O **atributo id** deve ser exclusivo dentro do manifesto.</span><span class="sxs-lookup"><span data-stu-id="5248d-108">The **id** attribute must be unique within the manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5248d-109">No Outlook no Mac, `CustomTab` o elemento não está disponível, portanto, você terá que usar o [OfficeTab.](officetab.md)</span><span class="sxs-lookup"><span data-stu-id="5248d-109">In Outlook on Mac, the `CustomTab` element is not available so you'll have to use [OfficeTab](officetab.md) instead.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5248d-110">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="5248d-110">Child elements</span></span>

|  <span data-ttu-id="5248d-111">Elemento</span><span class="sxs-lookup"><span data-stu-id="5248d-111">Element</span></span> |  <span data-ttu-id="5248d-112">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="5248d-112">Required</span></span>  |  <span data-ttu-id="5248d-113">Descrição</span><span class="sxs-lookup"><span data-stu-id="5248d-113">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5248d-114">Group</span><span class="sxs-lookup"><span data-stu-id="5248d-114">Group</span></span>](group.md)      | <span data-ttu-id="5248d-115">Não</span><span class="sxs-lookup"><span data-stu-id="5248d-115">No</span></span> |  <span data-ttu-id="5248d-116">Define um grupo de comandos</span><span class="sxs-lookup"><span data-stu-id="5248d-116">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="5248d-117">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="5248d-117">OfficeGroup</span></span>](#officegroup)      | <span data-ttu-id="5248d-118">Não</span><span class="sxs-lookup"><span data-stu-id="5248d-118">No</span></span> |  <span data-ttu-id="5248d-119">Representa um grupo de controles integrado do Office.</span><span class="sxs-lookup"><span data-stu-id="5248d-119">Represents a built-in Office control group.</span></span> <span data-ttu-id="5248d-120">**Importante:** não disponível no Outlook.</span><span class="sxs-lookup"><span data-stu-id="5248d-120">**Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="5248d-121">Label</span><span class="sxs-lookup"><span data-stu-id="5248d-121">Label</span></span>](#label-tab)      | <span data-ttu-id="5248d-122">Sim</span><span class="sxs-lookup"><span data-stu-id="5248d-122">Yes</span></span> |  <span data-ttu-id="5248d-123">O rótulo para CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="5248d-123">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="5248d-124">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="5248d-124">InsertAfter</span></span>](#insertafter)      | <span data-ttu-id="5248d-125">Não</span><span class="sxs-lookup"><span data-stu-id="5248d-125">No</span></span> |  <span data-ttu-id="5248d-126">Especifica que a guia personalizada deve estar imediatamente após uma guia do Office. **Importante:** não disponível no Outlook.</span><span class="sxs-lookup"><span data-stu-id="5248d-126">Specifies that the custom tab should be immediately after a specified built-in Office tab. **Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="5248d-127">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="5248d-127">InsertBefore</span></span>](#insertbefore)      | <span data-ttu-id="5248d-128">Não</span><span class="sxs-lookup"><span data-stu-id="5248d-128">No</span></span> |  <span data-ttu-id="5248d-129">Especifica que a guia personalizada deve estar imediatamente antes de uma guia do Office. **Importante:** não disponível no Outlook.</span><span class="sxs-lookup"><span data-stu-id="5248d-129">Specifies that the custom tab should be immediately before a specified built-in Office tab. **Important**: Not available in Outlook.</span></span> |
|  [<span data-ttu-id="5248d-130">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="5248d-130">OverriddenByRibbonApi</span></span>](overriddenbyribbonapi.md)      | <span data-ttu-id="5248d-131">Não</span><span class="sxs-lookup"><span data-stu-id="5248d-131">No</span></span> |  <span data-ttu-id="5248d-132">Especifica se a guia personalizada deve aparecer em combinações de aplicativos e plataformas que suportam guias contextuais personalizadas.</span><span class="sxs-lookup"><span data-stu-id="5248d-132">Specifies whether the custom tab should appear on application and platform combinations that support custom contextual tabs.</span></span> <span data-ttu-id="5248d-133">**Importante:** não disponível no Outlook.</span><span class="sxs-lookup"><span data-stu-id="5248d-133">**Important**: Not available in Outlook.</span></span> |

### <a name="group"></a><span data-ttu-id="5248d-134">Grupo</span><span class="sxs-lookup"><span data-stu-id="5248d-134">Group</span></span>

<span data-ttu-id="5248d-135">Opcional, mas se não estiver presente, deve haver pelo menos um **elemento OfficeGroup.**</span><span class="sxs-lookup"><span data-stu-id="5248d-135">Optional, but if not present there must be at least one **OfficeGroup** element.</span></span> <span data-ttu-id="5248d-136">Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="5248d-136">See [Group element](group.md).</span></span> <span data-ttu-id="5248d-137">A ordem do **Grupo** e **do OfficeGroup** no manifesto deve ser a ordem em que você deseja que apareçam na guia personalizada. Eles podem ser intercalados se houver vários elementos, mas todos devem estar acima do **elemento Label.**</span><span class="sxs-lookup"><span data-stu-id="5248d-137">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

### <a name="officegroup"></a><span data-ttu-id="5248d-138">OfficeGroup</span><span class="sxs-lookup"><span data-stu-id="5248d-138">OfficeGroup</span></span>

<span data-ttu-id="5248d-139">Opcional, mas se não estiver presente, deve haver pelo menos um **elemento Group.**</span><span class="sxs-lookup"><span data-stu-id="5248d-139">Optional, but if not present there must be at least one **Group** element.</span></span> <span data-ttu-id="5248d-140">Representa um grupo de controles integrado do Office.</span><span class="sxs-lookup"><span data-stu-id="5248d-140">Represents a built-in Office control group.</span></span> <span data-ttu-id="5248d-141">O **atributo id** especifica a ID do grupo do Office integrado.</span><span class="sxs-lookup"><span data-stu-id="5248d-141">The **id** attribute specifies the ID of the built-in Office group.</span></span> <span data-ttu-id="5248d-142">Para encontrar a ID de um grupo integrado, consulte [Encontrar as IDs de controles e grupos de controles.](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)</span><span class="sxs-lookup"><span data-stu-id="5248d-142">To find the ID of a built-in group, see [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).</span></span> <span data-ttu-id="5248d-143">A ordem do **Grupo** e **do OfficeGroup** no manifesto deve ser a ordem em que você deseja que apareçam na guia personalizada. Eles podem ser intercalados se houver vários elementos, mas todos devem estar acima do **elemento Label.**</span><span class="sxs-lookup"><span data-stu-id="5248d-143">The order of **Group** and **OfficeGroup** in the manifest should be the order you want them to appear on the custom tab. They can be intermingled if there are multiple elements, but all must be above the **Label** element.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5248d-144">O `OfficeGroup` elemento não está disponível no Outlook.</span><span class="sxs-lookup"><span data-stu-id="5248d-144">The `OfficeGroup` element is not available in Outlook.</span></span>

### <a name="label-tab"></a><span data-ttu-id="5248d-145">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="5248d-145">Label (Tab)</span></span>

<span data-ttu-id="5248d-146">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="5248d-146">Required.</span></span> <span data-ttu-id="5248d-147">O rótulo da guia personalizada. O **atributo resid** pode ter no máximo 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no elemento [Resources.](resources.md)</span><span class="sxs-lookup"><span data-stu-id="5248d-147">The label of the custom tab. The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="insertafter"></a><span data-ttu-id="5248d-148">InsertAfter</span><span class="sxs-lookup"><span data-stu-id="5248d-148">InsertAfter</span></span>

<span data-ttu-id="5248d-149">Opcional.</span><span class="sxs-lookup"><span data-stu-id="5248d-149">Optional.</span></span> <span data-ttu-id="5248d-150">Especifica que a guia personalizada deve ser imediatamente após uma guia do Office. O valor do elemento é a ID da guia integrado, como "TabHome" ou "TabReview".</span><span class="sxs-lookup"><span data-stu-id="5248d-150">Specifies that the custom tab should be immediately after a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="5248d-151">(Consulte [Encontrar as IDs de controles e grupos de controles.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups) Se presente, deve ser após o **elemento Label.**</span><span class="sxs-lookup"><span data-stu-id="5248d-151">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).) If present, must be after the **Label** element.</span></span> <span data-ttu-id="5248d-152">You cannot have both **InsertAfter** and **InsertBefore**.</span><span class="sxs-lookup"><span data-stu-id="5248d-152">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5248d-153">O `InsertAfter` elemento não está disponível no Outlook.</span><span class="sxs-lookup"><span data-stu-id="5248d-153">The `InsertAfter` element is not available in Outlook.</span></span>

### <a name="insertbefore"></a><span data-ttu-id="5248d-154">InsertBefore</span><span class="sxs-lookup"><span data-stu-id="5248d-154">InsertBefore</span></span>

<span data-ttu-id="5248d-155">Opcional.</span><span class="sxs-lookup"><span data-stu-id="5248d-155">Optional.</span></span> <span data-ttu-id="5248d-156">Especifica que a guia personalizada deve estar imediatamente antes de uma guia do Office. O valor do elemento é a ID da guia integrado, como "TabHome" ou "TabReview".</span><span class="sxs-lookup"><span data-stu-id="5248d-156">Specifies that the custom tab should be immediately before a specified built-in Office tab. The value of the element is the ID of the built-in tab, such as "TabHome" or "TabReview".</span></span> <span data-ttu-id="5248d-157">(Consulte [Encontrar as IDs de controles e grupos de controles.)](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups)  Se presente, deve ser após o **elemento Label.**</span><span class="sxs-lookup"><span data-stu-id="5248d-157">(See [Find the IDs of controls and control groups](../../design/built-in-button-integration.md#find-the-ids-of-controls-and-control-groups).)  If present, must be after the **Label** element.</span></span> <span data-ttu-id="5248d-158">You cannot have both **InsertAfter** and **InsertBefore**.</span><span class="sxs-lookup"><span data-stu-id="5248d-158">You cannot have both **InsertAfter** and **InsertBefore**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5248d-159">O `InsertBefore` elemento não está disponível no Outlook.</span><span class="sxs-lookup"><span data-stu-id="5248d-159">The `InsertBefore` element is not available in Outlook.</span></span>

### <a name="overriddenbyribbonapi"></a><span data-ttu-id="5248d-160">OverriddenByRibbonApi</span><span class="sxs-lookup"><span data-stu-id="5248d-160">OverriddenByRibbonApi</span></span>

<span data-ttu-id="5248d-161">Opcional (booliana).</span><span class="sxs-lookup"><span data-stu-id="5248d-161">Optional (boolean).</span></span> <span data-ttu-id="5248d-162">Especifica se a **CustomTab** ficará oculta em combinações de aplicativos e plataformas que suportam uma API que instala uma guia contextual personalizada na faixa de opções em tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="5248d-162">Specifies whether the **CustomTab** will be hidden on application and platform combinations that support an API that installs a custom contextual tab on the ribbon at runtime.</span></span> <span data-ttu-id="5248d-163">O valor padrão, se não estiver presente, é `false` .</span><span class="sxs-lookup"><span data-stu-id="5248d-163">The default value, if not present, is `false`.</span></span> <span data-ttu-id="5248d-164">Se usado, **OverriddenByRibbonApi** deve ser o *primeiro* filho de **CustomTab**.</span><span class="sxs-lookup"><span data-stu-id="5248d-164">If used, **OverriddenByRibbonApi** must be the *first* child of **CustomTab**.</span></span> <span data-ttu-id="5248d-165">Para obter mais informações, [consulte OverriddenByRibbonApi](overriddenbyribbonapi.md).</span><span class="sxs-lookup"><span data-stu-id="5248d-165">For more information, see [OverriddenByRibbonApi](overriddenbyribbonapi.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5248d-166">O `OverriddenByRibbonApi` elemento não está disponível no Outlook.</span><span class="sxs-lookup"><span data-stu-id="5248d-166">The `OverriddenByRibbonApi` element is not available in Outlook.</span></span>

## <a name="customtab-example"></a><span data-ttu-id="5248d-167">Exemplo de CustomTab</span><span class="sxs-lookup"><span data-stu-id="5248d-167">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <CustomTab id="TabCustom1">
    <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
    <Group id="ContosoCustomTab.grp1">
    </Group>
    <OfficeGroup id="Paragraph" />
    <Label resid="customTabLabel1"/>
    <InsertAfter>TabReview</InsertAfter>
  </CustomTab>
</ExtensionPoint>
```
