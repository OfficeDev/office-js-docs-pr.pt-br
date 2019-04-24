---
title: Elemento OfficeTab no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b61c245c000f8bf13eb71c991ec57a125993c2fc
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450461"
---
# <a name="officetab-element"></a><span data-ttu-id="4e0d9-102">Elemento OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4e0d9-102">OfficeTab element</span></span>

<span data-ttu-id="4e0d9-p101">Define a guia da faixa de opções no qual seu comando de suplemento é exibido. Pode ser a guia padrão (**Início**, **Mensagem** ou **Reunião**) ou uma guia personalizada definida pelo suplemento. Este elemento é obrigatório.</span><span class="sxs-lookup"><span data-stu-id="4e0d9-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="4e0d9-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="4e0d9-106">Child elements</span></span>

|  <span data-ttu-id="4e0d9-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="4e0d9-107">Element</span></span> |  <span data-ttu-id="4e0d9-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="4e0d9-108">Required</span></span>  |  <span data-ttu-id="4e0d9-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="4e0d9-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="4e0d9-110">Group</span><span class="sxs-lookup"><span data-stu-id="4e0d9-110">Group</span></span>      | <span data-ttu-id="4e0d9-111">Sim</span><span class="sxs-lookup"><span data-stu-id="4e0d9-111">Yes</span></span> |  <span data-ttu-id="4e0d9-p102">Define um grupo de comandos. Você pode adicionar apenas um grupo por suplemento à guia padrão.</span><span class="sxs-lookup"><span data-stu-id="4e0d9-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="4e0d9-114">A seguir estão os valores válidos de `id` por host.</span><span class="sxs-lookup"><span data-stu-id="4e0d9-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="4e0d9-115">Os valores em **negrito** têm suporte na área de trabalho e online (por exemplo, Word 2016 ou posterior para Windows e Word Online).</span><span class="sxs-lookup"><span data-stu-id="4e0d9-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="4e0d9-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="4e0d9-116">Outlook</span></span>

- <span data-ttu-id="4e0d9-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="4e0d9-118">Word</span><span class="sxs-lookup"><span data-stu-id="4e0d9-118">Word</span></span>

- <span data-ttu-id="4e0d9-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-119">**TabHome**</span></span>
- <span data-ttu-id="4e0d9-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-120">**TabInsert**</span></span>
- <span data-ttu-id="4e0d9-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="4e0d9-121">TabWordDesign</span></span>
- <span data-ttu-id="4e0d9-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="4e0d9-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="4e0d9-123">TabReferences</span></span>
- <span data-ttu-id="4e0d9-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="4e0d9-124">TabMailings</span></span>
- <span data-ttu-id="4e0d9-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="4e0d9-125">TabReviewWord</span></span>
- <span data-ttu-id="4e0d9-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-126">**TabView**</span></span>
- <span data-ttu-id="4e0d9-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="4e0d9-127">TabDeveloper</span></span>
- <span data-ttu-id="4e0d9-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="4e0d9-128">TabAddIns</span></span>
- <span data-ttu-id="4e0d9-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="4e0d9-129">TabBlogPost</span></span>
- <span data-ttu-id="4e0d9-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="4e0d9-130">TabBlogInsert</span></span>
- <span data-ttu-id="4e0d9-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="4e0d9-131">TabPrintPreview</span></span>
- <span data-ttu-id="4e0d9-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="4e0d9-132">TabOutlining</span></span>
- <span data-ttu-id="4e0d9-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="4e0d9-133">TabConflicts</span></span>
- <span data-ttu-id="4e0d9-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="4e0d9-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="4e0d9-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="4e0d9-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="4e0d9-136">Excel</span><span class="sxs-lookup"><span data-stu-id="4e0d9-136">Excel</span></span>

- <span data-ttu-id="4e0d9-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-137">**TabHome**</span></span>
- <span data-ttu-id="4e0d9-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-138">**TabInsert**</span></span>
- <span data-ttu-id="4e0d9-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="4e0d9-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="4e0d9-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="4e0d9-140">TabFormulas</span></span>
- <span data-ttu-id="4e0d9-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-141">**TabData**</span></span>
- <span data-ttu-id="4e0d9-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-142">**TabReview**</span></span>
- <span data-ttu-id="4e0d9-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-143">**TabView**</span></span>
- <span data-ttu-id="4e0d9-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="4e0d9-144">TabDeveloper</span></span>
- <span data-ttu-id="4e0d9-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="4e0d9-145">TabAddIns</span></span>
- <span data-ttu-id="4e0d9-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="4e0d9-146">TabPrintPreview</span></span>
- <span data-ttu-id="4e0d9-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="4e0d9-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="4e0d9-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4e0d9-148">PowerPoint</span></span>

- <span data-ttu-id="4e0d9-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-149">**TabHome**</span></span>
- <span data-ttu-id="4e0d9-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-150">**TabInsert**</span></span>
- <span data-ttu-id="4e0d9-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-151">**TabDesign**</span></span>
- <span data-ttu-id="4e0d9-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-152">**TabTransitions**</span></span>
- <span data-ttu-id="4e0d9-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-153">**TabAnimations**</span></span>
- <span data-ttu-id="4e0d9-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="4e0d9-154">TabSlideShow</span></span>
- <span data-ttu-id="4e0d9-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="4e0d9-155">TabReview</span></span>
- <span data-ttu-id="4e0d9-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-156">**TabView**</span></span>
- <span data-ttu-id="4e0d9-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="4e0d9-157">TabDeveloper</span></span>
- <span data-ttu-id="4e0d9-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="4e0d9-158">TabAddIns</span></span>
- <span data-ttu-id="4e0d9-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="4e0d9-159">TabPrintPreview</span></span>
- <span data-ttu-id="4e0d9-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="4e0d9-160">TabMerge</span></span>
- <span data-ttu-id="4e0d9-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="4e0d9-161">TabGrayscale</span></span>
- <span data-ttu-id="4e0d9-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="4e0d9-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="4e0d9-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="4e0d9-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="4e0d9-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="4e0d9-164">TabSlideMaster</span></span>
- <span data-ttu-id="4e0d9-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="4e0d9-165">TabHandoutMaster</span></span>
- <span data-ttu-id="4e0d9-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="4e0d9-166">TabNotesMaster</span></span>
- <span data-ttu-id="4e0d9-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="4e0d9-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="4e0d9-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="4e0d9-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="4e0d9-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="4e0d9-169">OneNote</span></span>

- <span data-ttu-id="4e0d9-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-170">**TabHome**</span></span>
- <span data-ttu-id="4e0d9-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-171">**TabInsert**</span></span>
- <span data-ttu-id="4e0d9-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="4e0d9-172">**TabView**</span></span>
- <span data-ttu-id="4e0d9-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="4e0d9-173">TabDeveloper</span></span>
- <span data-ttu-id="4e0d9-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="4e0d9-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="4e0d9-175">Group</span><span class="sxs-lookup"><span data-stu-id="4e0d9-175">Group</span></span>

<span data-ttu-id="4e0d9-p104">Um grupo de pontos de extensão de interface do usuário em uma guia. O grupo pode ter até seis controles. O atributo **id** é obrigatório, e cada **id** deve ser exclusiva dentro do manifesto. A **id** é uma cadeia de caracteres com, no máximo, 125 caracteres. Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="4e0d9-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="4e0d9-180">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="4e0d9-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
