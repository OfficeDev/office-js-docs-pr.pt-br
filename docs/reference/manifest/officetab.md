---
title: Elemento OfficeTab no arquivo de manifesto
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: b8458233ba93e98fe0bd8d51f5734b1fece65864
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324831"
---
# <a name="officetab-element"></a><span data-ttu-id="431bb-102">Elemento OfficeTab</span><span class="sxs-lookup"><span data-stu-id="431bb-102">OfficeTab element</span></span>

<span data-ttu-id="431bb-103">Define a guia da faixa de opções no qual seu comando de suplemento é exibido.</span><span class="sxs-lookup"><span data-stu-id="431bb-103">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="431bb-104">Pode ser a guia padrão ( **página inicial**, de **mensagem**ou **reunião**) ou uma guia personalizada definida pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="431bb-104">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="431bb-105">Este elemento é obrigatório.</span><span class="sxs-lookup"><span data-stu-id="431bb-105">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="431bb-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="431bb-106">Child elements</span></span>

|  <span data-ttu-id="431bb-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="431bb-107">Element</span></span> |  <span data-ttu-id="431bb-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="431bb-108">Required</span></span>  |  <span data-ttu-id="431bb-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="431bb-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="431bb-110">Group</span><span class="sxs-lookup"><span data-stu-id="431bb-110">Group</span></span>      | <span data-ttu-id="431bb-111">Sim</span><span class="sxs-lookup"><span data-stu-id="431bb-111">Yes</span></span> |  <span data-ttu-id="431bb-p102">Define um grupo de comandos. Você pode adicionar apenas um grupo por suplemento à guia padrão.</span><span class="sxs-lookup"><span data-stu-id="431bb-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="431bb-114">A seguir estão os valores válidos de `id` por host.</span><span class="sxs-lookup"><span data-stu-id="431bb-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="431bb-115">Os valores em **negrito** têm suporte na área de trabalho e online (por exemplo, o Word 2016 ou posterior no Windows e no Word na Web).</span><span class="sxs-lookup"><span data-stu-id="431bb-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="431bb-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="431bb-116">Outlook</span></span>

- <span data-ttu-id="431bb-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="431bb-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="431bb-118">Word</span><span class="sxs-lookup"><span data-stu-id="431bb-118">Word</span></span>

- <span data-ttu-id="431bb-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="431bb-119">**TabHome**</span></span>
- <span data-ttu-id="431bb-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="431bb-120">**TabInsert**</span></span>
- <span data-ttu-id="431bb-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="431bb-121">TabWordDesign</span></span>
- <span data-ttu-id="431bb-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="431bb-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="431bb-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="431bb-123">TabReferences</span></span>
- <span data-ttu-id="431bb-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="431bb-124">TabMailings</span></span>
- <span data-ttu-id="431bb-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="431bb-125">TabReviewWord</span></span>
- <span data-ttu-id="431bb-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="431bb-126">**TabView**</span></span>
- <span data-ttu-id="431bb-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="431bb-127">TabDeveloper</span></span>
- <span data-ttu-id="431bb-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="431bb-128">TabAddIns</span></span>
- <span data-ttu-id="431bb-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="431bb-129">TabBlogPost</span></span>
- <span data-ttu-id="431bb-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="431bb-130">TabBlogInsert</span></span>
- <span data-ttu-id="431bb-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="431bb-131">TabPrintPreview</span></span>
- <span data-ttu-id="431bb-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="431bb-132">TabOutlining</span></span>
- <span data-ttu-id="431bb-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="431bb-133">TabConflicts</span></span>
- <span data-ttu-id="431bb-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="431bb-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="431bb-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="431bb-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="431bb-136">Excel</span><span class="sxs-lookup"><span data-stu-id="431bb-136">Excel</span></span>

- <span data-ttu-id="431bb-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="431bb-137">**TabHome**</span></span>
- <span data-ttu-id="431bb-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="431bb-138">**TabInsert**</span></span>
- <span data-ttu-id="431bb-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="431bb-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="431bb-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="431bb-140">TabFormulas</span></span>
- <span data-ttu-id="431bb-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="431bb-141">**TabData**</span></span>
- <span data-ttu-id="431bb-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="431bb-142">**TabReview**</span></span>
- <span data-ttu-id="431bb-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="431bb-143">**TabView**</span></span>
- <span data-ttu-id="431bb-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="431bb-144">TabDeveloper</span></span>
- <span data-ttu-id="431bb-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="431bb-145">TabAddIns</span></span>
- <span data-ttu-id="431bb-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="431bb-146">TabPrintPreview</span></span>
- <span data-ttu-id="431bb-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="431bb-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="431bb-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="431bb-148">PowerPoint</span></span>

- <span data-ttu-id="431bb-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="431bb-149">**TabHome**</span></span>
- <span data-ttu-id="431bb-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="431bb-150">**TabInsert**</span></span>
- <span data-ttu-id="431bb-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="431bb-151">**TabDesign**</span></span>
- <span data-ttu-id="431bb-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="431bb-152">**TabTransitions**</span></span>
- <span data-ttu-id="431bb-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="431bb-153">**TabAnimations**</span></span>
- <span data-ttu-id="431bb-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="431bb-154">TabSlideShow</span></span>
- <span data-ttu-id="431bb-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="431bb-155">TabReview</span></span>
- <span data-ttu-id="431bb-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="431bb-156">**TabView**</span></span>
- <span data-ttu-id="431bb-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="431bb-157">TabDeveloper</span></span>
- <span data-ttu-id="431bb-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="431bb-158">TabAddIns</span></span>
- <span data-ttu-id="431bb-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="431bb-159">TabPrintPreview</span></span>
- <span data-ttu-id="431bb-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="431bb-160">TabMerge</span></span>
- <span data-ttu-id="431bb-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="431bb-161">TabGrayscale</span></span>
- <span data-ttu-id="431bb-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="431bb-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="431bb-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="431bb-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="431bb-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="431bb-164">TabSlideMaster</span></span>
- <span data-ttu-id="431bb-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="431bb-165">TabHandoutMaster</span></span>
- <span data-ttu-id="431bb-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="431bb-166">TabNotesMaster</span></span>
- <span data-ttu-id="431bb-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="431bb-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="431bb-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="431bb-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="431bb-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="431bb-169">OneNote</span></span>

- <span data-ttu-id="431bb-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="431bb-170">**TabHome**</span></span>
- <span data-ttu-id="431bb-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="431bb-171">**TabInsert**</span></span>
- <span data-ttu-id="431bb-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="431bb-172">**TabView**</span></span>
- <span data-ttu-id="431bb-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="431bb-173">TabDeveloper</span></span>
- <span data-ttu-id="431bb-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="431bb-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="431bb-175">Group</span><span class="sxs-lookup"><span data-stu-id="431bb-175">Group</span></span>

<span data-ttu-id="431bb-176">Um grupo de pontos de extensão da interface do usuário em uma guia. Um grupo pode ter até seis controles.</span><span class="sxs-lookup"><span data-stu-id="431bb-176">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="431bb-177">O atributo **ID** é obrigatório e cada **ID** deve ser exclusivo no manifesto.</span><span class="sxs-lookup"><span data-stu-id="431bb-177">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="431bb-178">A **ID** é uma cadeia de caracteres com, no máximo, 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="431bb-178">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="431bb-179">Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="431bb-179">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="431bb-180">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="431bb-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
