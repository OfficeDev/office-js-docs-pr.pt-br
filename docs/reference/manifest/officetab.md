---
title: Elemento OfficeTab no arquivo de manifesto
description: O elemento OfficeTab define a guia faixa de opções onde o comando de suplemento é exibido.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 9b07ce1e57329e796545610e0c61a2c11d1ed55d
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641435"
---
# <a name="officetab-element"></a><span data-ttu-id="bb52b-103">Elemento OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bb52b-103">OfficeTab element</span></span>

<span data-ttu-id="bb52b-104">Define a guia da faixa de opções no qual seu comando de suplemento é exibido.</span><span class="sxs-lookup"><span data-stu-id="bb52b-104">Defines the ribbon tab on which your add-in command appears.</span></span> <span data-ttu-id="bb52b-105">Pode ser a guia padrão ( **página inicial**, de **mensagem**ou **reunião**) ou uma guia personalizada definida pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="bb52b-105">This can either be the default tab (either **Home**, **Message**, or **Meeting**), or a custom tab defined by the add-in.</span></span> <span data-ttu-id="bb52b-106">Este elemento é obrigatório.</span><span class="sxs-lookup"><span data-stu-id="bb52b-106">This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="bb52b-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="bb52b-107">Child elements</span></span>

|  <span data-ttu-id="bb52b-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="bb52b-108">Element</span></span> |  <span data-ttu-id="bb52b-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="bb52b-109">Required</span></span>  |  <span data-ttu-id="bb52b-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="bb52b-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bb52b-111">Group</span><span class="sxs-lookup"><span data-stu-id="bb52b-111">Group</span></span>      | <span data-ttu-id="bb52b-112">Sim</span><span class="sxs-lookup"><span data-stu-id="bb52b-112">Yes</span></span> |  <span data-ttu-id="bb52b-p102">Define um grupo de comandos. Você pode adicionar apenas um grupo por suplemento à guia padrão.</span><span class="sxs-lookup"><span data-stu-id="bb52b-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="bb52b-115">A seguir estão os valores válidos de `id` por host.</span><span class="sxs-lookup"><span data-stu-id="bb52b-115">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="bb52b-116">Os valores em **negrito** têm suporte na área de trabalho e online (por exemplo, o Word 2016 ou posterior no Windows e no Word na Web).</span><span class="sxs-lookup"><span data-stu-id="bb52b-116">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word on the web).</span></span>

### <a name="outlook"></a><span data-ttu-id="bb52b-117">Outlook</span><span class="sxs-lookup"><span data-stu-id="bb52b-117">Outlook</span></span>

- <span data-ttu-id="bb52b-118">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="bb52b-118">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="bb52b-119">Word</span><span class="sxs-lookup"><span data-stu-id="bb52b-119">Word</span></span>

- <span data-ttu-id="bb52b-120">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bb52b-120">**TabHome**</span></span>
- <span data-ttu-id="bb52b-121">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bb52b-121">**TabInsert**</span></span>
- <span data-ttu-id="bb52b-122">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="bb52b-122">TabWordDesign</span></span>
- <span data-ttu-id="bb52b-123">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="bb52b-123">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="bb52b-124">TabReferences</span><span class="sxs-lookup"><span data-stu-id="bb52b-124">TabReferences</span></span>
- <span data-ttu-id="bb52b-125">TabMailings</span><span class="sxs-lookup"><span data-stu-id="bb52b-125">TabMailings</span></span>
- <span data-ttu-id="bb52b-126">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="bb52b-126">TabReviewWord</span></span>
- <span data-ttu-id="bb52b-127">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bb52b-127">**TabView**</span></span>
- <span data-ttu-id="bb52b-128">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bb52b-128">TabDeveloper</span></span>
- <span data-ttu-id="bb52b-129">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bb52b-129">TabAddIns</span></span>
- <span data-ttu-id="bb52b-130">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="bb52b-130">TabBlogPost</span></span>
- <span data-ttu-id="bb52b-131">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="bb52b-131">TabBlogInsert</span></span>
- <span data-ttu-id="bb52b-132">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="bb52b-132">TabPrintPreview</span></span>
- <span data-ttu-id="bb52b-133">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="bb52b-133">TabOutlining</span></span>
- <span data-ttu-id="bb52b-134">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="bb52b-134">TabConflicts</span></span>
- <span data-ttu-id="bb52b-135">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="bb52b-135">TabBackgroundRemoval</span></span>
- <span data-ttu-id="bb52b-136">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="bb52b-136">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="bb52b-137">Excel</span><span class="sxs-lookup"><span data-stu-id="bb52b-137">Excel</span></span>

- <span data-ttu-id="bb52b-138">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bb52b-138">**TabHome**</span></span>
- <span data-ttu-id="bb52b-139">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bb52b-139">**TabInsert**</span></span>
- <span data-ttu-id="bb52b-140">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="bb52b-140">TabPageLayoutExcel</span></span>
- <span data-ttu-id="bb52b-141">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="bb52b-141">TabFormulas</span></span>
- <span data-ttu-id="bb52b-142">**TabData**</span><span class="sxs-lookup"><span data-stu-id="bb52b-142">**TabData**</span></span>
- <span data-ttu-id="bb52b-143">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="bb52b-143">**TabReview**</span></span>
- <span data-ttu-id="bb52b-144">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bb52b-144">**TabView**</span></span>
- <span data-ttu-id="bb52b-145">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bb52b-145">TabDeveloper</span></span>
- <span data-ttu-id="bb52b-146">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bb52b-146">TabAddIns</span></span>
- <span data-ttu-id="bb52b-147">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="bb52b-147">TabPrintPreview</span></span>
- <span data-ttu-id="bb52b-148">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="bb52b-148">TabBackgroundRemoval</span></span>

### <a name="powerpoint"></a><span data-ttu-id="bb52b-149">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bb52b-149">PowerPoint</span></span>

- <span data-ttu-id="bb52b-150">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bb52b-150">**TabHome**</span></span>
- <span data-ttu-id="bb52b-151">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bb52b-151">**TabInsert**</span></span>
- <span data-ttu-id="bb52b-152">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="bb52b-152">**TabDesign**</span></span>
- <span data-ttu-id="bb52b-153">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="bb52b-153">**TabTransitions**</span></span>
- <span data-ttu-id="bb52b-154">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="bb52b-154">**TabAnimations**</span></span>
- <span data-ttu-id="bb52b-155">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="bb52b-155">TabSlideShow</span></span>
- <span data-ttu-id="bb52b-156">TabReview</span><span class="sxs-lookup"><span data-stu-id="bb52b-156">TabReview</span></span>
- <span data-ttu-id="bb52b-157">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bb52b-157">**TabView**</span></span>
- <span data-ttu-id="bb52b-158">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bb52b-158">TabDeveloper</span></span>
- <span data-ttu-id="bb52b-159">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bb52b-159">TabAddIns</span></span>
- <span data-ttu-id="bb52b-160">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="bb52b-160">TabPrintPreview</span></span>
- <span data-ttu-id="bb52b-161">TabMerge</span><span class="sxs-lookup"><span data-stu-id="bb52b-161">TabMerge</span></span>
- <span data-ttu-id="bb52b-162">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="bb52b-162">TabGrayscale</span></span>
- <span data-ttu-id="bb52b-163">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="bb52b-163">TabBlackAndWhite</span></span>
- <span data-ttu-id="bb52b-164">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="bb52b-164">TabBroadcastPresentation</span></span>
- <span data-ttu-id="bb52b-165">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="bb52b-165">TabSlideMaster</span></span>
- <span data-ttu-id="bb52b-166">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="bb52b-166">TabHandoutMaster</span></span>
- <span data-ttu-id="bb52b-167">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="bb52b-167">TabNotesMaster</span></span>
- <span data-ttu-id="bb52b-168">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="bb52b-168">TabBackgroundRemoval</span></span>
- <span data-ttu-id="bb52b-169">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="bb52b-169">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="bb52b-170">OneNote</span><span class="sxs-lookup"><span data-stu-id="bb52b-170">OneNote</span></span>

- <span data-ttu-id="bb52b-171">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bb52b-171">**TabHome**</span></span>
- <span data-ttu-id="bb52b-172">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bb52b-172">**TabInsert**</span></span>
- <span data-ttu-id="bb52b-173">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bb52b-173">**TabView**</span></span>
- <span data-ttu-id="bb52b-174">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bb52b-174">TabDeveloper</span></span>
- <span data-ttu-id="bb52b-175">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bb52b-175">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="bb52b-176">Group</span><span class="sxs-lookup"><span data-stu-id="bb52b-176">Group</span></span>

<span data-ttu-id="bb52b-177">Um grupo de pontos de extensão da interface do usuário em uma guia. Um grupo pode ter até seis controles.</span><span class="sxs-lookup"><span data-stu-id="bb52b-177">A group of UI extension points in a tab. A group can have up to six controls.</span></span> <span data-ttu-id="bb52b-178">O atributo **ID** é obrigatório e cada **ID** deve ser exclusivo no manifesto.</span><span class="sxs-lookup"><span data-stu-id="bb52b-178">The **id** attribute is required and each **id** must be unique within the manifest.</span></span> <span data-ttu-id="bb52b-179">A **ID** é uma cadeia de caracteres com, no máximo, 125 caracteres.</span><span class="sxs-lookup"><span data-stu-id="bb52b-179">The **id** is a string with a maximum of 125 characters.</span></span> <span data-ttu-id="bb52b-180">Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="bb52b-180">See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="bb52b-181">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bb52b-181">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
