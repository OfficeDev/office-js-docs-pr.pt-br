---
title: Elemento OfficeTab no arquivo de manifesto
description: ''
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: 1bf9f1d1e08a8147b52f93923229ef8fb8556fcf
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952268"
---
# <a name="officetab-element"></a><span data-ttu-id="bcdd3-102">Elemento OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bcdd3-102">OfficeTab element</span></span>

<span data-ttu-id="bcdd3-p101">Define a guia da faixa de opções no qual seu comando de suplemento é exibido. Pode ser a guia padrão (**Início**, **Mensagem** ou **Reunião**) ou uma guia personalizada definida pelo suplemento. Este elemento é obrigatório.</span><span class="sxs-lookup"><span data-stu-id="bcdd3-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="bcdd3-106">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="bcdd3-106">Child elements</span></span>

|  <span data-ttu-id="bcdd3-107">Elemento</span><span class="sxs-lookup"><span data-stu-id="bcdd3-107">Element</span></span> |  <span data-ttu-id="bcdd3-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="bcdd3-108">Required</span></span>  |  <span data-ttu-id="bcdd3-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="bcdd3-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bcdd3-110">Group</span><span class="sxs-lookup"><span data-stu-id="bcdd3-110">Group</span></span>      | <span data-ttu-id="bcdd3-111">Sim</span><span class="sxs-lookup"><span data-stu-id="bcdd3-111">Yes</span></span> |  <span data-ttu-id="bcdd3-p102">Define um grupo de comandos. Você pode adicionar apenas um grupo por suplemento à guia padrão.</span><span class="sxs-lookup"><span data-stu-id="bcdd3-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="bcdd3-114">A seguir estão os valores válidos de `id` por host.</span><span class="sxs-lookup"><span data-stu-id="bcdd3-114">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="bcdd3-115">Os valores em **negrito** têm suporte na área de trabalho e online (por exemplo, Word 2016 ou posterior no Windows e no Word online).</span><span class="sxs-lookup"><span data-stu-id="bcdd3-115">Values in **bold** are supported in both desktop and online (for example, Word 2016 or later on Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="bcdd3-116">Outlook</span><span class="sxs-lookup"><span data-stu-id="bcdd3-116">Outlook</span></span>

- <span data-ttu-id="bcdd3-117">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-117">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="bcdd3-118">Word</span><span class="sxs-lookup"><span data-stu-id="bcdd3-118">Word</span></span>

- <span data-ttu-id="bcdd3-119">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-119">**TabHome**</span></span>
- <span data-ttu-id="bcdd3-120">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-120">**TabInsert**</span></span>
- <span data-ttu-id="bcdd3-121">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="bcdd3-121">TabWordDesign</span></span>
- <span data-ttu-id="bcdd3-122">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-122">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="bcdd3-123">TabReferences</span><span class="sxs-lookup"><span data-stu-id="bcdd3-123">TabReferences</span></span>
- <span data-ttu-id="bcdd3-124">TabMailings</span><span class="sxs-lookup"><span data-stu-id="bcdd3-124">TabMailings</span></span>
- <span data-ttu-id="bcdd3-125">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="bcdd3-125">TabReviewWord</span></span>
- <span data-ttu-id="bcdd3-126">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-126">**TabView**</span></span>
- <span data-ttu-id="bcdd3-127">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bcdd3-127">TabDeveloper</span></span>
- <span data-ttu-id="bcdd3-128">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bcdd3-128">TabAddIns</span></span>
- <span data-ttu-id="bcdd3-129">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="bcdd3-129">TabBlogPost</span></span>
- <span data-ttu-id="bcdd3-130">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="bcdd3-130">TabBlogInsert</span></span>
- <span data-ttu-id="bcdd3-131">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="bcdd3-131">TabPrintPreview</span></span>
- <span data-ttu-id="bcdd3-132">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="bcdd3-132">TabOutlining</span></span>
- <span data-ttu-id="bcdd3-133">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="bcdd3-133">TabConflicts</span></span>
- <span data-ttu-id="bcdd3-134">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="bcdd3-134">TabBackgroundRemoval</span></span>
- <span data-ttu-id="bcdd3-135">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="bcdd3-135">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="bcdd3-136">Excel</span><span class="sxs-lookup"><span data-stu-id="bcdd3-136">Excel</span></span>

- <span data-ttu-id="bcdd3-137">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-137">**TabHome**</span></span>
- <span data-ttu-id="bcdd3-138">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-138">**TabInsert**</span></span>
- <span data-ttu-id="bcdd3-139">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="bcdd3-139">TabPageLayoutExcel</span></span>
- <span data-ttu-id="bcdd3-140">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="bcdd3-140">TabFormulas</span></span>
- <span data-ttu-id="bcdd3-141">**TabData**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-141">**TabData**</span></span>
- <span data-ttu-id="bcdd3-142">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-142">**TabReview**</span></span>
- <span data-ttu-id="bcdd3-143">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-143">**TabView**</span></span>
- <span data-ttu-id="bcdd3-144">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bcdd3-144">TabDeveloper</span></span>
- <span data-ttu-id="bcdd3-145">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bcdd3-145">TabAddIns</span></span>
- <span data-ttu-id="bcdd3-146">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="bcdd3-146">TabPrintPreview</span></span>
- <span data-ttu-id="bcdd3-147">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="bcdd3-147">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="bcdd3-148">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bcdd3-148">PowerPoint</span></span>

- <span data-ttu-id="bcdd3-149">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-149">**TabHome**</span></span>
- <span data-ttu-id="bcdd3-150">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-150">**TabInsert**</span></span>
- <span data-ttu-id="bcdd3-151">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-151">**TabDesign**</span></span>
- <span data-ttu-id="bcdd3-152">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-152">**TabTransitions**</span></span>
- <span data-ttu-id="bcdd3-153">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-153">**TabAnimations**</span></span>
- <span data-ttu-id="bcdd3-154">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="bcdd3-154">TabSlideShow</span></span>
- <span data-ttu-id="bcdd3-155">TabReview</span><span class="sxs-lookup"><span data-stu-id="bcdd3-155">TabReview</span></span>
- <span data-ttu-id="bcdd3-156">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-156">**TabView**</span></span>
- <span data-ttu-id="bcdd3-157">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bcdd3-157">TabDeveloper</span></span>
- <span data-ttu-id="bcdd3-158">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bcdd3-158">TabAddIns</span></span>
- <span data-ttu-id="bcdd3-159">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="bcdd3-159">TabPrintPreview</span></span>
- <span data-ttu-id="bcdd3-160">TabMerge</span><span class="sxs-lookup"><span data-stu-id="bcdd3-160">TabMerge</span></span>
- <span data-ttu-id="bcdd3-161">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="bcdd3-161">TabGrayscale</span></span>
- <span data-ttu-id="bcdd3-162">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="bcdd3-162">TabBlackAndWhite</span></span>
- <span data-ttu-id="bcdd3-163">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="bcdd3-163">TabBroadcastPresentation</span></span>
- <span data-ttu-id="bcdd3-164">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="bcdd3-164">TabSlideMaster</span></span>
- <span data-ttu-id="bcdd3-165">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="bcdd3-165">TabHandoutMaster</span></span>
- <span data-ttu-id="bcdd3-166">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="bcdd3-166">TabNotesMaster</span></span>
- <span data-ttu-id="bcdd3-167">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="bcdd3-167">TabBackgroundRemoval</span></span>
- <span data-ttu-id="bcdd3-168">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="bcdd3-168">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="bcdd3-169">OneNote</span><span class="sxs-lookup"><span data-stu-id="bcdd3-169">OneNote</span></span>

- <span data-ttu-id="bcdd3-170">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-170">**TabHome**</span></span>
- <span data-ttu-id="bcdd3-171">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-171">**TabInsert**</span></span>
- <span data-ttu-id="bcdd3-172">**TabView**</span><span class="sxs-lookup"><span data-stu-id="bcdd3-172">**TabView**</span></span>
- <span data-ttu-id="bcdd3-173">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="bcdd3-173">TabDeveloper</span></span>
- <span data-ttu-id="bcdd3-174">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="bcdd3-174">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="bcdd3-175">Group</span><span class="sxs-lookup"><span data-stu-id="bcdd3-175">Group</span></span>

<span data-ttu-id="bcdd3-p104">Um grupo de pontos de extensão de interface do usuário em uma guia. O grupo pode ter até seis controles. O atributo **id** é obrigatório, e cada **id** deve ser exclusiva dentro do manifesto. A **id** é uma cadeia de caracteres com, no máximo, 125 caracteres. Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="bcdd3-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="bcdd3-180">Exemplo de OfficeTab</span><span class="sxs-lookup"><span data-stu-id="bcdd3-180">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
