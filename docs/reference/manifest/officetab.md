# <a name="officetab-element"></a><span data-ttu-id="1bbe5-101">Elemento OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1bbe5-101">OfficeTab element</span></span>

<span data-ttu-id="1bbe5-p101">Define a guia da faixa de opções no qual seu comando de suplemento é exibido. Pode ser a guia padrão (**Início**, **Mensagem** ou **Reunião**) ou uma guia personalizada definida pelo suplemento. Este elemento é obrigatório.</span><span class="sxs-lookup"><span data-stu-id="1bbe5-p101">Defines the ribbon tab on which your add-in command appears. This can either be the default tab (either  **Home**,  **Message**, or  **Meeting**), or a custom tab defined by the add-in. This element is required.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1bbe5-105">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="1bbe5-105">Child elements</span></span>

|  <span data-ttu-id="1bbe5-106">Elemento</span><span class="sxs-lookup"><span data-stu-id="1bbe5-106">Element</span></span> |  <span data-ttu-id="1bbe5-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="1bbe5-107">Required</span></span>  |  <span data-ttu-id="1bbe5-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="1bbe5-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="1bbe5-109">Grupo</span><span class="sxs-lookup"><span data-stu-id="1bbe5-109">Group</span></span>      | <span data-ttu-id="1bbe5-110">Sim</span><span class="sxs-lookup"><span data-stu-id="1bbe5-110">Yes</span></span> |  <span data-ttu-id="1bbe5-p102">Define um grupo de comandos. Você pode adicionar apenas um grupo por suplemento à guia padrão.</span><span class="sxs-lookup"><span data-stu-id="1bbe5-p102">Defines a group of commands. You can add only one group per add-in to the default tab.</span></span>  |

<span data-ttu-id="1bbe5-113">A seguir estão os valores válidos da guia `id` por host.</span><span class="sxs-lookup"><span data-stu-id="1bbe5-113">The following are valid tab `id` values by host.</span></span> <span data-ttu-id="1bbe5-114">Os valores em **negrito** são suportados na área de trabalho e on-line (por exemplo, Word 2016 ou posterior para Windows e Word Online).</span><span class="sxs-lookup"><span data-stu-id="1bbe5-114">Values in **bold** are supported in both desktop and online (for example, Word 2016 for Windows and Word Online).</span></span>

### <a name="outlook"></a><span data-ttu-id="1bbe5-115">Outlook</span><span class="sxs-lookup"><span data-stu-id="1bbe5-115">Outlook</span></span>

- <span data-ttu-id="1bbe5-116">**TabDefault**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-116">**TabDefault**</span></span>

### <a name="word"></a><span data-ttu-id="1bbe5-117">Word</span><span class="sxs-lookup"><span data-stu-id="1bbe5-117">Word</span></span>

- <span data-ttu-id="1bbe5-118">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-118">**TabHome**</span></span>
- <span data-ttu-id="1bbe5-119">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-119">**TabInsert**</span></span>
- <span data-ttu-id="1bbe5-120">TabWordDesign</span><span class="sxs-lookup"><span data-stu-id="1bbe5-120">TabWordDesign</span></span>
- <span data-ttu-id="1bbe5-121">**TabPageLayoutWord**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-121">**TabPageLayoutWord**</span></span>
- <span data-ttu-id="1bbe5-122">TabReferences</span><span class="sxs-lookup"><span data-stu-id="1bbe5-122">TabReferences</span></span>
- <span data-ttu-id="1bbe5-123">TabMailings</span><span class="sxs-lookup"><span data-stu-id="1bbe5-123">TabMailings</span></span>
- <span data-ttu-id="1bbe5-124">TabReviewWord</span><span class="sxs-lookup"><span data-stu-id="1bbe5-124">TabReviewWord</span></span>
- <span data-ttu-id="1bbe5-125">**TabView**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-125">**TabView**</span></span>
- <span data-ttu-id="1bbe5-126">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="1bbe5-126">TabDeveloper</span></span>
- <span data-ttu-id="1bbe5-127">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="1bbe5-127">TabAddIns</span></span>
- <span data-ttu-id="1bbe5-128">TabBlogPost</span><span class="sxs-lookup"><span data-stu-id="1bbe5-128">TabBlogPost</span></span>
- <span data-ttu-id="1bbe5-129">TabBlogInsert</span><span class="sxs-lookup"><span data-stu-id="1bbe5-129">TabBlogInsert</span></span>
- <span data-ttu-id="1bbe5-130">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="1bbe5-130">TabPrintPreview</span></span>
- <span data-ttu-id="1bbe5-131">TabOutlining</span><span class="sxs-lookup"><span data-stu-id="1bbe5-131">TabOutlining</span></span>
- <span data-ttu-id="1bbe5-132">TabConflicts</span><span class="sxs-lookup"><span data-stu-id="1bbe5-132">TabConflicts</span></span>
- <span data-ttu-id="1bbe5-133">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="1bbe5-133">TabBackgroundRemoval</span></span>
- <span data-ttu-id="1bbe5-134">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="1bbe5-134">TabBroadcastPresentation</span></span>

### <a name="excel"></a><span data-ttu-id="1bbe5-135">Excel</span><span class="sxs-lookup"><span data-stu-id="1bbe5-135">Excel</span></span>

- <span data-ttu-id="1bbe5-136">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-136">**TabHome**</span></span>
- <span data-ttu-id="1bbe5-137">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-137">**TabInsert**</span></span>
- <span data-ttu-id="1bbe5-138">TabPageLayoutExcel</span><span class="sxs-lookup"><span data-stu-id="1bbe5-138">TabPageLayoutExcel</span></span>
- <span data-ttu-id="1bbe5-139">TabFormulas</span><span class="sxs-lookup"><span data-stu-id="1bbe5-139">TabFormulas</span></span>
- <span data-ttu-id="1bbe5-140">**TabData**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-140">**TabData**</span></span>
- <span data-ttu-id="1bbe5-141">**TabReview**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-141">**TabReview**</span></span>
- <span data-ttu-id="1bbe5-142">**TabView**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-142">**TabView**</span></span>
- <span data-ttu-id="1bbe5-143">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="1bbe5-143">TabDeveloper</span></span>
- <span data-ttu-id="1bbe5-144">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="1bbe5-144">TabAddIns</span></span>
- <span data-ttu-id="1bbe5-145">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="1bbe5-145">TabPrintPreview</span></span>
- <span data-ttu-id="1bbe5-146">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="1bbe5-146">TabBackgroundRemoval</span></span> 

### <a name="powerpoint"></a><span data-ttu-id="1bbe5-147">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1bbe5-147">PowerPoint</span></span>

- <span data-ttu-id="1bbe5-148">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-148">**TabHome**</span></span>
- <span data-ttu-id="1bbe5-149">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-149">**TabInsert**</span></span>
- <span data-ttu-id="1bbe5-150">**TabDesign**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-150">**TabDesign**</span></span>
- <span data-ttu-id="1bbe5-151">**TabTransitions**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-151">**TabTransitions**</span></span>
- <span data-ttu-id="1bbe5-152">**TabAnimations**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-152">**TabAnimations**</span></span>
- <span data-ttu-id="1bbe5-153">TabSlideShow</span><span class="sxs-lookup"><span data-stu-id="1bbe5-153">TabSlideShow</span></span>
- <span data-ttu-id="1bbe5-154">TabReview</span><span class="sxs-lookup"><span data-stu-id="1bbe5-154">TabReview</span></span>
- <span data-ttu-id="1bbe5-155">**TabView**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-155">**TabView**</span></span>
- <span data-ttu-id="1bbe5-156">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="1bbe5-156">TabDeveloper</span></span>
- <span data-ttu-id="1bbe5-157">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="1bbe5-157">TabAddIns</span></span>
- <span data-ttu-id="1bbe5-158">TabPrintPreview</span><span class="sxs-lookup"><span data-stu-id="1bbe5-158">TabPrintPreview</span></span>
- <span data-ttu-id="1bbe5-159">TabMerge</span><span class="sxs-lookup"><span data-stu-id="1bbe5-159">TabMerge</span></span>
- <span data-ttu-id="1bbe5-160">TabGrayscale</span><span class="sxs-lookup"><span data-stu-id="1bbe5-160">TabGrayscale</span></span>
- <span data-ttu-id="1bbe5-161">TabBlackAndWhite</span><span class="sxs-lookup"><span data-stu-id="1bbe5-161">TabBlackAndWhite</span></span>
- <span data-ttu-id="1bbe5-162">TabBroadcastPresentation</span><span class="sxs-lookup"><span data-stu-id="1bbe5-162">TabBroadcastPresentation</span></span>
- <span data-ttu-id="1bbe5-163">TabSlideMaster</span><span class="sxs-lookup"><span data-stu-id="1bbe5-163">TabSlideMaster</span></span>
- <span data-ttu-id="1bbe5-164">TabHandoutMaster</span><span class="sxs-lookup"><span data-stu-id="1bbe5-164">TabHandoutMaster</span></span>
- <span data-ttu-id="1bbe5-165">TabNotesMaster</span><span class="sxs-lookup"><span data-stu-id="1bbe5-165">TabNotesMaster</span></span>
- <span data-ttu-id="1bbe5-166">TabBackgroundRemoval</span><span class="sxs-lookup"><span data-stu-id="1bbe5-166">TabBackgroundRemoval</span></span>
- <span data-ttu-id="1bbe5-167">TabSlideMasterHome</span><span class="sxs-lookup"><span data-stu-id="1bbe5-167">TabSlideMasterHome</span></span>

### <a name="onenote"></a><span data-ttu-id="1bbe5-168">OneNote</span><span class="sxs-lookup"><span data-stu-id="1bbe5-168">OneNote</span></span>

- <span data-ttu-id="1bbe5-169">**TabHome**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-169">**TabHome**</span></span>
- <span data-ttu-id="1bbe5-170">**TabInsert**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-170">**TabInsert**</span></span>
- <span data-ttu-id="1bbe5-171">**TabView**</span><span class="sxs-lookup"><span data-stu-id="1bbe5-171">**TabView**</span></span>
- <span data-ttu-id="1bbe5-172">TabDeveloper</span><span class="sxs-lookup"><span data-stu-id="1bbe5-172">TabDeveloper</span></span>
- <span data-ttu-id="1bbe5-173">TabAddIns</span><span class="sxs-lookup"><span data-stu-id="1bbe5-173">TabAddIns</span></span>

## <a name="group"></a><span data-ttu-id="1bbe5-174">Grupo</span><span class="sxs-lookup"><span data-stu-id="1bbe5-174">Group</span></span>

<span data-ttu-id="1bbe5-p104">Um grupo de pontos de extensão de interface do usuário em uma guia. Um grupo pode ter até seis controles. O atributo **id** é obrigatório e cada **id** deve ser exclusiva do manifesto. A **id** é uma sequência de caracteres com, no máximo, 125 caracteres. Confira [Elemento Group](group.md)</span><span class="sxs-lookup"><span data-stu-id="1bbe5-p104">A group of UI extension points in a tab. A group can have up to six controls. The  **id** attribute is required and each **id** must be unique within the manifest. The **id** is a string with a maximum of 125 characters. See [Group element](group.md).</span></span>

## <a name="officetab-example"></a><span data-ttu-id="1bbe5-179">Exemplo OfficeTab</span><span class="sxs-lookup"><span data-stu-id="1bbe5-179">OfficeTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgreadTabMessage.grp1">
        <!-- Group Definition -->
    </Group>
  </OfficeTab>
</ExtensionPoint>
```
