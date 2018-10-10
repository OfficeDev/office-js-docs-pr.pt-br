# <a name="group-element"></a><span data-ttu-id="11d4e-101">Elemento Group</span><span class="sxs-lookup"><span data-stu-id="11d4e-101">Group element</span></span>

<span data-ttu-id="11d4e-p101">Define um grupo de controles de interface do usuário em uma guia. Em guias personalizadas, o suplemento pode criar até 10 grupos. Cada grupo é limitado a 6 controles, independentemente de qual guia aparece. Suplementos são limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="11d4e-p101">Defines a group of UI controls in a tab.  On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

## <a name="attributes"></a><span data-ttu-id="11d4e-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="11d4e-105">Attributes</span></span>

|  <span data-ttu-id="11d4e-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="11d4e-106">Attribute</span></span>  |  <span data-ttu-id="11d4e-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="11d4e-107">Required</span></span>  |  <span data-ttu-id="11d4e-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="11d4e-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="11d4e-109">id</span><span class="sxs-lookup"><span data-stu-id="11d4e-109">id</span></span>](#id-attribute)  |  <span data-ttu-id="11d4e-110">Sim</span><span class="sxs-lookup"><span data-stu-id="11d4e-110">Yes</span></span>  | <span data-ttu-id="11d4e-111">Uma ID exclusiva do grupo.</span><span class="sxs-lookup"><span data-stu-id="11d4e-111">A unique ID for the group.</span></span>|

### <a name="id-attribute"></a><span data-ttu-id="11d4e-112">Atributo id</span><span class="sxs-lookup"><span data-stu-id="11d4e-112">id attribute</span></span>

<span data-ttu-id="11d4e-p102">Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Esse valor deve ser exclusivo dentro o manifesto, ou o grupo não será processado.</span><span class="sxs-lookup"><span data-stu-id="11d4e-p102">Required. Unique identifier for the group. It is a string with a maximum of 125 characters. This must be unique within the manifest or the group will fail to render.</span></span>

## <a name="child-elements"></a><span data-ttu-id="11d4e-117">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="11d4e-117">Child elements</span></span>
|  <span data-ttu-id="11d4e-118">Elemento</span><span class="sxs-lookup"><span data-stu-id="11d4e-118">Element</span></span> |  <span data-ttu-id="11d4e-119">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="11d4e-119">Required</span></span>  |  <span data-ttu-id="11d4e-120">Descrição</span><span class="sxs-lookup"><span data-stu-id="11d4e-120">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="11d4e-121">Rótulo</span><span class="sxs-lookup"><span data-stu-id="11d4e-121">Label</span></span>](#label)      | <span data-ttu-id="11d4e-122">Sim</span><span class="sxs-lookup"><span data-stu-id="11d4e-122">Yes</span></span> |  <span data-ttu-id="11d4e-123">O rótulo para a CustomTab ou um grupo.</span><span class="sxs-lookup"><span data-stu-id="11d4e-123">The label for the CustomTab or a group.</span></span>  |
|  [<span data-ttu-id="11d4e-124">Control</span><span class="sxs-lookup"><span data-stu-id="11d4e-124">Control</span></span>](#control)    | <span data-ttu-id="11d4e-125">Sim</span><span class="sxs-lookup"><span data-stu-id="11d4e-125">Yes</span></span> |  <span data-ttu-id="11d4e-126">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="11d4e-126">Collection of one or more Control objects.</span></span>  |

### <a name="label"></a><span data-ttu-id="11d4e-127">Rótulo</span><span class="sxs-lookup"><span data-stu-id="11d4e-127">Label</span></span> 

<span data-ttu-id="11d4e-p103">Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="11d4e-p103">Required. The label of the group. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="control"></a><span data-ttu-id="11d4e-131">Control</span><span class="sxs-lookup"><span data-stu-id="11d4e-131">Control</span></span>
<span data-ttu-id="11d4e-132">Um grupo exige pelo menos um controle.</span><span class="sxs-lookup"><span data-stu-id="11d4e-132">A group requires at least one control.</span></span>

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```