# <a name="customtab-element"></a><span data-ttu-id="535f2-101">Elemento CustomTab</span><span class="sxs-lookup"><span data-stu-id="535f2-101">CustomTab element</span></span>

<span data-ttu-id="535f2-p101">Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento. Isso pode ser realizado na guia padrão (**Página inicial**, **Mensagem** ou **Reunião**) ou em uma guia personalizada definida pelo suplemento.</span><span class="sxs-lookup"><span data-stu-id="535f2-p101">On the ribbon, you specify which tab and group for their add-in commands. This can either be on the default tab (either  **Home**,  **Message**, or  **Meeting**), or on a custom tab defined by the add-in.</span></span>

<span data-ttu-id="535f2-p102">Nas guias personalizadas, o suplemento poderá criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual aparece. Os suplementos estão limitados a uma guia personalizada.</span><span class="sxs-lookup"><span data-stu-id="535f2-p102">On custom tabs, the add-in can create up to 10 groups. Each group is limited to 6 controls, regardless of which tab it appears on. Add-ins are limited to one custom tab.</span></span>

<span data-ttu-id="535f2-107">O atributo **id** deve ser único dentro do manifesto.</span><span class="sxs-lookup"><span data-stu-id="535f2-107">The  **id** attribute must be unique within the manifest.</span></span>

## <a name="child-elements"></a><span data-ttu-id="535f2-108">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="535f2-108">Child elements</span></span>

|  <span data-ttu-id="535f2-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="535f2-109">Element</span></span> |  <span data-ttu-id="535f2-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="535f2-110">Required</span></span>  |  <span data-ttu-id="535f2-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="535f2-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="535f2-112">Group</span><span class="sxs-lookup"><span data-stu-id="535f2-112">Group</span></span>](group.md)      | <span data-ttu-id="535f2-113">Sim</span><span class="sxs-lookup"><span data-stu-id="535f2-113">Yes</span></span> |  <span data-ttu-id="535f2-114">Define um grupo de comandos</span><span class="sxs-lookup"><span data-stu-id="535f2-114">Defines a Group of commands.</span></span>  |
|  [<span data-ttu-id="535f2-115">Label</span><span class="sxs-lookup"><span data-stu-id="535f2-115">Label</span></span>](#label-tab)      | <span data-ttu-id="535f2-116">Sim</span><span class="sxs-lookup"><span data-stu-id="535f2-116">Yes</span></span> |  <span data-ttu-id="535f2-117">O rótulo para CustomTab ou Group.</span><span class="sxs-lookup"><span data-stu-id="535f2-117">The label for the CustomTab or a Group.</span></span>  |
|  [<span data-ttu-id="535f2-118">Control</span><span class="sxs-lookup"><span data-stu-id="535f2-118">Control</span></span>](control.md)    | <span data-ttu-id="535f2-119">Sim</span><span class="sxs-lookup"><span data-stu-id="535f2-119">Yes</span></span> |  <span data-ttu-id="535f2-120">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="535f2-120">A collection of one or more Control objects.</span></span>  |

### <a name="group"></a><span data-ttu-id="535f2-121">Group</span><span class="sxs-lookup"><span data-stu-id="535f2-121">Group</span></span>

<span data-ttu-id="535f2-p103">Obrigatório. Confira [Elemento Group](group.md).</span><span class="sxs-lookup"><span data-stu-id="535f2-p103">Required. See [Group element](group.md).</span></span>

### <a name="label-tab"></a><span data-ttu-id="535f2-124">Label (Tab)</span><span class="sxs-lookup"><span data-stu-id="535f2-124">Label (Tab)</span></span>

<span data-ttu-id="535f2-p104">Obrigatório. O rótulo da guia personalizada. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="535f2-p104">Required. The label of the custom tab. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>


## <a name="customtab-example"></a><span data-ttu-id="535f2-127">Exemplo CustomTab</span><span class="sxs-lookup"><span data-stu-id="535f2-127">CustomTab example</span></span>

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```