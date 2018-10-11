# <a name="officemenu-element"></a><span data-ttu-id="6acfc-101">Elemento OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="6acfc-101">OfficeMenu element</span></span>

<span data-ttu-id="6acfc-p101">Define um conjunto de controles que serão adicionados ao menu de contexto do Office. Aplica-se aos suplementos do Word, do Excel, do PowerPoint e do OneNote.</span><span class="sxs-lookup"><span data-stu-id="6acfc-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="6acfc-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="6acfc-104">Attributes</span></span>

| <span data-ttu-id="6acfc-105">Atributo</span><span class="sxs-lookup"><span data-stu-id="6acfc-105">Attribute</span></span>            | <span data-ttu-id="6acfc-106">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="6acfc-106">Required</span></span> | <span data-ttu-id="6acfc-107">Descrição</span><span class="sxs-lookup"><span data-stu-id="6acfc-107">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="6acfc-108">xsi:type</span><span class="sxs-lookup"><span data-stu-id="6acfc-108">xsi:type</span></span>](#xsitype) | <span data-ttu-id="6acfc-109">Sim</span><span class="sxs-lookup"><span data-stu-id="6acfc-109">Yes</span></span>      | <span data-ttu-id="6acfc-110">O tipo de OfficeMenu que está sendo definido.</span><span class="sxs-lookup"><span data-stu-id="6acfc-110">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="6acfc-111">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="6acfc-111">Child elements</span></span>

|  <span data-ttu-id="6acfc-112">Elemento</span><span class="sxs-lookup"><span data-stu-id="6acfc-112">Element</span></span> |  <span data-ttu-id="6acfc-113">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="6acfc-113">Required</span></span>  |  <span data-ttu-id="6acfc-114">Descrição</span><span class="sxs-lookup"><span data-stu-id="6acfc-114">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6acfc-115">Control</span><span class="sxs-lookup"><span data-stu-id="6acfc-115">Control</span></span>](#control)    | <span data-ttu-id="6acfc-116">Sim</span><span class="sxs-lookup"><span data-stu-id="6acfc-116">Yes</span></span> |  <span data-ttu-id="6acfc-117">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="6acfc-117">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="6acfc-118">xsi:type</span><span class="sxs-lookup"><span data-stu-id="6acfc-118">xsi:type</span></span>

<span data-ttu-id="6acfc-119">Especifica um menu interno do aplicativo cliente do Office ao qual você deseja adicionar esse suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="6acfc-119">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="6acfc-p102">`ContextMenuText` - Exibe o item no menu de contexto quando o texto for selecionado e o usuário abrir o menu de contexto (clicando com o botão direito do mouse) no texto selecionado. Aplica-se a Word, Excel, PowerPoint e OneNote.</span><span class="sxs-lookup"><span data-stu-id="6acfc-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="6acfc-p103">`ContextMenuCell` - Exibe o item no menu de contexto quando o usuário abrir o menu de contexto (clicando com o botão direito do mouse) em uma célula na planilha. Aplica-se ao Excel.</span><span class="sxs-lookup"><span data-stu-id="6acfc-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="6acfc-124">Control</span><span class="sxs-lookup"><span data-stu-id="6acfc-124">Control</span></span>

<span data-ttu-id="6acfc-125">Cada elemento **OfficeMenu** requer um ou mais controles de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="6acfc-125">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="6acfc-126">Exemplo</span><span class="sxs-lookup"><span data-stu-id="6acfc-126">Example</span></span>

```xml
<OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Menu" id="myMenuID">
      <Label resid="residLabel3" />
      <Supertip>
          <Title resid="residLabel" />
          <Description resid="residToolTip" />
      </Supertip>   
      <Icon>
        <bt:Image size="16" resid="icon1_16x16" />
        <bt:Image size="32" resid="icon1_32x32" />
        <bt:Image size="80" resid="icon1_80x80" />
      </Icon>    
      <Items>
        <Item id="myMenuItemID">
          <Label resid="residLabel3"/>
          <Supertip>
            <Title resid="residLabel" />
            <Description resid="residToolTip" />
          </Supertip>
          <Icon>
            <bt:Image size="16" resid="icon1_16x16" />
            <bt:Image size="32" resid="icon1_32x32" />
            <bt:Image size="80" resid="icon1_80x80" />
          </Icon>    
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl2" />    
          </Action>    
        </Item>
      </Items>
    </Control>   
</OfficeMenu>
```
