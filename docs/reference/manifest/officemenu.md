---
title: Elemento OfficeMenu no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 20d020b8ab826049ef0271cbdb8d51201ee88ab4
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452015"
---
# <a name="officemenu-element"></a><span data-ttu-id="e394d-102">Elemento OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="e394d-102">OfficeMenu element</span></span>

<span data-ttu-id="e394d-p101">Define um conjunto de controles que serão adicionados ao menu de contexto do Office. Aplica-se aos suplementos do Word, do Excel, do PowerPoint e do OneNote.</span><span class="sxs-lookup"><span data-stu-id="e394d-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="e394d-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="e394d-105">Attributes</span></span>

| <span data-ttu-id="e394d-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="e394d-106">Attribute</span></span>            | <span data-ttu-id="e394d-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="e394d-107">Required</span></span> | <span data-ttu-id="e394d-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="e394d-108">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="e394d-109">xsi:type</span><span class="sxs-lookup"><span data-stu-id="e394d-109">xsi:type</span></span>](#xsitype) | <span data-ttu-id="e394d-110">Sim</span><span class="sxs-lookup"><span data-stu-id="e394d-110">Yes</span></span>      | <span data-ttu-id="e394d-111">O tipo de OfficeMenu está sendo definido.</span><span class="sxs-lookup"><span data-stu-id="e394d-111">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="e394d-112">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="e394d-112">Child elements</span></span>

|  <span data-ttu-id="e394d-113">Elemento</span><span class="sxs-lookup"><span data-stu-id="e394d-113">Element</span></span> |  <span data-ttu-id="e394d-114">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="e394d-114">Required</span></span>  |  <span data-ttu-id="e394d-115">Descrição</span><span class="sxs-lookup"><span data-stu-id="e394d-115">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e394d-116">Control</span><span class="sxs-lookup"><span data-stu-id="e394d-116">Control</span></span>](#control)    | <span data-ttu-id="e394d-117">Sim</span><span class="sxs-lookup"><span data-stu-id="e394d-117">Yes</span></span> |  <span data-ttu-id="e394d-118">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="e394d-118">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="e394d-119">xsi:type</span><span class="sxs-lookup"><span data-stu-id="e394d-119">xsi:type</span></span>

<span data-ttu-id="e394d-120">Especifica um menu interno do aplicativo cliente do Office no qual você deseja adicionar esse suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="e394d-120">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="e394d-p102">`ContextMenuText` -  Exibe o item no menu de contexto quando o texto for selecionado e o usuário abre o menu de contexto (clica com o botão direito do mouse) no texto selecionado. Aplica-se a Word, Excel, PowerPoint e OneNote.</span><span class="sxs-lookup"><span data-stu-id="e394d-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="e394d-p103">`ContextMenuCell` -  Exibe o item no menu de contexto quando o usuário abre o menu de contexto (clica com o botão direito do mouse) em uma célula na planilha. Aplica-se ao Excel.</span><span class="sxs-lookup"><span data-stu-id="e394d-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="e394d-125">Control</span><span class="sxs-lookup"><span data-stu-id="e394d-125">Control</span></span>

<span data-ttu-id="e394d-126">Cada elemento **OfficeMenu** requer um ou mais controles de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="e394d-126">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="e394d-127">Exemplo</span><span class="sxs-lookup"><span data-stu-id="e394d-127">Example</span></span>

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
