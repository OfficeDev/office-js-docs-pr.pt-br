---
title: Elemento OfficeMenu no arquivo de manifesto
description: O elemento OfficeMenu define uma coleção de controles a serem adicionados ao menu de contexto do Office.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 89503533f7310898a420eb805d5fd66f096ad5f2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718045"
---
# <a name="officemenu-element"></a><span data-ttu-id="3c14e-103">Elemento OfficeMenu</span><span class="sxs-lookup"><span data-stu-id="3c14e-103">OfficeMenu element</span></span>

<span data-ttu-id="3c14e-p101">Define um conjunto de controles que serão adicionados ao menu de contexto do Office. Aplica-se aos suplementos do Word, do Excel, do PowerPoint e do OneNote.</span><span class="sxs-lookup"><span data-stu-id="3c14e-p101">Defines a collection of controls to be added to the Office context menu. Applies to Word, Excel, PowerPoint, and OneNote add-ins.</span></span>

## <a name="attributes"></a><span data-ttu-id="3c14e-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="3c14e-106">Attributes</span></span>

| <span data-ttu-id="3c14e-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="3c14e-107">Attribute</span></span>            | <span data-ttu-id="3c14e-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3c14e-108">Required</span></span> | <span data-ttu-id="3c14e-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="3c14e-109">Description</span></span>                          |
|:---------------------|:--------:|:-------------------------------------|
| [<span data-ttu-id="3c14e-110">xsi:type</span><span class="sxs-lookup"><span data-stu-id="3c14e-110">xsi:type</span></span>](#xsitype) | <span data-ttu-id="3c14e-111">Sim</span><span class="sxs-lookup"><span data-stu-id="3c14e-111">Yes</span></span>      | <span data-ttu-id="3c14e-112">O tipo de OfficeMenu está sendo definido.</span><span class="sxs-lookup"><span data-stu-id="3c14e-112">The type of OfficeMenu being defined.</span></span>|

## <a name="child-elements"></a><span data-ttu-id="3c14e-113">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="3c14e-113">Child elements</span></span>

|  <span data-ttu-id="3c14e-114">Elemento</span><span class="sxs-lookup"><span data-stu-id="3c14e-114">Element</span></span> |  <span data-ttu-id="3c14e-115">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="3c14e-115">Required</span></span>  |  <span data-ttu-id="3c14e-116">Descrição</span><span class="sxs-lookup"><span data-stu-id="3c14e-116">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3c14e-117">Control</span><span class="sxs-lookup"><span data-stu-id="3c14e-117">Control</span></span>](#control)    | <span data-ttu-id="3c14e-118">Sim</span><span class="sxs-lookup"><span data-stu-id="3c14e-118">Yes</span></span> |  <span data-ttu-id="3c14e-119">Conjunto de um ou mais objetos Control.</span><span class="sxs-lookup"><span data-stu-id="3c14e-119">A collection of one or more Control objects.</span></span>  |

## <a name="xsitype"></a><span data-ttu-id="3c14e-120">xsi:type</span><span class="sxs-lookup"><span data-stu-id="3c14e-120">xsi:type</span></span>

<span data-ttu-id="3c14e-121">Especifica um menu interno do aplicativo cliente do Office no qual você deseja adicionar esse suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="3c14e-121">Specifies a built-in menu of the Office client application on which to add this Office Add-in.</span></span>

- <span data-ttu-id="3c14e-p102">`ContextMenuText` -  Exibe o item no menu de contexto quando o texto for selecionado e o usuário abre o menu de contexto (clica com o botão direito do mouse) no texto selecionado. Aplica-se a Word, Excel, PowerPoint e OneNote.</span><span class="sxs-lookup"><span data-stu-id="3c14e-p102">`ContextMenuText` -  Displays the item on the context menu when text is selected and the user opens the context menu (right-clicks) on the selected text. Applies to Word, Excel, PowerPoint, and OneNote.</span></span>
- <span data-ttu-id="3c14e-p103">`ContextMenuCell` -  Exibe o item no menu de contexto quando o usuário abre o menu de contexto (clica com o botão direito do mouse) em uma célula na planilha. Aplica-se ao Excel.</span><span class="sxs-lookup"><span data-stu-id="3c14e-p103">`ContextMenuCell` -  Displays the item on the context menu when the user opens the context menu (right-clicks) on a cell on the spreadsheet. Applies to Excel.</span></span> 

## <a name="control"></a><span data-ttu-id="3c14e-126">Control</span><span class="sxs-lookup"><span data-stu-id="3c14e-126">Control</span></span>

<span data-ttu-id="3c14e-127">Cada elemento **OfficeMenu** requer um ou mais controles de [menu](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="3c14e-127">Each **OfficeMenu** element requires at one or more [menu](control.md#menu-dropdown-button-controls) controls.</span></span> 

## <a name="example"></a><span data-ttu-id="3c14e-128">Exemplo</span><span class="sxs-lookup"><span data-stu-id="3c14e-128">Example</span></span>

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
