---
title: Componente ChoiceGroup no Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 78da2fae781039663bfe2bac159bfbe50192c023
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="choicegroup-component-in-office-ui-fabric"></a><span data-ttu-id="87812-102">Componente ChoiceGroup no Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="87812-102">ChoiceGroup component in Office UI Fabric</span></span>

<span data-ttu-id="87812-p101">O componente ChoiceGroup, tamb?m conhecido como um bot?o de op??o, apresenta aos usu?rios duas ou mais op??es mutuamente exclusivas. Os usu?rios podem selecionar apenas um bot?o do ChoiceGroup em um grupo. Cada op??o ? representada por um bot?o do ChoiceGroup.</span><span class="sxs-lookup"><span data-stu-id="87812-p101">The ChoiceGroup component, also known as a radio button, presents users with two or more mutually exclusive options. Users can select only one ChoiceGroup button in a group. Each option is represented by one ChoiceGroup button.</span></span> 
  
#### <a name="example-choicegroup-in-a-task-pane"></a><span data-ttu-id="87812-106">Exemplo: ChoiceGroup em um painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="87812-106">Example: ChoiceGroup in a task pane</span></span>

 ![Imagem mostrando um ChoiceGroup](../images/overview-with-app-choicegroup.png)

## <a name="best-practices"></a><span data-ttu-id="87812-108">Pr?ticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="87812-108">Best practices</span></span>

|<span data-ttu-id="87812-109">**Fa?a**</span><span class="sxs-lookup"><span data-stu-id="87812-109">**Do**</span></span>|<span data-ttu-id="87812-110">**N?o fa?a**</span><span class="sxs-lookup"><span data-stu-id="87812-110">**Don't**</span></span>|
|:------------|:--------------|
|<span data-ttu-id="87812-111">Mantenha as op??es de ChoiceGroup no mesmo n?vel.</span><span class="sxs-lookup"><span data-stu-id="87812-111">Keep ChoiceGroup options at the same level.</span></span><br/><br/>![Exemplo do que fazer com ChoiceGroup](../images/choice-do.png)<br/>|<span data-ttu-id="87812-113">N?o utilize ChoiceGroups ou caixas de sele??o aninhados.</span><span class="sxs-lookup"><span data-stu-id="87812-113">Don't use nested ChoiceGroups or check boxes.</span></span><br/><br/>![Exemplo do que n?o fazer com ChoiceGroup](../images/choice-dont.png)<br/>|
|<span data-ttu-id="87812-115">Use ChoiceGroups com duas a sete op??es, verificando se h? espa?o suficiente na tela para mostrar todas as op??es.</span><span class="sxs-lookup"><span data-stu-id="87812-115">Use ChoiceGroups with 2-7 options, ensuring there is enough screen space to show all options.</span></span> <span data-ttu-id="87812-116">Caso contr?rio, use uma caixa de sele??o ou lista suspensa.</span><span class="sxs-lookup"><span data-stu-id="87812-116">Otherwise, use a check box or drop-down list.</span></span>|<span data-ttu-id="87812-p103">N?o use quando as op??es forem n?meros com uma grada??o fixa, por exemplo, 10, 20, 30 e assim por diante. Em vez disso, use um componente de controle deslizante.</span><span class="sxs-lookup"><span data-stu-id="87812-p103">Don't use when the options are numbers with a fixed step, for example 10, 20, 30, and so on. Instead, use a slider component.</span></span>|
|<span data-ttu-id="87812-119">Se os usu?rios n?o puderem escolher nenhuma das op??es, considere incluir uma op??o como **Nenhum** ou **N?o se aplica**.</span><span class="sxs-lookup"><span data-stu-id="87812-119">If users may not choose any of the options, consider including an option such as **None** or **Does not apply**.</span></span>|<span data-ttu-id="87812-120">N?o use dois bot?es de ChoiceGroup para uma ?nica op??o bin?ria.</span><span class="sxs-lookup"><span data-stu-id="87812-120">Don?t use two ChoiceGroup buttons for a single binary choice.</span></span>|
|<span data-ttu-id="87812-p104">Se poss?vel, alinhe os bot?es de ChoiceGroup verticalmente em vez de horizontalmente. O alinhamento horizontal ? mais dif?cil de ler e localizar.</span><span class="sxs-lookup"><span data-stu-id="87812-p104">If possible, align ChoiceGroup buttons vertically instead of horizontally. Horizontal alignment is harder to read and localize.</span></span>||
|<span data-ttu-id="87812-123">Liste as op??es em ordem l?gica, por exemplo, da op??o mais prov?vel a ser selecionada at? a menos, da opera??o mais simples at? a mais complexa ou do menor risco para o maior risco.</span><span class="sxs-lookup"><span data-stu-id="87812-123">List options in logical order, for example, the most likely option to be selected to the least, the simplest operation to the most complex, or the least risk to the highest risk.</span></span> |<span data-ttu-id="87812-124">N?o use ordena??o alfab?tica porque ? dependente do idioma.</span><span class="sxs-lookup"><span data-stu-id="87812-124">Don't use alphabetical ordering because it is language dependent.</span></span>|

## <a name="variants"></a><span data-ttu-id="87812-125">Variantes</span><span class="sxs-lookup"><span data-stu-id="87812-125">Variants</span></span>

|<span data-ttu-id="87812-126">**Varia??o**</span><span class="sxs-lookup"><span data-stu-id="87812-126">**Variation**</span></span>|<span data-ttu-id="87812-127">**Descri??o**</span><span class="sxs-lookup"><span data-stu-id="87812-127">**Description**</span></span>|<span data-ttu-id="87812-128">**Exemplo**</span><span class="sxs-lookup"><span data-stu-id="87812-128">**Example**</span></span>|
|:------------|:--------------|:----------|
|<span data-ttu-id="87812-129">**ChoiceGroups**</span><span class="sxs-lookup"><span data-stu-id="87812-129">**ChoiceGroups**</span></span>|<span data-ttu-id="87812-130">Use quando n?o forem necess?rias imagens para fazer uma escolha.</span><span class="sxs-lookup"><span data-stu-id="87812-130">Use when imagery is not necessary for making a selection.</span></span>|![Imagem da variante de ChoiceGroup](../images/radio.png)<br/>|
|<span data-ttu-id="87812-132">**ChoiceGroups usando imagens**</span><span class="sxs-lookup"><span data-stu-id="87812-132">**ChoiceGroups using images**</span></span>|<span data-ttu-id="87812-133">Use quando forem necess?rias imagens para fazer uma escolha.</span><span class="sxs-lookup"><span data-stu-id="87812-133">Use when imagery is necessary for making a selection.</span></span>|![Variante de ChoiceGroup com imagem](../images/radio-image.png)<br/>|

## <a name="implementation"></a><span data-ttu-id="87812-135">Implementa??o</span><span class="sxs-lookup"><span data-stu-id="87812-135">Implementation</span></span>

<span data-ttu-id="87812-136">Para saber mais, confira [ChoiceGroup](https://dev.office.com/fabric#/components/choicegroup) e [Primeiros passos com exemplo de c?digo do Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span><span class="sxs-lookup"><span data-stu-id="87812-136">For details, see [ChoiceGroup](https://dev.office.com/fabric#/components/choicegroup) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span></span>

## <a name="see-also"></a><span data-ttu-id="87812-137">Veja tamb?m</span><span class="sxs-lookup"><span data-stu-id="87812-137">See also</span></span>

- [<span data-ttu-id="87812-138">Padr?es de design da experi?ncia do usu?rio</span><span class="sxs-lookup"><span data-stu-id="87812-138">UX Design Patterns</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="87812-139">Office UI Fabric em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="87812-139">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
