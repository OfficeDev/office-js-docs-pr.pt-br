---
title: Componente ChoiceGroup no Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 78da2fae781039663bfe2bac159bfbe50192c023
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437210"
---
# <a name="choicegroup-component-in-office-ui-fabric"></a><span data-ttu-id="17124-102">Componente ChoiceGroup no Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="17124-102">ChoiceGroup component in Office UI Fabric</span></span>

<span data-ttu-id="17124-p101">O componente ChoiceGroup, também conhecido como um botão de opção, apresenta aos usuários duas ou mais opções mutuamente exclusivas. Os usuários podem selecionar apenas um botão do ChoiceGroup em um grupo. Cada opção é representada por um botão do ChoiceGroup.</span><span class="sxs-lookup"><span data-stu-id="17124-p101">The ChoiceGroup component, also known as a radio button, presents users with two or more mutually exclusive options. Users can select only one ChoiceGroup button in a group. Each option is represented by one ChoiceGroup button.</span></span> 
  
#### <a name="example-choicegroup-in-a-task-pane"></a><span data-ttu-id="17124-106">Exemplo: ChoiceGroup em um painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="17124-106">Example: ChoiceGroup in a task pane</span></span>

 ![Imagem mostrando um ChoiceGroup](../images/overview-with-app-choicegroup.png)

## <a name="best-practices"></a><span data-ttu-id="17124-108">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="17124-108">Best practices</span></span>

|<span data-ttu-id="17124-109">**Faça**</span><span class="sxs-lookup"><span data-stu-id="17124-109">**Do**</span></span>|<span data-ttu-id="17124-110">**Não faça**</span><span class="sxs-lookup"><span data-stu-id="17124-110">**Don't**</span></span>|
|:------------|:--------------|
|<span data-ttu-id="17124-111">Mantenha as opções de ChoiceGroup no mesmo nível.</span><span class="sxs-lookup"><span data-stu-id="17124-111">Keep ChoiceGroup options at the same level.</span></span><br/><br/>![Exemplo do que fazer com ChoiceGroup](../images/choice-do.png)<br/>|<span data-ttu-id="17124-113">Não utilize ChoiceGroups ou caixas de seleção aninhados.</span><span class="sxs-lookup"><span data-stu-id="17124-113">Don't use nested ChoiceGroups or check boxes.</span></span><br/><br/>![Exemplo do que não fazer com ChoiceGroup](../images/choice-dont.png)<br/>|
|<span data-ttu-id="17124-115">Use ChoiceGroups com duas a sete opções, verificando se há espaço suficiente na tela para mostrar todas as opções.</span><span class="sxs-lookup"><span data-stu-id="17124-115">Use ChoiceGroups with 2-7 options, ensuring there is enough screen space to show all options.</span></span> <span data-ttu-id="17124-116">Caso contrário, use uma caixa de seleção ou lista suspensa.</span><span class="sxs-lookup"><span data-stu-id="17124-116">Otherwise, use a check box or drop-down list.</span></span>|<span data-ttu-id="17124-p103">Não use quando as opções forem números com uma gradação fixa, por exemplo, 10, 20, 30 e assim por diante. Em vez disso, use um componente de controle deslizante.</span><span class="sxs-lookup"><span data-stu-id="17124-p103">Don't use when the options are numbers with a fixed step, for example 10, 20, 30, and so on. Instead, use a slider component.</span></span>|
|<span data-ttu-id="17124-119">Se os usuários não puderem escolher nenhuma das opções, considere incluir uma opção como **Nenhum** ou **Não se aplica**.</span><span class="sxs-lookup"><span data-stu-id="17124-119">If users may not choose any of the options, consider including an option such as **None** or **Does not apply**.</span></span>|<span data-ttu-id="17124-120">Não use dois botões de ChoiceGroup para uma única opção binária.</span><span class="sxs-lookup"><span data-stu-id="17124-120">Don’t use two ChoiceGroup buttons for a single binary choice.</span></span>|
|<span data-ttu-id="17124-p104">Se possível, alinhe os botões de ChoiceGroup verticalmente em vez de horizontalmente. O alinhamento horizontal é mais difícil de ler e localizar.</span><span class="sxs-lookup"><span data-stu-id="17124-p104">If possible, align ChoiceGroup buttons vertically instead of horizontally. Horizontal alignment is harder to read and localize.</span></span>||
|<span data-ttu-id="17124-123">Liste as opções em ordem lógica, por exemplo, da opção mais provável a ser selecionada até a menos, da operação mais simples até a mais complexa ou do menor risco para o maior risco.</span><span class="sxs-lookup"><span data-stu-id="17124-123">List options in logical order, for example, the most likely option to be selected to the least, the simplest operation to the most complex, or the least risk to the highest risk.</span></span> |<span data-ttu-id="17124-124">Não use ordenação alfabética porque é dependente do idioma.</span><span class="sxs-lookup"><span data-stu-id="17124-124">Don't use alphabetical ordering because it is language dependent.</span></span>|

## <a name="variants"></a><span data-ttu-id="17124-125">Variantes</span><span class="sxs-lookup"><span data-stu-id="17124-125">Variants</span></span>

|<span data-ttu-id="17124-126">**Variação**</span><span class="sxs-lookup"><span data-stu-id="17124-126">**Variation**</span></span>|<span data-ttu-id="17124-127">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="17124-127">**Description**</span></span>|<span data-ttu-id="17124-128">**Exemplo**</span><span class="sxs-lookup"><span data-stu-id="17124-128">**Example**</span></span>|
|:------------|:--------------|:----------|
|<span data-ttu-id="17124-129">**ChoiceGroups**</span><span class="sxs-lookup"><span data-stu-id="17124-129">**ChoiceGroups**</span></span>|<span data-ttu-id="17124-130">Use quando não forem necessárias imagens para fazer uma escolha.</span><span class="sxs-lookup"><span data-stu-id="17124-130">Use when imagery is not necessary for making a selection.</span></span>|![Imagem da variante de ChoiceGroup](../images/radio.png)<br/>|
|<span data-ttu-id="17124-132">**ChoiceGroups usando imagens**</span><span class="sxs-lookup"><span data-stu-id="17124-132">**ChoiceGroups using images**</span></span>|<span data-ttu-id="17124-133">Use quando forem necessárias imagens para fazer uma escolha.</span><span class="sxs-lookup"><span data-stu-id="17124-133">Use when imagery is necessary for making a selection.</span></span>|![Variante de ChoiceGroup com imagem](../images/radio-image.png)<br/>|

## <a name="implementation"></a><span data-ttu-id="17124-135">Implementação</span><span class="sxs-lookup"><span data-stu-id="17124-135">Implementation</span></span>

<span data-ttu-id="17124-136">Para saber mais, confira [ChoiceGroup](https://dev.office.com/fabric#/components/choicegroup) e [Primeiros passos com exemplo de código do Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span><span class="sxs-lookup"><span data-stu-id="17124-136">For details, see [ChoiceGroup](https://dev.office.com/fabric#/components/choicegroup) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span></span>

## <a name="see-also"></a><span data-ttu-id="17124-137">Veja também</span><span class="sxs-lookup"><span data-stu-id="17124-137">See also</span></span>

- [<span data-ttu-id="17124-138">Padrões de design da experiência do usuário</span><span class="sxs-lookup"><span data-stu-id="17124-138">UX Design Patterns</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="17124-139">Office UI Fabric em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="17124-139">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
