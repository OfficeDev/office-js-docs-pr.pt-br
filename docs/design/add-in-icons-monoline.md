---
title: Diretrizes de ícone de estilo monoline para suplementos do Office
description: Obter diretrizes para usar ícones de ícone de estilo monoline em suplementos do Office.
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 264aa9e01bd70924cfee01a864c515c8c7a4d138
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132197"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a><span data-ttu-id="46cb7-103">Diretrizes de ícone de estilo monoline para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="46cb7-103">Monoline style icon guidelines for Office Add-ins</span></span>

<span data-ttu-id="46cb7-104">O estilo monoline iconografia são usados no Office 365.</span><span class="sxs-lookup"><span data-stu-id="46cb7-104">Monoline style iconography are used in Office 365.</span></span> <span data-ttu-id="46cb7-105">Se você preferir que seus ícones correspondam ao novo estilo de não assinatura do Office 2013 +, confira [diretrizes de ícone de estilo atualizado para suplementos do Office](add-in-icons-fresh.md).</span><span class="sxs-lookup"><span data-stu-id="46cb7-105">If you would prefer that your icons match the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).</span></span>

## <a name="office-monoline-visual-style"></a><span data-ttu-id="46cb7-106">Estilo visual monoline do Office</span><span class="sxs-lookup"><span data-stu-id="46cb7-106">Office Monoline visual style</span></span>

<span data-ttu-id="46cb7-107">O objetivo do estilo de monolinha ter um iconografia consistente, claro e acessível para comunicar ações e recursos com visuais simples, garantir que os ícones estejam acessíveis a todos os usuários e ter um estilo consistente com aqueles usados em qualquer lugar no Windows.</span><span class="sxs-lookup"><span data-stu-id="46cb7-107">The goal of the Monoline style to have consistent, clear, and accessible iconography to communicate action and features with simple visuals, ensure the icons are accessible to all users, and have a style that is consistent with those used elsewhere in Windows.</span></span>

<span data-ttu-id="46cb7-108">As diretrizes a seguir são para desenvolvedores de terceiros que desejam criar ícones para recursos que serão consistentes com os ícones já presentes nos produtos do Office.</span><span class="sxs-lookup"><span data-stu-id="46cb7-108">The following guidelines are for 3rd party developers who want to create icons for features that will be consistent with the icons already present Office products.</span></span>

### <a name="design-principles"></a><span data-ttu-id="46cb7-109">Princípios de design</span><span class="sxs-lookup"><span data-stu-id="46cb7-109">Design principles</span></span>

- <span data-ttu-id="46cb7-110">Simples, limpo, claro.</span><span class="sxs-lookup"><span data-stu-id="46cb7-110">Simple, clean, clear.</span></span>
- <span data-ttu-id="46cb7-111">Conter apenas elementos necessários.</span><span class="sxs-lookup"><span data-stu-id="46cb7-111">Contain only necessary elements.</span></span>
- <span data-ttu-id="46cb7-112">Estilo de ícone do Windows inspirado.</span><span class="sxs-lookup"><span data-stu-id="46cb7-112">Inspired by Windows icon style.</span></span>
- <span data-ttu-id="46cb7-113">Acessível a todos os usuários.</span><span class="sxs-lookup"><span data-stu-id="46cb7-113">Accessible to all users.</span></span>

#### <a name="conveying-meaning"></a><span data-ttu-id="46cb7-114">Transmitir significado</span><span class="sxs-lookup"><span data-stu-id="46cb7-114">Conveying meaning</span></span>

- <span data-ttu-id="46cb7-115">Use elementos descritivos, como uma página para representar um documento ou envelope para representar emails.</span><span class="sxs-lookup"><span data-stu-id="46cb7-115">Use descriptive elements such as a page to represent a document or an envelope to represent mail.</span></span>
- <span data-ttu-id="46cb7-116">Use o mesmo elemento para representar o mesmo conceito, ou seja, mail é sempre representado por um envelope, não um carimbo.</span><span class="sxs-lookup"><span data-stu-id="46cb7-116">Use the same element to represent the same concept, i.e., mail is always represented by an envelope, not a stamp.</span></span>
- <span data-ttu-id="46cb7-117">Use uma metáfora principal durante o desenvolvimento do conceito.</span><span class="sxs-lookup"><span data-stu-id="46cb7-117">Use a core metaphor during concept development.</span></span>

#### <a name="reduction-of-elements"></a><span data-ttu-id="46cb7-118">Redução dos elementos</span><span class="sxs-lookup"><span data-stu-id="46cb7-118">Reduction of Elements</span></span>

- <span data-ttu-id="46cb7-119">Reduza o ícone ao seu significado principal, usando apenas os elementos essenciais para a metáfora.</span><span class="sxs-lookup"><span data-stu-id="46cb7-119">Reduce the icon to its core meaning, using only elements that are essential to the metaphor.</span></span>
- <span data-ttu-id="46cb7-120">Limitar o número de elementos em um ícone a dois, independentemente do tamanho do ícone.</span><span class="sxs-lookup"><span data-stu-id="46cb7-120">Limit the number of elements in an icon to two, regardless of icon size.</span></span>

#### <a name="consistency"></a><span data-ttu-id="46cb7-121">Consistência</span><span class="sxs-lookup"><span data-stu-id="46cb7-121">Consistency</span></span>

<span data-ttu-id="46cb7-122">Os tamanhos, a organização e a cor dos ícones devem ser consistentes.</span><span class="sxs-lookup"><span data-stu-id="46cb7-122">Sizes, arrangement, and color of icons should be consistent.</span></span>

#### <a name="styling"></a><span data-ttu-id="46cb7-123">Estilo</span><span class="sxs-lookup"><span data-stu-id="46cb7-123">Styling</span></span>

##### <a name="perspective"></a><span data-ttu-id="46cb7-124">Perspectiva</span><span class="sxs-lookup"><span data-stu-id="46cb7-124">Perspective</span></span>

<span data-ttu-id="46cb7-125">Os ícones monoline estão voltados para o avanço por padrão.</span><span class="sxs-lookup"><span data-stu-id="46cb7-125">Monoline icons are forward-facing by default.</span></span> <span data-ttu-id="46cb7-126">Determinados elementos que exigem perspectiva e/ou rotação, como um cubo, são permitidos, mas as exceções devem ser mantidas no mínimo.</span><span class="sxs-lookup"><span data-stu-id="46cb7-126">Certain elements that require perspective and/or rotation, such as a cube, are allowed, but exceptions should be kept to a minimum.</span></span>

##### <a name="embellishment"></a><span data-ttu-id="46cb7-127">Ornamento</span><span class="sxs-lookup"><span data-stu-id="46cb7-127">Embellishment</span></span>

<span data-ttu-id="46cb7-128">Monolinha é um estilo mínimo limpo.</span><span class="sxs-lookup"><span data-stu-id="46cb7-128">Monoline is a clean minimal style.</span></span> <span data-ttu-id="46cb7-129">Tudo usa cor plana, o que significa que não há gradientes, texturas ou fontes de luz.</span><span class="sxs-lookup"><span data-stu-id="46cb7-129">Everything uses flat color, which means there are no gradients, textures, or light sources.</span></span>

## <a name="designing"></a><span data-ttu-id="46cb7-130">Planejamento</span><span class="sxs-lookup"><span data-stu-id="46cb7-130">Designing</span></span>

### <a name="sizes"></a><span data-ttu-id="46cb7-131">Coincidi</span><span class="sxs-lookup"><span data-stu-id="46cb7-131">Sizes</span></span>

<span data-ttu-id="46cb7-132">Recomendamos que você produza cada ícone em todos esses tamanhos para suportar dispositivos DPI alto.</span><span class="sxs-lookup"><span data-stu-id="46cb7-132">We recommend that you produce each icon in all these sizes to support high DPI devices.</span></span> <span data-ttu-id="46cb7-133">Os tamanhos absolutamente *necessários* são 16 px, 20 px e 32 px, pois esses são os tamanhos de 100%.</span><span class="sxs-lookup"><span data-stu-id="46cb7-133">The absolutely *required* sizes are 16 px, 20 px, and 32 px, as those are the 100% sizes.</span></span>

<span data-ttu-id="46cb7-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span><span class="sxs-lookup"><span data-stu-id="46cb7-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span></span>

### <a name="layout"></a><span data-ttu-id="46cb7-135">Layout</span><span class="sxs-lookup"><span data-stu-id="46cb7-135">Layout</span></span>

<span data-ttu-id="46cb7-136">Veja a seguir um exemplo de layout de ícone com um modificador.</span><span class="sxs-lookup"><span data-stu-id="46cb7-136">The following is an example of icon layout with a modifier.</span></span>

![Diagrama de ícone com modificador no canto inferior direito](../images/monolineicon1.png)  ![Diagrama do mesmo ícone com plano de fundo de grade e textos explicativos adicionados para base, modificador, enchimento e recorte](../images/monolineicon2.png)

#### <a name="elements"></a><span data-ttu-id="46cb7-139">Elementos</span><span class="sxs-lookup"><span data-stu-id="46cb7-139">Elements</span></span>

- <span data-ttu-id="46cb7-140">**Base**: o conceito principal que o ícone representa.</span><span class="sxs-lookup"><span data-stu-id="46cb7-140">**Base**: The main concept that the icon represents.</span></span> <span data-ttu-id="46cb7-141">Isso geralmente é o único Visual necessário para o ícone, mas às vezes o conceito principal pode ser aprimorado com um elemento secundário, um modificador.</span><span class="sxs-lookup"><span data-stu-id="46cb7-141">This is usually the only visual needed for the icon, but sometimes the main concept can be enhanced with a secondary element, a modifier.</span></span>

- <span data-ttu-id="46cb7-142">**Modificador** Qualquer elemento que sobrepõe a base; ou seja, um modificador que normalmente representa uma ação ou um status.</span><span class="sxs-lookup"><span data-stu-id="46cb7-142">**Modifier** Any element that overlays the base; that is, a modifier that typically represents an action or a status.</span></span> <span data-ttu-id="46cb7-143">Ele modifica o elemento base agindo como uma adição, alteração ou descritor.</span><span class="sxs-lookup"><span data-stu-id="46cb7-143">It modifies the base element by acting as an addition, alteration, or a descriptor.</span></span>

![Diagrama de grade com áreas de base e modificadores chamadas out](../images/monolineicon3.png)

### <a name="construction"></a><span data-ttu-id="46cb7-145">Construção</span><span class="sxs-lookup"><span data-stu-id="46cb7-145">Construction</span></span>

#### <a name="element-placement"></a><span data-ttu-id="46cb7-146">Posicionamento do elemento</span><span class="sxs-lookup"><span data-stu-id="46cb7-146">Element placement</span></span>

<span data-ttu-id="46cb7-147">Os elementos base são colocados no centro do ícone dentro do preenchimento.</span><span class="sxs-lookup"><span data-stu-id="46cb7-147">Base elements are placed in the center of the icon within the padding.</span></span> <span data-ttu-id="46cb7-148">Se ele não puder ser colocado perfeitamente centralizado, a base deverá ter um erro no canto superior direito.</span><span class="sxs-lookup"><span data-stu-id="46cb7-148">If it can't be placed perfectly centered, then the base should err to the top right.</span></span> <span data-ttu-id="46cb7-149">No exemplo a seguir, o ícone está perfeitamente centralizado.</span><span class="sxs-lookup"><span data-stu-id="46cb7-149">In the following example, the icon is perfectly centered.</span></span>

![Diagrama mostrando o ícone perfeitamente centralizado](../images/monolineicon4.png)

<span data-ttu-id="46cb7-151">No exemplo a seguir, o ícone é erring à esquerda.</span><span class="sxs-lookup"><span data-stu-id="46cb7-151">In the following example, the icon is erring to the left.</span></span>

![Diagrama mostrando o ícone que ERRs à esquerda em 1 px](../images/monolineicon5.png)

<span data-ttu-id="46cb7-153">Modificadores quase sempre são colocados no canto inferior direito da tela de ícones.</span><span class="sxs-lookup"><span data-stu-id="46cb7-153">Modifiers are almost always placed in the bottom right corner of the icon canvas.</span></span> <span data-ttu-id="46cb7-154">Em alguns casos raros, os modificadores são colocados em um canto diferente.</span><span class="sxs-lookup"><span data-stu-id="46cb7-154">In some rare cases, modifiers are placed in a different corner.</span></span> <span data-ttu-id="46cb7-155">Por exemplo, se o elemento base não puder ser reconhecível com o modificador no canto inferior direito, considere colocá-lo no canto superior esquerdo.</span><span class="sxs-lookup"><span data-stu-id="46cb7-155">For example, if the base element would be unrecognizable with the modifier in the bottom right corner, then consider placing it in the upper left corner.</span></span>

![Diagrama mostrando quatro ícones com o modificador no canto inferior direito e um ícone com o modificador na parte superior esquerda](../images/monolineicon6.png)

#### <a name="padding"></a><span data-ttu-id="46cb7-157">Padding</span><span class="sxs-lookup"><span data-stu-id="46cb7-157">Padding</span></span>

<span data-ttu-id="46cb7-158">Cada ícone de tamanho tem uma quantidade especificada de preenchimento em torno do ícone.</span><span class="sxs-lookup"><span data-stu-id="46cb7-158">Each size icon has a specified amount of padding around the icon.</span></span> <span data-ttu-id="46cb7-159">O elemento base permanece dentro do preenchimento, mas o modificador deve arredondar para a borda da tela, estendendo-o de fora do preenchimento para a borda da borda do ícone.</span><span class="sxs-lookup"><span data-stu-id="46cb7-159">The base element stays within the padding, but the modifier should butt up to the edge of the canvas, extending outside of the padding to the edge of the icon border.</span></span> <span data-ttu-id="46cb7-160">As imagens a seguir mostram o preenchimento recomendado a ser usado para cada um dos tamanhos de ícone.</span><span class="sxs-lookup"><span data-stu-id="46cb7-160">The following images show the recommended padding to use for each of the icon sizes.</span></span>

|<span data-ttu-id="46cb7-161">**16px**</span><span class="sxs-lookup"><span data-stu-id="46cb7-161">**16px**</span></span>|<span data-ttu-id="46cb7-162">**20px**</span><span class="sxs-lookup"><span data-stu-id="46cb7-162">**20px**</span></span>|<span data-ttu-id="46cb7-163">**24px**</span><span class="sxs-lookup"><span data-stu-id="46cb7-163">**24px**</span></span>|<span data-ttu-id="46cb7-164">**32px**</span><span class="sxs-lookup"><span data-stu-id="46cb7-164">**32px**</span></span>|<span data-ttu-id="46cb7-165">**40px**</span><span class="sxs-lookup"><span data-stu-id="46cb7-165">**40px**</span></span>|<span data-ttu-id="46cb7-166">**48px**</span><span class="sxs-lookup"><span data-stu-id="46cb7-166">**48px**</span></span>|<span data-ttu-id="46cb7-167">**64px**</span><span class="sxs-lookup"><span data-stu-id="46cb7-167">**64px**</span></span>|<span data-ttu-id="46cb7-168">**80px**</span><span class="sxs-lookup"><span data-stu-id="46cb7-168">**80px**</span></span>|<span data-ttu-id="46cb7-169">**96px**</span><span class="sxs-lookup"><span data-stu-id="46cb7-169">**96px**</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![ícone 16 px com preenchimento 0px](../images/monolineicon7.png)|![ícone de 20 px com preenchimento 1 px](../images/monolineicon8.png)|![ícone de 24 PX com preenchimento 1 px](../images/monolineicon9.png)|![32 px Icon com preenchimento 2 px](../images/monolineicon10.png)|![40 PX Icon com preenchimento 2 px](../images/monolineicon11.png)|![48 PX Icon com preenchimento 3 px](../images/monolineicon12.png)|![64 PX Icon com preenchimento 4px](../images/monolineicon13.png)|![80 PX Icon com preenchimento 5px](../images/monolineicon14.png)|![96 PX Icon com preenchimento 6px](../images/monolineicon15.png)|

#### <a name="line-weights"></a><span data-ttu-id="46cb7-179">Espessuras de linha</span><span class="sxs-lookup"><span data-stu-id="46cb7-179">Line weights</span></span>

<span data-ttu-id="46cb7-180">Monolinha é um estilo dominado por formas de linha e contorno.</span><span class="sxs-lookup"><span data-stu-id="46cb7-180">Monoline is a style dominated by line and outlined shapes.</span></span> <span data-ttu-id="46cb7-181">Dependendo de qual tamanho você está produzindo, o ícone deve usar os pesos de linha a seguir.</span><span class="sxs-lookup"><span data-stu-id="46cb7-181">Depending on what size you are producing the icon should use the following line weights.</span></span>

|<span data-ttu-id="46cb7-182">Tamanho do ícone:</span><span class="sxs-lookup"><span data-stu-id="46cb7-182">Icon Size:</span></span>|<span data-ttu-id="46cb7-183">16px</span><span class="sxs-lookup"><span data-stu-id="46cb7-183">16px</span></span>|<span data-ttu-id="46cb7-184">20px</span><span class="sxs-lookup"><span data-stu-id="46cb7-184">20px</span></span>|<span data-ttu-id="46cb7-185">24px</span><span class="sxs-lookup"><span data-stu-id="46cb7-185">24px</span></span>|<span data-ttu-id="46cb7-186">32px</span><span class="sxs-lookup"><span data-stu-id="46cb7-186">32px</span></span>|<span data-ttu-id="46cb7-187">40px</span><span class="sxs-lookup"><span data-stu-id="46cb7-187">40px</span></span>|<span data-ttu-id="46cb7-188">48px</span><span class="sxs-lookup"><span data-stu-id="46cb7-188">48px</span></span>|<span data-ttu-id="46cb7-189">64px</span><span class="sxs-lookup"><span data-stu-id="46cb7-189">64px</span></span>|<span data-ttu-id="46cb7-190">80px</span><span class="sxs-lookup"><span data-stu-id="46cb7-190">80px</span></span>|<span data-ttu-id="46cb7-191">96px</span><span class="sxs-lookup"><span data-stu-id="46cb7-191">96px</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|<span data-ttu-id="46cb7-192">**Espessura da linha:**</span><span class="sxs-lookup"><span data-stu-id="46cb7-192">**Line Weight:**</span></span>|<span data-ttu-id="46cb7-193">1px</span><span class="sxs-lookup"><span data-stu-id="46cb7-193">1px</span></span>|<span data-ttu-id="46cb7-194">1px</span><span class="sxs-lookup"><span data-stu-id="46cb7-194">1px</span></span>|<span data-ttu-id="46cb7-195">1px</span><span class="sxs-lookup"><span data-stu-id="46cb7-195">1px</span></span>|<span data-ttu-id="46cb7-196">1px</span><span class="sxs-lookup"><span data-stu-id="46cb7-196">1px</span></span>|<span data-ttu-id="46cb7-197">2px</span><span class="sxs-lookup"><span data-stu-id="46cb7-197">2px</span></span>|<span data-ttu-id="46cb7-198">2px</span><span class="sxs-lookup"><span data-stu-id="46cb7-198">2px</span></span>|<span data-ttu-id="46cb7-199">2px</span><span class="sxs-lookup"><span data-stu-id="46cb7-199">2px</span></span>|<span data-ttu-id="46cb7-200">2px</span><span class="sxs-lookup"><span data-stu-id="46cb7-200">2px</span></span>|<span data-ttu-id="46cb7-201">3px</span><span class="sxs-lookup"><span data-stu-id="46cb7-201">3px</span></span>|
|<span data-ttu-id="46cb7-202">**Ícone de exemplo:**</span><span class="sxs-lookup"><span data-stu-id="46cb7-202">**Example icon:**</span></span>|![ícone 16 px](../images/monolineicon16.png)|![ícone de 20 px](../images/monolineicon17.png)|![ícone de 24 px](../images/monolineicon18.png)|![ícone da 32 px](../images/monolineicon19.png)|![ícone da 40 px](../images/monolineicon20.png)|![ícone da 48 px](../images/monolineicon21.png)|![ícone da 64 px](../images/monolineicon22.png)|![ícone da 80 px](../images/monolineicon23.png)|![ícone da 96 px](../images/monolineicon24.png)|

#### <a name="cutouts"></a><span data-ttu-id="46cb7-212">Recortes</span><span class="sxs-lookup"><span data-stu-id="46cb7-212">Cutouts</span></span>

<span data-ttu-id="46cb7-213">Quando um elemento Icon é colocado na parte superior de outro elemento, um recorte (do elemento inferior) é usado para fornecer espaço entre os dois elementos, principalmente para fins de legibilidade.</span><span class="sxs-lookup"><span data-stu-id="46cb7-213">When an icon element is placed on top of another element, a cutout (of the bottom element) is used to provide space between the two elements, mainly for readability purposes.</span></span> <span data-ttu-id="46cb7-214">Isso geralmente ocorre quando um modificador é colocado na parte superior de um elemento base, mas também há casos em que nenhum dos elementos é um modificador.</span><span class="sxs-lookup"><span data-stu-id="46cb7-214">This usually happens when a modifier is placed on top of a base element, but there are also cases where neither of the elements is a modifier.</span></span> <span data-ttu-id="46cb7-215">Esses recortes entre os dois elementos são, às vezes, chamados de "Gap".</span><span class="sxs-lookup"><span data-stu-id="46cb7-215">These cutouts between the two elements is sometimes referred to as a "gap".</span></span>

<span data-ttu-id="46cb7-216">O tamanho da lacuna deve ter a mesma largura que a espessura da linha usada nesse tamanho.</span><span class="sxs-lookup"><span data-stu-id="46cb7-216">The size of the gap should be the same width as the line weight used on that size.</span></span> <span data-ttu-id="46cb7-217">Se estiver fazendo um ícone de 16 px, a largura do espaço seria 1 px e, se for um ícone de 48 px, a lacuna deverá ser 2 px.</span><span class="sxs-lookup"><span data-stu-id="46cb7-217">If making a 16 px icon, the gap width would be 1px and if it is a 48 px icon then the gap should be 2px.</span></span> <span data-ttu-id="46cb7-218">O exemplo a seguir mostra um ícone de 32 px com uma lacuna de 1 px entre o modificador e a base subjacente.</span><span class="sxs-lookup"><span data-stu-id="46cb7-218">The following example shows a 32 px icon with a gap of 1px between the modifier and the underlying base.</span></span>

![32 px Icon com uma lacuna de 1 px entre o modificador e a base subjacente](../images/monolineicon25.png)

<span data-ttu-id="46cb7-220">Em alguns casos, a lacuna pode ser aumentada por uma 1/2 px se o modificador tiver uma borda diagonal ou curva e a lacuna padrão não fornecer separação suficiente.</span><span class="sxs-lookup"><span data-stu-id="46cb7-220">In some cases, the gap can be increase by a 1/2 px if the modifier has a diagonal or curved edge and the standard gap doesn't provide enough separation.</span></span> <span data-ttu-id="46cb7-221">Isso provavelmente afetará somente os ícones com espessura de linha 1 PX: 16 px, 20 px, 24 PX e 32 px.</span><span class="sxs-lookup"><span data-stu-id="46cb7-221">This will likely only affect the icons with 1px line weight: 16 px, 20 px, 24 px, and 32 px.</span></span>

#### <a name="background-fills"></a><span data-ttu-id="46cb7-222">Preenchimentos de plano de fundo</span><span class="sxs-lookup"><span data-stu-id="46cb7-222">Background fills</span></span>

<span data-ttu-id="46cb7-223">A maioria dos ícones no conjunto de ícones monoline exige preenchimentos de plano de fundo.</span><span class="sxs-lookup"><span data-stu-id="46cb7-223">Most icons in the Monoline icon set require background fills.</span></span> <span data-ttu-id="46cb7-224">No entanto, há casos em que o objeto não teria um preenchimento naturalmente, portanto, nenhum preenchimento deve ser aplicado.</span><span class="sxs-lookup"><span data-stu-id="46cb7-224">However, there are cases where the object would not naturally have a fill, so no fill should be applied.</span></span> <span data-ttu-id="46cb7-225">Os ícones a seguir têm um preenchimento branco.</span><span class="sxs-lookup"><span data-stu-id="46cb7-225">The following icons have a white fill.</span></span>

![Compilação de cinco ícones com preenchimento branco](../images/monolineicon26.png)

<span data-ttu-id="46cb7-227">Os ícones a seguir não têm preenchimento.</span><span class="sxs-lookup"><span data-stu-id="46cb7-227">The following icons have no fill.</span></span> <span data-ttu-id="46cb7-228">(O ícone de engrenagem é incluído para mostrar que o orifício central não está preenchido.)</span><span class="sxs-lookup"><span data-stu-id="46cb7-228">(The gear icon is included to show that the center hole is not filled.)</span></span>

![Compilação de cinco ícones sem preenchimento](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a><span data-ttu-id="46cb7-230">Práticas recomendadas para preenchimentos</span><span class="sxs-lookup"><span data-stu-id="46cb7-230">Best practices for fills</span></span>

###### <a name="dos"></a><span data-ttu-id="46cb7-231">Ataque</span><span class="sxs-lookup"><span data-stu-id="46cb7-231">Dos:</span></span>

- <span data-ttu-id="46cb7-232">Preencha qualquer elemento que tenha um limite definido e, naturalmente, teria um preenchimento.</span><span class="sxs-lookup"><span data-stu-id="46cb7-232">Fill any element that has a defined boundary, and would naturally have a fill.</span></span>
- <span data-ttu-id="46cb7-233">Use uma forma separada para criar o preenchimento do plano de fundo.</span><span class="sxs-lookup"><span data-stu-id="46cb7-233">Use a separate shape to create the background fill.</span></span>
- <span data-ttu-id="46cb7-234">Usar **preenchimento de plano de fundo** da [paleta de cores](#color).</span><span class="sxs-lookup"><span data-stu-id="46cb7-234">Use **Background Fill** from the [color palette](#color).</span></span>
- <span data-ttu-id="46cb7-235">Manter a separação de pixels entre elementos sobrepostos.</span><span class="sxs-lookup"><span data-stu-id="46cb7-235">Maintain the pixel separation between overlapping elements.</span></span>
- <span data-ttu-id="46cb7-236">Preencher entre vários objetos.</span><span class="sxs-lookup"><span data-stu-id="46cb7-236">Fill between multiple objects.</span></span>

###### <a name="donts"></a><span data-ttu-id="46cb7-237">Permitido</span><span class="sxs-lookup"><span data-stu-id="46cb7-237">Don'ts:</span></span>

- <span data-ttu-id="46cb7-238">Não preencha objetos que não seriam naturalmente preenchidos; por exemplo, um clipe de clipe.</span><span class="sxs-lookup"><span data-stu-id="46cb7-238">Don't fill objects that would not naturally be filled; for example, a paperclip.</span></span>
- <span data-ttu-id="46cb7-239">Não preencha os colchetes.</span><span class="sxs-lookup"><span data-stu-id="46cb7-239">Don't fill brackets.</span></span>
- <span data-ttu-id="46cb7-240">Não preencha números ou caracteres alfabéticos.</span><span class="sxs-lookup"><span data-stu-id="46cb7-240">Don't fill behind numbers or alpha characters.</span></span>

### <a name="color"></a><span data-ttu-id="46cb7-241">Cor</span><span class="sxs-lookup"><span data-stu-id="46cb7-241">Color</span></span>

<span data-ttu-id="46cb7-242">A paleta de cores foi projetada para simplificar e acessibilidade.</span><span class="sxs-lookup"><span data-stu-id="46cb7-242">The color palette has been designed for simplicity and accessibility.</span></span> <span data-ttu-id="46cb7-243">Ele contém 4 cores neutras e duas variações de azul, verde, amarelo, vermelho e roxo.</span><span class="sxs-lookup"><span data-stu-id="46cb7-243">It contains 4 neutral colors and two variations for blue, green, yellow, red, and purple.</span></span> <span data-ttu-id="46cb7-244">A cor laranja não é incluída intencionalmente na paleta de cores do ícone monoline.</span><span class="sxs-lookup"><span data-stu-id="46cb7-244">Orange is intentionally not included in the Monoline icon color palette.</span></span> <span data-ttu-id="46cb7-245">Cada cor deve ser usada de formas específicas, conforme descrito nesta seção.</span><span class="sxs-lookup"><span data-stu-id="46cb7-245">Each color is intended to be used in specific ways as outlined in this section.</span></span>

#### <a name="palette"></a><span data-ttu-id="46cb7-246">Paleta</span><span class="sxs-lookup"><span data-stu-id="46cb7-246">Palette</span></span>

![Quatro tonalidades de cinza em monolinha: cinza escuro para autônomo ou contorno, cinza médio para contorno ou conteúdo, cinza muito claro para preenchimento de plano de fundo e cinza claro para preenchimento](../images/monoline-grayshades.png)

![A paleta de cores em monoline inclui uma tonalidade azul, verde, amarelo, vermelho e roxo para autônomo, contorno e preenchimento](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a><span data-ttu-id="46cb7-249">Como usar cores</span><span class="sxs-lookup"><span data-stu-id="46cb7-249">How to use color</span></span>

<span data-ttu-id="46cb7-250">Na paleta de cores monoline, todas as cores têm variações autônomas, de estrutura de tópicos e de preenchimento.</span><span class="sxs-lookup"><span data-stu-id="46cb7-250">In the Monoline color palette, all colors have Standalone, Outline, and Fill variations.</span></span> <span data-ttu-id="46cb7-251">Geralmente, os elementos são construídos com um preenchimento e uma borda.</span><span class="sxs-lookup"><span data-stu-id="46cb7-251">Generally, elements are constructed with a fill and a border.</span></span> <span data-ttu-id="46cb7-252">As cores são aplicadas em um dos seguintes padrões:</span><span class="sxs-lookup"><span data-stu-id="46cb7-252">The colors are applied in one of the following patterns:</span></span>

- <span data-ttu-id="46cb7-253">A cor autônoma sozinho para objetos que não têm preenchimento.</span><span class="sxs-lookup"><span data-stu-id="46cb7-253">The Standalone color alone for objects that have no fill.</span></span>
- <span data-ttu-id="46cb7-254">A borda usa a cor de contorno e o preenchimento usa a cor de preenchimento.</span><span class="sxs-lookup"><span data-stu-id="46cb7-254">The border uses the Outline color and the fill uses the Fill color.</span></span>
- <span data-ttu-id="46cb7-255">A borda usa a cor autônoma e o preenchimento usa a cor de preenchimento de plano de fundo.</span><span class="sxs-lookup"><span data-stu-id="46cb7-255">The border uses the Standalone color and the fill uses the Background Fill color.</span></span>

<span data-ttu-id="46cb7-256">A seguir estão exemplos de como usar cores.</span><span class="sxs-lookup"><span data-stu-id="46cb7-256">The following are examples of using color.</span></span>

![Compilação de três ícones com cor em uma borda ou preenchimento ou ambos](../images/monolineicon28.png)

<span data-ttu-id="46cb7-258">A situação mais comum será ter um elemento usando cinza escuro autônomo com preenchimento de plano de fundo.</span><span class="sxs-lookup"><span data-stu-id="46cb7-258">The most common situation will be to have an element use Dark Gray Standalone with Background Fill.</span></span>

<span data-ttu-id="46cb7-259">Ao usar um preenchimento colorido, ele sempre deve estar com sua cor de contorno correspondente.</span><span class="sxs-lookup"><span data-stu-id="46cb7-259">When using a colored Fill, it should always be with its corresponding Outline color.</span></span> <span data-ttu-id="46cb7-260">Por exemplo, o preenchimento azul deve ser usado apenas com o contorno azul.</span><span class="sxs-lookup"><span data-stu-id="46cb7-260">For example, Blue Fill should only be used with Blue Outline.</span></span> <span data-ttu-id="46cb7-261">Mas há duas exceções a essa regra geral:</span><span class="sxs-lookup"><span data-stu-id="46cb7-261">But there are two exceptions to this general rule:</span></span>

- <span data-ttu-id="46cb7-262">O preenchimento de plano de fundo pode ser usado com qualquer cor independente.</span><span class="sxs-lookup"><span data-stu-id="46cb7-262">Background Fill can be used with any color Standalone.</span></span>
- <span data-ttu-id="46cb7-263">O preenchimento cinza claro pode ser usado com duas cores de contorno diferentes: cinza escuro ou cinza médio.</span><span class="sxs-lookup"><span data-stu-id="46cb7-263">Light Gray Fill can be used with two different Outline colors: Dark Gray or Medium Gray.</span></span>

#### <a name="when-to-use-color"></a><span data-ttu-id="46cb7-264">Quando usar cores</span><span class="sxs-lookup"><span data-stu-id="46cb7-264">When to use color</span></span>

<span data-ttu-id="46cb7-265">A cor deve ser usada para transmitir o significado do ícone, em vez de um ornamento.</span><span class="sxs-lookup"><span data-stu-id="46cb7-265">Color should be used to convey the meaning of the icon rather than for embellishment.</span></span> <span data-ttu-id="46cb7-266">Ela deve **realçar a ação** para o usuário.</span><span class="sxs-lookup"><span data-stu-id="46cb7-266">It should **highlight the action** to the user.</span></span> <span data-ttu-id="46cb7-267">Quando um modificador é adicionado a um elemento base que tem cor, o elemento base é normalmente transformado em cinza escuro e preenchimento de plano de fundo para que o modificador possa ser o elemento de cor, como o caso abaixo com o modificador "X" sendo adicionado à base da imagem no ícone da extrema esquerda do conjunto a seguir.</span><span class="sxs-lookup"><span data-stu-id="46cb7-267">When a modifier is added to a base element that has color, the base element is typically turned into Dark Gray and Background Fill so that the modifier can be the element of color, such as the case below with the "X" modifier being added to the picture base in the leftmost icon of the following set.</span></span>

![Compilação de cinco ícones que usam cores](../images/monolineicon29.png)

<span data-ttu-id="46cb7-269">Você deve limitar seus ícones a **uma** cor adicional, diferente da estrutura de tópicos e do preenchimento mencionados acima.</span><span class="sxs-lookup"><span data-stu-id="46cb7-269">You should limit your icons to **one** additional color, other than the Outline and Fill mentioned above.</span></span> <span data-ttu-id="46cb7-270">No entanto, é possível usar mais cores se for vital para a metáfora, com um limite de duas cores adicionais além de cinza.</span><span class="sxs-lookup"><span data-stu-id="46cb7-270">However, more colors can be used if it is vital for its metaphor, with a limit of two additional colors other than gray.</span></span> <span data-ttu-id="46cb7-271">Em casos raros, há exceções quando são necessárias mais cores.</span><span class="sxs-lookup"><span data-stu-id="46cb7-271">In rare cases, there are exceptions when more colors are needed.</span></span> <span data-ttu-id="46cb7-272">Estes são bons exemplos de ícones que usam apenas uma cor.</span><span class="sxs-lookup"><span data-stu-id="46cb7-272">The following are good examples of icons that use just one color.</span></span>

  ![Compilação de cinco ícones que cada um usa uma cor](../images/monolineicon30.png)

<span data-ttu-id="46cb7-274">Mas os ícones a seguir usam muitas cores.</span><span class="sxs-lookup"><span data-stu-id="46cb7-274">But the following icons use too many colors.</span></span>

  ![Compilação de cinco ícones que cada um usa várias cores](../images/monolineicon31.png)

<span data-ttu-id="46cb7-276">Use **cinza médio** para "conteúdo" interno, como linhas de grade em um ícone de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="46cb7-276">Use **Medium Gray** for interior "content", such as grid lines in an icon of a spreadsheet.</span></span> <span data-ttu-id="46cb7-277">Cores interiores adicionais são usadas quando o conteúdo precisa mostrar o comportamento do controle.</span><span class="sxs-lookup"><span data-stu-id="46cb7-277">Additional interior colors are used when the content needs to show the behavior of the control.</span></span>

![Compilação de cinco ícones com elementos interiores de cinza médio](../images/monolineicon32.png)

#### <a name="text-lines"></a><span data-ttu-id="46cb7-279">Linhas de texto</span><span class="sxs-lookup"><span data-stu-id="46cb7-279">Text lines</span></span>

<span data-ttu-id="46cb7-280">Quando as linhas de texto estão em um "contêiner" (por exemplo, texto em um documento), use cinza médio.</span><span class="sxs-lookup"><span data-stu-id="46cb7-280">When text lines are in a "container" (for example, text on a document), use medium gray.</span></span> <span data-ttu-id="46cb7-281">As linhas de texto que não estão em um contêiner devem ser **cinza escuro**.</span><span class="sxs-lookup"><span data-stu-id="46cb7-281">Text lines not in a container should be **Dark Gray**.</span></span>

### <a name="text"></a><span data-ttu-id="46cb7-282">Texto</span><span class="sxs-lookup"><span data-stu-id="46cb7-282">Text</span></span>

<span data-ttu-id="46cb7-283">Evite usar caracteres de texto em ícones.</span><span class="sxs-lookup"><span data-stu-id="46cb7-283">Avoid using text characters in icons.</span></span> <span data-ttu-id="46cb7-284">Como os produtos do Office são usados em todo o mundo, desejamos manter os ícones da forma mais neutra possível.</span><span class="sxs-lookup"><span data-stu-id="46cb7-284">Since Office products are used around the world, we want to keep icons as language neutral as possible.</span></span>

## <a name="production"></a><span data-ttu-id="46cb7-285">Produção</span><span class="sxs-lookup"><span data-stu-id="46cb7-285">Production</span></span>

### <a name="icon-file-format"></a><span data-ttu-id="46cb7-286">Formato de arquivo de ícone</span><span class="sxs-lookup"><span data-stu-id="46cb7-286">Icon file format</span></span>

<span data-ttu-id="46cb7-287">Os ícones finais devem ser salvos como arquivos de imagem. png.</span><span class="sxs-lookup"><span data-stu-id="46cb7-287">The final icons should be saved as .png image files.</span></span> <span data-ttu-id="46cb7-288">Use o formato PNG com um plano de fundo transparente e tenha profundidade de 32 bits.</span><span class="sxs-lookup"><span data-stu-id="46cb7-288">Use PNG format with a transparent background and have 32-bit depth.</span></span>
