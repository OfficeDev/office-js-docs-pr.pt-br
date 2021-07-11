---
title: Diretrizes de ícone de estilo monolinha para Office de complementos
description: Diretrizes para usar ícones de estilo monoline em Office de complementos.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: a59574f1f49ccb8b7b6fd485d08f83e39d760a48
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349340"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a><span data-ttu-id="a041a-103">Diretrizes de ícone de estilo monolinha para Office de complementos</span><span class="sxs-lookup"><span data-stu-id="a041a-103">Monoline style icon guidelines for Office Add-ins</span></span>

<span data-ttu-id="a041a-104">Iconografia de estilo monoline é usada em Office aplicativos.</span><span class="sxs-lookup"><span data-stu-id="a041a-104">Monoline style iconography are used in Office apps.</span></span> <span data-ttu-id="a041a-105">Se você preferir que seus ícones corresponderem ao estilo Fresh de não assinatura Office 2013+, consulte Diretrizes de ícone de estilo novo [para Office Add-ins](add-in-icons-fresh.md).</span><span class="sxs-lookup"><span data-stu-id="a041a-105">If you would prefer that your icons match the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).</span></span>

## <a name="office-monoline-visual-style"></a><span data-ttu-id="a041a-106">Office Estilo visual monoline</span><span class="sxs-lookup"><span data-stu-id="a041a-106">Office Monoline visual style</span></span>

<span data-ttu-id="a041a-107">O objetivo do estilo Monoline é ter uma iconografia consistente, clara e acessível para comunicar ações e recursos com elementos visuais simples, garantir que os ícones sejam acessíveis a todos os usuários e tenham um estilo consistente com os usados em outros locais Windows.</span><span class="sxs-lookup"><span data-stu-id="a041a-107">The goal of the Monoline style to have consistent, clear, and accessible iconography to communicate action and features with simple visuals, ensure the icons are accessible to all users, and have a style that is consistent with those used elsewhere in Windows.</span></span>

<span data-ttu-id="a041a-108">As diretrizes a seguir são para desenvolvedores de terceiros que querem criar ícones para recursos que serão consistentes com os ícones já presentes Office produtos.</span><span class="sxs-lookup"><span data-stu-id="a041a-108">The following guidelines are for 3rd party developers who want to create icons for features that will be consistent with the icons already present Office products.</span></span>

### <a name="design-principles"></a><span data-ttu-id="a041a-109">Princípios de design</span><span class="sxs-lookup"><span data-stu-id="a041a-109">Design principles</span></span>

- <span data-ttu-id="a041a-110">Simples, limpo, claro.</span><span class="sxs-lookup"><span data-stu-id="a041a-110">Simple, clean, clear.</span></span>
- <span data-ttu-id="a041a-111">Contém apenas elementos necessários.</span><span class="sxs-lookup"><span data-stu-id="a041a-111">Contain only necessary elements.</span></span>
- <span data-ttu-id="a041a-112">Inspirado no estilo Windows ícone.</span><span class="sxs-lookup"><span data-stu-id="a041a-112">Inspired by Windows icon style.</span></span>
- <span data-ttu-id="a041a-113">Acessível a todos os usuários.</span><span class="sxs-lookup"><span data-stu-id="a041a-113">Accessible to all users.</span></span>

#### <a name="conveying-meaning"></a><span data-ttu-id="a041a-114">Transmitindo significado</span><span class="sxs-lookup"><span data-stu-id="a041a-114">Conveying meaning</span></span>

- <span data-ttu-id="a041a-115">Use elementos descritivos, como uma página, para representar um documento ou um envelope para representar o email.</span><span class="sxs-lookup"><span data-stu-id="a041a-115">Use descriptive elements such as a page to represent a document or an envelope to represent mail.</span></span>
- <span data-ttu-id="a041a-116">Use o mesmo elemento para representar o mesmo conceito, ou seja, o email é sempre representado por um envelope, não por um carimbo.</span><span class="sxs-lookup"><span data-stu-id="a041a-116">Use the same element to represent the same concept, i.e., mail is always represented by an envelope, not a stamp.</span></span>
- <span data-ttu-id="a041a-117">Use uma metáfora principal durante o desenvolvimento de conceitos.</span><span class="sxs-lookup"><span data-stu-id="a041a-117">Use a core metaphor during concept development.</span></span>

#### <a name="reduction-of-elements"></a><span data-ttu-id="a041a-118">Redução de elementos</span><span class="sxs-lookup"><span data-stu-id="a041a-118">Reduction of Elements</span></span>

- <span data-ttu-id="a041a-119">Reduza o ícone ao seu significado principal, usando apenas elementos essenciais à metáfora.</span><span class="sxs-lookup"><span data-stu-id="a041a-119">Reduce the icon to its core meaning, using only elements that are essential to the metaphor.</span></span>
- <span data-ttu-id="a041a-120">Limite o número de elementos em um ícone para dois, independentemente do tamanho do ícone.</span><span class="sxs-lookup"><span data-stu-id="a041a-120">Limit the number of elements in an icon to two, regardless of icon size.</span></span>

#### <a name="consistency"></a><span data-ttu-id="a041a-121">Consistência</span><span class="sxs-lookup"><span data-stu-id="a041a-121">Consistency</span></span>

<span data-ttu-id="a041a-122">Tamanhos, disposição e cor dos ícones devem ser consistentes.</span><span class="sxs-lookup"><span data-stu-id="a041a-122">Sizes, arrangement, and color of icons should be consistent.</span></span>

#### <a name="styling"></a><span data-ttu-id="a041a-123">Estilo</span><span class="sxs-lookup"><span data-stu-id="a041a-123">Styling</span></span>

##### <a name="perspective"></a><span data-ttu-id="a041a-124">Perspectiva</span><span class="sxs-lookup"><span data-stu-id="a041a-124">Perspective</span></span>

<span data-ttu-id="a041a-125">Os ícones monoline são voltados para frente por padrão.</span><span class="sxs-lookup"><span data-stu-id="a041a-125">Monoline icons are forward-facing by default.</span></span> <span data-ttu-id="a041a-126">Certos elementos que exigem perspectiva e/ou rotação, como um cubo, são permitidos, mas as exceções devem ser mantidas no mínimo.</span><span class="sxs-lookup"><span data-stu-id="a041a-126">Certain elements that require perspective and/or rotation, such as a cube, are allowed, but exceptions should be kept to a minimum.</span></span>

##### <a name="embellishment"></a><span data-ttu-id="a041a-127">Embelezamento</span><span class="sxs-lookup"><span data-stu-id="a041a-127">Embellishment</span></span>

<span data-ttu-id="a041a-128">Monoline é um estilo mínimo limpo.</span><span class="sxs-lookup"><span data-stu-id="a041a-128">Monoline is a clean minimal style.</span></span> <span data-ttu-id="a041a-129">Tudo usa cor plana, o que significa que não há gradientes, texturas ou fontes de luz.</span><span class="sxs-lookup"><span data-stu-id="a041a-129">Everything uses flat color, which means there are no gradients, textures, or light sources.</span></span>

## <a name="designing"></a><span data-ttu-id="a041a-130">Designing</span><span class="sxs-lookup"><span data-stu-id="a041a-130">Designing</span></span>

### <a name="sizes"></a><span data-ttu-id="a041a-131">Tamanhos</span><span class="sxs-lookup"><span data-stu-id="a041a-131">Sizes</span></span>

<span data-ttu-id="a041a-132">Recomendamos que você produza cada ícone em todos esses tamanhos para dar suporte a dispositivos DPI altos.</span><span class="sxs-lookup"><span data-stu-id="a041a-132">We recommend that you produce each icon in all these sizes to support high DPI devices.</span></span> <span data-ttu-id="a041a-133">Os *tamanhos absolutamente* necessários são 16 px, 20 px e 32 px, pois esses são os tamanhos 100%.</span><span class="sxs-lookup"><span data-stu-id="a041a-133">The absolutely *required* sizes are 16 px, 20 px, and 32 px, as those are the 100% sizes.</span></span>

<span data-ttu-id="a041a-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span><span class="sxs-lookup"><span data-stu-id="a041a-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a041a-135">Para uma imagem que é o ícone representativo do seu complemento, consulte [Create effective listings in AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) and within Office for size and other requirements.</span><span class="sxs-lookup"><span data-stu-id="a041a-135">For an image that is your add-in's representative icon, see [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) for size and other requirements.</span></span>

### <a name="layout"></a><span data-ttu-id="a041a-136">Layout</span><span class="sxs-lookup"><span data-stu-id="a041a-136">Layout</span></span>

<span data-ttu-id="a041a-137">A seguir, um exemplo de layout de ícone com um modificador.</span><span class="sxs-lookup"><span data-stu-id="a041a-137">The following is an example of icon layout with a modifier.</span></span>

![Diagrama do ícone com modificador no canto inferior direito.](../images/monolineicon1.png)  ![Diagrama do mesmo ícone com plano de fundo de grade adicionado e textos explicadores para a base, modificador, preenchimento e recorte.](../images/monolineicon2.png)

#### <a name="elements"></a><span data-ttu-id="a041a-140">Elementos</span><span class="sxs-lookup"><span data-stu-id="a041a-140">Elements</span></span>

- <span data-ttu-id="a041a-141">**Base**: O conceito principal que o ícone representa.</span><span class="sxs-lookup"><span data-stu-id="a041a-141">**Base**: The main concept that the icon represents.</span></span> <span data-ttu-id="a041a-142">Normalmente, esse é o único visual necessário para o ícone, mas às vezes o conceito principal pode ser aprimorado com um elemento secundário, um modificador.</span><span class="sxs-lookup"><span data-stu-id="a041a-142">This is usually the only visual needed for the icon, but sometimes the main concept can be enhanced with a secondary element, a modifier.</span></span>

- <span data-ttu-id="a041a-143">**Modificador** Qualquer elemento que sobrepõe a base; ou seja, um modificador que normalmente representa uma ação ou um status.</span><span class="sxs-lookup"><span data-stu-id="a041a-143">**Modifier** Any element that overlays the base; that is, a modifier that typically represents an action or a status.</span></span> <span data-ttu-id="a041a-144">Modifica o elemento base agindo como uma adição, alteração ou descritor.</span><span class="sxs-lookup"><span data-stu-id="a041a-144">It modifies the base element by acting as an addition, alteration, or a descriptor.</span></span>

![Diagrama de grade com áreas base e modificadora chamadas.](../images/monolineicon3.png)

### <a name="construction"></a><span data-ttu-id="a041a-146">Construção</span><span class="sxs-lookup"><span data-stu-id="a041a-146">Construction</span></span>

#### <a name="element-placement"></a><span data-ttu-id="a041a-147">Posicionamento do elemento</span><span class="sxs-lookup"><span data-stu-id="a041a-147">Element placement</span></span>

<span data-ttu-id="a041a-148">Os elementos base são colocados no centro do ícone dentro do preenchimento.</span><span class="sxs-lookup"><span data-stu-id="a041a-148">Base elements are placed in the center of the icon within the padding.</span></span> <span data-ttu-id="a041a-149">Se não puder ser colocado perfeitamente centralizado, a base deve errá-la para a direita superior.</span><span class="sxs-lookup"><span data-stu-id="a041a-149">If it can't be placed perfectly centered, then the base should err to the top right.</span></span> <span data-ttu-id="a041a-150">No exemplo a seguir, o ícone é perfeitamente centralizado.</span><span class="sxs-lookup"><span data-stu-id="a041a-150">In the following example, the icon is perfectly centered.</span></span>

![Diagrama mostrando ícone perfeitamente centralizado.](../images/monolineicon4.png)

<span data-ttu-id="a041a-152">No exemplo a seguir, o ícone está errando para a esquerda.</span><span class="sxs-lookup"><span data-stu-id="a041a-152">In the following example, the icon is erring to the left.</span></span>

![Diagrama mostrando o ícone que erra para a esquerda por 1 px.](../images/monolineicon5.png)

<span data-ttu-id="a041a-154">Os modificadores quase sempre são colocados no canto inferior direito da tela de ícone.</span><span class="sxs-lookup"><span data-stu-id="a041a-154">Modifiers are almost always placed in the bottom right corner of the icon canvas.</span></span> <span data-ttu-id="a041a-155">Em alguns casos raros, os modificadores são colocados em um canto diferente.</span><span class="sxs-lookup"><span data-stu-id="a041a-155">In some rare cases, modifiers are placed in a different corner.</span></span> <span data-ttu-id="a041a-156">Por exemplo, se o elemento base não for reconhecido com o modificador no canto inferior direito, considere colocá-lo no canto superior esquerdo.</span><span class="sxs-lookup"><span data-stu-id="a041a-156">For example, if the base element would be unrecognizable with the modifier in the bottom right corner, then consider placing it in the upper left corner.</span></span>

![Diagrama mostrando quatro ícones com o modificador na parte inferior direita e um ícone com o modificador no canto superior esquerdo.](../images/monolineicon6.png)

#### <a name="padding"></a><span data-ttu-id="a041a-158">Padding</span><span class="sxs-lookup"><span data-stu-id="a041a-158">Padding</span></span>

<span data-ttu-id="a041a-159">Cada ícone de tamanho tem uma quantidade especificada de preenchimento ao redor do ícone.</span><span class="sxs-lookup"><span data-stu-id="a041a-159">Each size icon has a specified amount of padding around the icon.</span></span> <span data-ttu-id="a041a-160">O elemento base permanece dentro do preenchimento, mas o modificador deve ficar até a borda da tela, estendendo-se fora do preenchimento até a borda da borda do ícone.</span><span class="sxs-lookup"><span data-stu-id="a041a-160">The base element stays within the padding, but the modifier should butt up to the edge of the canvas, extending outside of the padding to the edge of the icon border.</span></span> <span data-ttu-id="a041a-161">As imagens a seguir mostram o preenchimento recomendado a ser usado para cada um dos tamanhos de ícone.</span><span class="sxs-lookup"><span data-stu-id="a041a-161">The following images show the recommended padding to use for each of the icon sizes.</span></span>

|<span data-ttu-id="a041a-162">**16px**</span><span class="sxs-lookup"><span data-stu-id="a041a-162">**16px**</span></span>|<span data-ttu-id="a041a-163">**20px**</span><span class="sxs-lookup"><span data-stu-id="a041a-163">**20px**</span></span>|<span data-ttu-id="a041a-164">**24px**</span><span class="sxs-lookup"><span data-stu-id="a041a-164">**24px**</span></span>|<span data-ttu-id="a041a-165">**32px**</span><span class="sxs-lookup"><span data-stu-id="a041a-165">**32px**</span></span>|<span data-ttu-id="a041a-166">**40px**</span><span class="sxs-lookup"><span data-stu-id="a041a-166">**40px**</span></span>|<span data-ttu-id="a041a-167">**48px**</span><span class="sxs-lookup"><span data-stu-id="a041a-167">**48px**</span></span>|<span data-ttu-id="a041a-168">**64px**</span><span class="sxs-lookup"><span data-stu-id="a041a-168">**64px**</span></span>|<span data-ttu-id="a041a-169">**80px**</span><span class="sxs-lookup"><span data-stu-id="a041a-169">**80px**</span></span>|<span data-ttu-id="a041a-170">**96px**</span><span class="sxs-lookup"><span data-stu-id="a041a-170">**96px**</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![Ícone de 16 px com preenchimento 0px.](../images/monolineicon7.png)|![Ícone de 20 px com preenchimento de 1px.](../images/monolineicon8.png)|![Ícone de 24 px com preenchimento de 1px.](../images/monolineicon9.png)|![Ícone de 32 px com preenchimento de 2px.](../images/monolineicon10.png)|![Ícone de 40 px com preenchimento de 2px.](../images/monolineicon11.png)|![Ícone de 48 px com preenchimento de 3px.](../images/monolineicon12.png)|![Ícone de 64 px com preenchimento de 4px.](../images/monolineicon13.png)|![Ícone de 80 px com preenchimento de 5px.](../images/monolineicon14.png)|![Ícone px de 96 com preenchimento de 6px.](../images/monolineicon15.png)|

#### <a name="line-weights"></a><span data-ttu-id="a041a-180">Pesos de linha</span><span class="sxs-lookup"><span data-stu-id="a041a-180">Line weights</span></span>

<span data-ttu-id="a041a-181">Monoline é um estilo dominado por formas de linha e delineadas.</span><span class="sxs-lookup"><span data-stu-id="a041a-181">Monoline is a style dominated by line and outlined shapes.</span></span> <span data-ttu-id="a041a-182">Dependendo do tamanho que você está produzindo, o ícone deve usar os seguintes pesos de linha.</span><span class="sxs-lookup"><span data-stu-id="a041a-182">Depending on what size you are producing the icon should use the following line weights.</span></span>

|<span data-ttu-id="a041a-183">Tamanho do ícone:</span><span class="sxs-lookup"><span data-stu-id="a041a-183">Icon Size:</span></span>|<span data-ttu-id="a041a-184">16px</span><span class="sxs-lookup"><span data-stu-id="a041a-184">16px</span></span>|<span data-ttu-id="a041a-185">20px</span><span class="sxs-lookup"><span data-stu-id="a041a-185">20px</span></span>|<span data-ttu-id="a041a-186">24px</span><span class="sxs-lookup"><span data-stu-id="a041a-186">24px</span></span>|<span data-ttu-id="a041a-187">32px</span><span class="sxs-lookup"><span data-stu-id="a041a-187">32px</span></span>|<span data-ttu-id="a041a-188">40px</span><span class="sxs-lookup"><span data-stu-id="a041a-188">40px</span></span>|<span data-ttu-id="a041a-189">48px</span><span class="sxs-lookup"><span data-stu-id="a041a-189">48px</span></span>|<span data-ttu-id="a041a-190">64px</span><span class="sxs-lookup"><span data-stu-id="a041a-190">64px</span></span>|<span data-ttu-id="a041a-191">80px</span><span class="sxs-lookup"><span data-stu-id="a041a-191">80px</span></span>|<span data-ttu-id="a041a-192">96px</span><span class="sxs-lookup"><span data-stu-id="a041a-192">96px</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|<span data-ttu-id="a041a-193">**Peso da linha:**</span><span class="sxs-lookup"><span data-stu-id="a041a-193">**Line Weight:**</span></span>|<span data-ttu-id="a041a-194">1px</span><span class="sxs-lookup"><span data-stu-id="a041a-194">1px</span></span>|<span data-ttu-id="a041a-195">1px</span><span class="sxs-lookup"><span data-stu-id="a041a-195">1px</span></span>|<span data-ttu-id="a041a-196">1px</span><span class="sxs-lookup"><span data-stu-id="a041a-196">1px</span></span>|<span data-ttu-id="a041a-197">1px</span><span class="sxs-lookup"><span data-stu-id="a041a-197">1px</span></span>|<span data-ttu-id="a041a-198">2px</span><span class="sxs-lookup"><span data-stu-id="a041a-198">2px</span></span>|<span data-ttu-id="a041a-199">2px</span><span class="sxs-lookup"><span data-stu-id="a041a-199">2px</span></span>|<span data-ttu-id="a041a-200">2px</span><span class="sxs-lookup"><span data-stu-id="a041a-200">2px</span></span>|<span data-ttu-id="a041a-201">2px</span><span class="sxs-lookup"><span data-stu-id="a041a-201">2px</span></span>|<span data-ttu-id="a041a-202">3px</span><span class="sxs-lookup"><span data-stu-id="a041a-202">3px</span></span>|
|<span data-ttu-id="a041a-203">**Ícone de exemplo:**</span><span class="sxs-lookup"><span data-stu-id="a041a-203">**Example icon:**</span></span>|![Ícone de 16 px.](../images/monolineicon16.png)|![Ícone de 20 px.](../images/monolineicon17.png)|![Ícone de 24 px.](../images/monolineicon18.png)|![Ícone de 32 px.](../images/monolineicon19.png)|![Ícone de 40 px.](../images/monolineicon20.png)|![Ícone de 48 px.](../images/monolineicon21.png)|![Ícone de 64 px.](../images/monolineicon22.png)|![Ícone de 80 px.](../images/monolineicon23.png)|![Ícone px de 96.](../images/monolineicon24.png)|

#### <a name="cutouts"></a><span data-ttu-id="a041a-213">Recortes</span><span class="sxs-lookup"><span data-stu-id="a041a-213">Cutouts</span></span>

<span data-ttu-id="a041a-214">Quando um elemento icon é colocado sobre outro elemento, um recorte (do elemento inferior) é usado para fornecer espaço entre os dois elementos, principalmente para fins de leitura.</span><span class="sxs-lookup"><span data-stu-id="a041a-214">When an icon element is placed on top of another element, a cutout (of the bottom element) is used to provide space between the two elements, mainly for readability purposes.</span></span> <span data-ttu-id="a041a-215">Isso geralmente acontece quando um modificador é colocado sobre um elemento base, mas também há casos em que nenhum dos elementos é um modificador.</span><span class="sxs-lookup"><span data-stu-id="a041a-215">This usually happens when a modifier is placed on top of a base element, but there are also cases where neither of the elements is a modifier.</span></span> <span data-ttu-id="a041a-216">Esses recortes entre os dois elementos são às vezes chamados de "lacuna".</span><span class="sxs-lookup"><span data-stu-id="a041a-216">These cutouts between the two elements is sometimes referred to as a "gap".</span></span>

<span data-ttu-id="a041a-217">O tamanho da lacuna deve ter a mesma largura que o peso da linha usado nesse tamanho.</span><span class="sxs-lookup"><span data-stu-id="a041a-217">The size of the gap should be the same width as the line weight used on that size.</span></span> <span data-ttu-id="a041a-218">Se você criar um ícone de 16 px, a largura da lacuna será 1px e se for um ícone de 48 px, o intervalo deverá ser de 2px.</span><span class="sxs-lookup"><span data-stu-id="a041a-218">If making a 16 px icon, the gap width would be 1px and if it is a 48 px icon then the gap should be 2px.</span></span> <span data-ttu-id="a041a-219">O exemplo a seguir mostra um ícone de 32 px com um intervalo de 1px entre o modificador e a base subjacente.</span><span class="sxs-lookup"><span data-stu-id="a041a-219">The following example shows a 32 px icon with a gap of 1px between the modifier and the underlying base.</span></span>

![Ícone de 32 px com um intervalo de 1px entre o modificador e a base subjacente.](../images/monolineicon25.png)

<span data-ttu-id="a041a-221">Em alguns casos, o intervalo pode ser maior em um px de 1/2 se o modificador tiver uma borda diagonal ou curva e o intervalo padrão não fornecer separação suficiente.</span><span class="sxs-lookup"><span data-stu-id="a041a-221">In some cases, the gap can be increase by a 1/2 px if the modifier has a diagonal or curved edge and the standard gap doesn't provide enough separation.</span></span> <span data-ttu-id="a041a-222">Isso provavelmente afetará apenas os ícones com peso de linha de 1px: 16 px, 20 px, 24 px e 32 px.</span><span class="sxs-lookup"><span data-stu-id="a041a-222">This will likely only affect the icons with 1px line weight: 16 px, 20 px, 24 px, and 32 px.</span></span>

#### <a name="background-fills"></a><span data-ttu-id="a041a-223">Preenchimentos em segundo plano</span><span class="sxs-lookup"><span data-stu-id="a041a-223">Background fills</span></span>

<span data-ttu-id="a041a-224">A maioria dos ícones no conjunto de ícones monoline exige preenchimentos em segundo plano.</span><span class="sxs-lookup"><span data-stu-id="a041a-224">Most icons in the Monoline icon set require background fills.</span></span> <span data-ttu-id="a041a-225">No entanto, há casos em que o objeto não teria naturalmente um preenchimento, portanto, nenhum preenchimento deve ser aplicado.</span><span class="sxs-lookup"><span data-stu-id="a041a-225">However, there are cases where the object would not naturally have a fill, so no fill should be applied.</span></span> <span data-ttu-id="a041a-226">Os ícones a seguir têm um preenchimento em branco.</span><span class="sxs-lookup"><span data-stu-id="a041a-226">The following icons have a white fill.</span></span>

![Compilação de cinco ícones com preenchimento branco.](../images/monolineicon26.png)

<span data-ttu-id="a041a-228">Os ícones a seguir não têm preenchimento.</span><span class="sxs-lookup"><span data-stu-id="a041a-228">The following icons have no fill.</span></span> <span data-ttu-id="a041a-229">(O ícone de engrenagem está incluído para mostrar que o buraco central não está preenchido.)</span><span class="sxs-lookup"><span data-stu-id="a041a-229">(The gear icon is included to show that the center hole is not filled.)</span></span>

![Compilação de cinco ícones sem preenchimento.](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a><span data-ttu-id="a041a-231">Práticas recomendadas para preenchimentos</span><span class="sxs-lookup"><span data-stu-id="a041a-231">Best practices for fills</span></span>

###### <a name="dos"></a><span data-ttu-id="a041a-232">Dos:</span><span class="sxs-lookup"><span data-stu-id="a041a-232">Dos:</span></span>

- <span data-ttu-id="a041a-233">Preencha qualquer elemento que tenha um limite definido e que tenha um preenchimento naturalmente.</span><span class="sxs-lookup"><span data-stu-id="a041a-233">Fill any element that has a defined boundary, and would naturally have a fill.</span></span>
- <span data-ttu-id="a041a-234">Use uma forma separada para criar o preenchimento em segundo plano.</span><span class="sxs-lookup"><span data-stu-id="a041a-234">Use a separate shape to create the background fill.</span></span>
- <span data-ttu-id="a041a-235">Use **o Preenchimento de** Plano de Fundo da [paleta de cores](#color).</span><span class="sxs-lookup"><span data-stu-id="a041a-235">Use **Background Fill** from the [color palette](#color).</span></span>
- <span data-ttu-id="a041a-236">Mantenha a separação de pixels entre elementos sobrepostos.</span><span class="sxs-lookup"><span data-stu-id="a041a-236">Maintain the pixel separation between overlapping elements.</span></span>
- <span data-ttu-id="a041a-237">Preencha entre vários objetos.</span><span class="sxs-lookup"><span data-stu-id="a041a-237">Fill between multiple objects.</span></span>

###### <a name="donts"></a><span data-ttu-id="a041a-238">Não:</span><span class="sxs-lookup"><span data-stu-id="a041a-238">Don'ts:</span></span>

- <span data-ttu-id="a041a-239">Não preencha objetos que não seriam preenchidos naturalmente; por exemplo, um clipe de papel.</span><span class="sxs-lookup"><span data-stu-id="a041a-239">Don't fill objects that would not naturally be filled; for example, a paperclip.</span></span>
- <span data-ttu-id="a041a-240">Não preencha colchetes.</span><span class="sxs-lookup"><span data-stu-id="a041a-240">Don't fill brackets.</span></span>
- <span data-ttu-id="a041a-241">Não preencha por trás de números ou caracteres alfa.</span><span class="sxs-lookup"><span data-stu-id="a041a-241">Don't fill behind numbers or alpha characters.</span></span>

### <a name="color"></a><span data-ttu-id="a041a-242">Cor</span><span class="sxs-lookup"><span data-stu-id="a041a-242">Color</span></span>

<span data-ttu-id="a041a-243">A paleta de cores foi projetada para simplicidade e acessibilidade.</span><span class="sxs-lookup"><span data-stu-id="a041a-243">The color palette has been designed for simplicity and accessibility.</span></span> <span data-ttu-id="a041a-244">Ele contém 4 cores neutras e duas variações para azul, verde, amarelo, vermelho e roxo.</span><span class="sxs-lookup"><span data-stu-id="a041a-244">It contains 4 neutral colors and two variations for blue, green, yellow, red, and purple.</span></span> <span data-ttu-id="a041a-245">Laranja não está incluída intencionalmente na paleta de cores do ícone monoline.</span><span class="sxs-lookup"><span data-stu-id="a041a-245">Orange is intentionally not included in the Monoline icon color palette.</span></span> <span data-ttu-id="a041a-246">Cada cor destina-se a ser usada de maneiras específicas, conforme descrito nesta seção.</span><span class="sxs-lookup"><span data-stu-id="a041a-246">Each color is intended to be used in specific ways as outlined in this section.</span></span>

#### <a name="palette"></a><span data-ttu-id="a041a-247">Paleta</span><span class="sxs-lookup"><span data-stu-id="a041a-247">Palette</span></span>

![Os quatro tons de cinza em monoline: cinza escuro para autônomo ou contorno, cinza médio para contorno ou conteúdo, cinza muito claro para preenchimento de plano de fundo e cinza claro para preenchimento.](../images/monoline-grayshades.png)

![A paleta de cores em monoline inclui um tom de azul, verde, amarelo, vermelho e roxo para autônomo, contorno e preenchimento.](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a><span data-ttu-id="a041a-250">Como usar cor</span><span class="sxs-lookup"><span data-stu-id="a041a-250">How to use color</span></span>

<span data-ttu-id="a041a-251">Na paleta de cores Monoline, todas as cores têm variações Autônomas, Contornos e Preenchimento.</span><span class="sxs-lookup"><span data-stu-id="a041a-251">In the Monoline color palette, all colors have Standalone, Outline, and Fill variations.</span></span> <span data-ttu-id="a041a-252">Geralmente, os elementos são construídos com um preenchimento e uma borda.</span><span class="sxs-lookup"><span data-stu-id="a041a-252">Generally, elements are constructed with a fill and a border.</span></span> <span data-ttu-id="a041a-253">As cores são aplicadas em um dos padrões a seguir.</span><span class="sxs-lookup"><span data-stu-id="a041a-253">The colors are applied in one of the following patterns.</span></span>

- <span data-ttu-id="a041a-254">A cor autônoma sozinha para objetos que não têm preenchimento.</span><span class="sxs-lookup"><span data-stu-id="a041a-254">The Standalone color alone for objects that have no fill.</span></span>
- <span data-ttu-id="a041a-255">A borda usa a cor Outline e o preenchimento usa a cor Preenchimento.</span><span class="sxs-lookup"><span data-stu-id="a041a-255">The border uses the Outline color and the fill uses the Fill color.</span></span>
- <span data-ttu-id="a041a-256">A borda usa a cor Autônoma e o preenchimento usa a cor Preenchimento de Plano de Fundo.</span><span class="sxs-lookup"><span data-stu-id="a041a-256">The border uses the Standalone color and the fill uses the Background Fill color.</span></span>

<span data-ttu-id="a041a-257">A seguir estão exemplos de uso de cor.</span><span class="sxs-lookup"><span data-stu-id="a041a-257">The following are examples of using color.</span></span>

![Compilação de três ícones com cor em uma borda ou preenchimento ou ambos.](../images/monolineicon28.png)

<span data-ttu-id="a041a-259">A situação mais comum será ter um elemento que use Dark Gray Standalone with Background Fill.</span><span class="sxs-lookup"><span data-stu-id="a041a-259">The most common situation will be to have an element use Dark Gray Standalone with Background Fill.</span></span>

<span data-ttu-id="a041a-260">Ao usar um Preenchimento colorido, ele sempre deve estar com sua cor Delineada correspondente.</span><span class="sxs-lookup"><span data-stu-id="a041a-260">When using a colored Fill, it should always be with its corresponding Outline color.</span></span> <span data-ttu-id="a041a-261">Por exemplo, Preenchimento Azul só deve ser usado com o Contorno Azul.</span><span class="sxs-lookup"><span data-stu-id="a041a-261">For example, Blue Fill should only be used with Blue Outline.</span></span> <span data-ttu-id="a041a-262">Mas há duas exceções para esta regra geral:</span><span class="sxs-lookup"><span data-stu-id="a041a-262">But there are two exceptions to this general rule:</span></span>

- <span data-ttu-id="a041a-263">O Preenchimento de Plano de Fundo pode ser usado com qualquer cor Autônoma.</span><span class="sxs-lookup"><span data-stu-id="a041a-263">Background Fill can be used with any color Standalone.</span></span>
- <span data-ttu-id="a041a-264">O Preenchimento Cinza Claro pode ser usado com duas cores de outline diferentes: Cinza Escuro ou Cinza Médio.</span><span class="sxs-lookup"><span data-stu-id="a041a-264">Light Gray Fill can be used with two different Outline colors: Dark Gray or Medium Gray.</span></span>

#### <a name="when-to-use-color"></a><span data-ttu-id="a041a-265">Quando usar cor</span><span class="sxs-lookup"><span data-stu-id="a041a-265">When to use color</span></span>

<span data-ttu-id="a041a-266">A cor deve ser usada para transmitir o significado do ícone em vez de para embelezamento.</span><span class="sxs-lookup"><span data-stu-id="a041a-266">Color should be used to convey the meaning of the icon rather than for embellishment.</span></span> <span data-ttu-id="a041a-267">Ele deve **realçar a ação** para o usuário.</span><span class="sxs-lookup"><span data-stu-id="a041a-267">It should **highlight the action** to the user.</span></span> <span data-ttu-id="a041a-268">Quando um modificador é adicionado a um elemento base que tem cor, o elemento base normalmente é transformado em Cinza Escuro e Preenchimento de Plano de Fundo para que o modificador possa ser o elemento de cor, como o caso abaixo com o modificador "X" sendo adicionado à base de imagem no ícone mais à esquerda do conjunto a seguir.</span><span class="sxs-lookup"><span data-stu-id="a041a-268">When a modifier is added to a base element that has color, the base element is typically turned into Dark Gray and Background Fill so that the modifier can be the element of color, such as the case below with the "X" modifier being added to the picture base in the leftmost icon of the following set.</span></span>

![Compilação de cinco ícones que usam cor.](../images/monolineicon29.png)

<span data-ttu-id="a041a-270">Você deve limitar seus ícones **a uma** cor adicional, diferente das opções Outline e Fill mencionadas acima.</span><span class="sxs-lookup"><span data-stu-id="a041a-270">You should limit your icons to **one** additional color, other than the Outline and Fill mentioned above.</span></span> <span data-ttu-id="a041a-271">No entanto, mais cores podem ser usadas se for vital para sua metáfora, com um limite de duas cores adicionais que não sejam cinza.</span><span class="sxs-lookup"><span data-stu-id="a041a-271">However, more colors can be used if it is vital for its metaphor, with a limit of two additional colors other than gray.</span></span> <span data-ttu-id="a041a-272">Em casos raros, há exceções quando mais cores são necessárias.</span><span class="sxs-lookup"><span data-stu-id="a041a-272">In rare cases, there are exceptions when more colors are needed.</span></span> <span data-ttu-id="a041a-273">A seguir estão bons exemplos de ícones que usam apenas uma cor.</span><span class="sxs-lookup"><span data-stu-id="a041a-273">The following are good examples of icons that use just one color.</span></span>

  ![Compilação de cinco ícones que cada um usa uma cor.](../images/monolineicon30.png)

<span data-ttu-id="a041a-275">Mas os ícones a seguir usam muitas cores.</span><span class="sxs-lookup"><span data-stu-id="a041a-275">But the following icons use too many colors.</span></span>

  ![Compilação de cinco ícones que cada um usa várias cores.](../images/monolineicon31.png)

<span data-ttu-id="a041a-277">Use **Cinza Médio** para "conteúdo" interno, como linhas de grade em um ícone de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="a041a-277">Use **Medium Gray** for interior "content", such as grid lines in an icon of a spreadsheet.</span></span> <span data-ttu-id="a041a-278">Cores internas adicionais são usadas quando o conteúdo precisa mostrar o comportamento do controle.</span><span class="sxs-lookup"><span data-stu-id="a041a-278">Additional interior colors are used when the content needs to show the behavior of the control.</span></span>

![Compilação de cinco ícones com elementos internos cinza médios.](../images/monolineicon32.png)

#### <a name="text-lines"></a><span data-ttu-id="a041a-280">Linhas de texto</span><span class="sxs-lookup"><span data-stu-id="a041a-280">Text lines</span></span>

<span data-ttu-id="a041a-281">Quando as linhas de texto estão em um "contêiner" (por exemplo, texto em um documento), use cinza médio.</span><span class="sxs-lookup"><span data-stu-id="a041a-281">When text lines are in a "container" (for example, text on a document), use medium gray.</span></span> <span data-ttu-id="a041a-282">Linhas de texto que não estão em um contêiner devem ser **Cinza Escuro**.</span><span class="sxs-lookup"><span data-stu-id="a041a-282">Text lines not in a container should be **Dark Gray**.</span></span>

### <a name="text"></a><span data-ttu-id="a041a-283">Texto</span><span class="sxs-lookup"><span data-stu-id="a041a-283">Text</span></span>

<span data-ttu-id="a041a-284">Evite usar caracteres de texto em ícones.</span><span class="sxs-lookup"><span data-stu-id="a041a-284">Avoid using text characters in icons.</span></span> <span data-ttu-id="a041a-285">Como Office produtos são usados em todo o mundo, queremos manter os ícones o mais neutro possível.</span><span class="sxs-lookup"><span data-stu-id="a041a-285">Since Office products are used around the world, we want to keep icons as language neutral as possible.</span></span>

## <a name="production"></a><span data-ttu-id="a041a-286">Produção</span><span class="sxs-lookup"><span data-stu-id="a041a-286">Production</span></span>

### <a name="icon-file-format"></a><span data-ttu-id="a041a-287">Formato de arquivo icon</span><span class="sxs-lookup"><span data-stu-id="a041a-287">Icon file format</span></span>

<span data-ttu-id="a041a-288">Os ícones finais devem ser salvos como arquivos .png imagem.</span><span class="sxs-lookup"><span data-stu-id="a041a-288">The final icons should be saved as .png image files.</span></span> <span data-ttu-id="a041a-289">Use o formato PNG com um plano de fundo transparente e tenha profundidade de 32 bits.</span><span class="sxs-lookup"><span data-stu-id="a041a-289">Use PNG format with a transparent background and have 32-bit depth.</span></span>

## <a name="see-also"></a><span data-ttu-id="a041a-290">Confira também</span><span class="sxs-lookup"><span data-stu-id="a041a-290">See also</span></span>

- [<span data-ttu-id="a041a-291">Elemento de manifesto de ícone</span><span class="sxs-lookup"><span data-stu-id="a041a-291">Icon manifest element</span></span>](../reference/manifest/icon.md)
- [<span data-ttu-id="a041a-292">Elemento de manifesto IconUrl</span><span class="sxs-lookup"><span data-stu-id="a041a-292">IconUrl manifest element</span></span>](../reference/manifest/iconurl.md)
- [<span data-ttu-id="a041a-293">Elemento de manifesto HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="a041a-293">HighResolutionIconUrl manifest element</span></span>](../reference/manifest/highresolutioniconurl.md)
- [<span data-ttu-id="a041a-294">Criar um ícone para o seu suplemento</span><span class="sxs-lookup"><span data-stu-id="a041a-294">Create an icon for your add-in</span></span>](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)
