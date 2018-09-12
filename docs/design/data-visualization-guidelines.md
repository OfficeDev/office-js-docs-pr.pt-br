---
title: Diretrizes de estilo de visualização de dados para Suplementos do Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 27de6b6b2f4352488ad8f63c3b6e1250cbfbb324
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945789"
---
# <a name="data-visualization-style-guidelines-for-office-add-ins"></a><span data-ttu-id="f0960-102">Diretrizes de estilo de visualização de dados para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="f0960-102">Data visualization style guidelines for Office Add-ins</span></span>

<span data-ttu-id="f0960-p101">Boas visualizações de dados ajudam os usuários a encontrarem informações em seus dados. Eles podem usar essas informações para contar histórias que informam e convencem. Este artigo fornece diretrizes para ajudá-lo a criar visualizações de dados eficazes em seus suplementos para o Excel e outros aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="f0960-p101">Good data visualizations help users find insights in their data. They can use those insights to tell stories that inform and persuade. This article provides guidelines to help you design effective data visualizations in your add-ins for Excel and other Office apps.</span></span>

<span data-ttu-id="f0960-p102">Recomendamos que você use o [Office UI Fabric](https://developer.microsoft.com/fabric) para criar o cromado para suas visualizações de dados. O Office UI Fabric inclui estilos e componentes que se integram perfeitamente com a aparência do Office.</span><span class="sxs-lookup"><span data-stu-id="f0960-p102">We recommend that you use [Office UI Fabric](https://developer.microsoft.com/fabric) to create the chrome for your data visualizations. Office UI Fabric includes styles and components that integrate seamlessly with the Office look and feel.</span></span> 

<!--The following figure shows a data visualization in an add-in that uses Fabric.

![Image of a data visualization with Fabric elements applied**](../images/fabric-data-visualization.png) 

-->

## <a name="data-visualization-elements"></a><span data-ttu-id="f0960-108">Elementos de visualização de dados</span><span class="sxs-lookup"><span data-stu-id="f0960-108">Data visualization elements</span></span>

<span data-ttu-id="f0960-109">As visualizações de dados compartilham uma estrutura geral e elementos comuns visuais e interativos, incluindo títulos, rótulos e plotagens de dados, conforme mostrado nas figuras a seguir.</span><span class="sxs-lookup"><span data-stu-id="f0960-109">Data visualizations share a general framework and common visual and interactive elements, including titles, labels, and data plots, as shown in the following figures.</span></span>

<span data-ttu-id="f0960-110">![Imagem de um gráfico de linhas com título, eixos, legenda e uma área de plotagem rotulada](../images/data-visualization-line-chart.png)
![Imagem de um gráfico de coluna com eixos, linhas de grade, legenda e plotagem de dados rotulada](../images/data-visualization-column-chart.png)</span><span class="sxs-lookup"><span data-stu-id="f0960-110">![Image of a line chart with title, axes, legend, and plot area labeled](../images/data-visualization-line-chart.png)
![Image of a column chart with axes, gridlines, legend, and data plot labeled](../images/data-visualization-column-chart.png)</span></span>

### <a name="chart-titles"></a><span data-ttu-id="f0960-111">Títulos de gráfico</span><span class="sxs-lookup"><span data-stu-id="f0960-111">Chart titles</span></span>

<span data-ttu-id="f0960-112">Siga estas diretrizes para títulos de gráfico:</span><span class="sxs-lookup"><span data-stu-id="f0960-112">Follow these guidelines for chart titles:</span></span>

- <span data-ttu-id="f0960-p103">Deixe seus títulos de gráfico bem legíveis. Posicione-os para criar uma hierarquia visual em relação ao restante do gráfico.</span><span class="sxs-lookup"><span data-stu-id="f0960-p103">Make your chart titles easily readable. Position them to create a clear visual hierarchy in relation to the rest of the chart.</span></span>
- <span data-ttu-id="f0960-p104">Em geral, use maiúsculas nas frases (a primeira letra da primeira palavra em letra maiúscula). Para criar o contraste ou reforçar hierarquias, você poderá usar todas em maiúsculas, mas use isso com moderação.</span><span class="sxs-lookup"><span data-stu-id="f0960-p104">In general, use sentence capitalization (capitalize the first word). To create contrast or to reinforce hierarchies, you can use all caps, but all caps should be used sparingly.</span></span>
- <span data-ttu-id="f0960-p105">Incorpore a [typeramp do Office UI Fabric](https://developer.microsoft.com/fabric#/styles/typography) para deixar seus gráficos consistentes com a interface de usuário do Office, que usa o Segoe. Você também pode usar outra fonte para diferenciar o conteúdo do gráfico da interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="f0960-p105">Incorporate the [Office UI Fabric type ramp](https://developer.microsoft.com/fabric#/styles/typography) to make your charts consistent with the Office UI, which uses Segoe. You can also use a different typeface to differentiate chart content from the UI.</span></span>
- <span data-ttu-id="f0960-119">Use tipos sem serifa com contadores grandes.</span><span class="sxs-lookup"><span data-stu-id="f0960-119">Use sans-serif typefaces with large counters.</span></span>

<span data-ttu-id="f0960-p106">Os exemplos a seguir mostram tipos com serifa e sem serifa usados em títulos de gráfico. Observe como o contraste de escala e o uso eficaz do espaço em branco criam uma hierarquia visual forte.</span><span class="sxs-lookup"><span data-stu-id="f0960-p106">The following examples show serif and sans-serif typefaces used in chart titles. Notice how the scale contrast and effective use of white space create a strong visual hierarchy.</span></span>

<span data-ttu-id="f0960-122">![Imagem de uma visualização de dados com fontes com serifa](../images/data-visualization-serif.png)
![Imagem de uma visualização de dados com fontes sem serifa](../images/data-visualization-sans-serif.png)</span><span class="sxs-lookup"><span data-stu-id="f0960-122">![Image of a data visualization with serif font](../images/data-visualization-serif.png)
![Image of a data visualization with sans-serif font](../images/data-visualization-sans-serif.png)</span></span>

### <a name="axis-labels"></a><span data-ttu-id="f0960-123">Rótulos dos eixos</span><span class="sxs-lookup"><span data-stu-id="f0960-123">Axis labels</span></span>

<span data-ttu-id="f0960-p107">Deixe os rótulos dos eixos escuros para serem lidos claramente, com um bom contraste entre as cores do plano de fundo e do texto. Verifique se não estão tão escuros que competem com a tinta dos dados.</span><span class="sxs-lookup"><span data-stu-id="f0960-p107">Make your axis labels dark enough to read clearly, with adequate contrast ratios between the text and background colors. Make sure that they are not so dark that they compete with data ink.</span></span>

<span data-ttu-id="f0960-p108">Cinza claro é mais eficaz para rótulos dos eixos. Se você estiver usando o Fabric, consulte a [Paleta de cores neutras](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="f0960-p108">Light grays are most effective for axis labels. If you’re using Fabric, see the [Neutral Colors palette](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

### <a name="data-ink"></a><span data-ttu-id="f0960-128">Tinta dos dados</span><span class="sxs-lookup"><span data-stu-id="f0960-128">Data ink</span></span>

<span data-ttu-id="f0960-p109">Os pixels que representam os dados reais em um gráfico são chamados de tinta dos dados. Esse deve ser o foco central da visualização. Evite o uso de sombras, contornos pesados ou elementos de design desnecessários que distorcem ou competem com os dados. Use gradientes apenas quando os valores dos dados estiverem vinculados aos valores das cores. Evite gráficos tridimensionais, a menos que um valor mensurável e objetivo seja associado a uma terceira dimensão.</span><span class="sxs-lookup"><span data-stu-id="f0960-p109">The pixels that represent the actual data in a chart are referred to as data ink. This should be the central focus of the visualization. Avoid the use of drop shadows, heavy outlines, or unnecessary design elements that distort or compete with the data. Use gradients only when data values are tied to color values. Avoid three-dimensional charts unless a measurable, objective value is bound to a third dimension.</span></span>

### <a name="color"></a><span data-ttu-id="f0960-134">Cor</span><span class="sxs-lookup"><span data-stu-id="f0960-134">Color</span></span>

<span data-ttu-id="f0960-p110">Escolha cores que acompanham os temas do sistema operacional ou aplicativo em vez de cores codificadas. Ao mesmo tempo, não deixe que as cores que você aplica distorçam os dados. O uso incorreto de cores nas visualizações de dados pode resultar em distorção de dados e leitura incorreta de informações.</span><span class="sxs-lookup"><span data-stu-id="f0960-p110">Choose colors that follow operating system or application themes rather than hardcoded colors. At the same time, make sure that the colors you apply do not distort the data. Misuse of color in data visualizations can result in data distortion and incorrect reading of information.</span></span>

<span data-ttu-id="f0960-138">Para ver as práticas recomendadas para o uso de cores nas visualizações de dados, consulte o seguinte:</span><span class="sxs-lookup"><span data-stu-id="f0960-138">For best practices for use of color in data visualizations, see the following:</span></span>


- [<span data-ttu-id="f0960-139">Por que as cores do arco-íris não são a melhor opção para as visualizações de dados</span><span class="sxs-lookup"><span data-stu-id="f0960-139">Why rainbow colors aren't the best option for data visualizations</span></span>](http://www.poynter.org/2013/why-rainbow-colors-arent-always-the-best-options-for-data-visualizations/224413/)
- [<span data-ttu-id="f0960-140">Color Brewer 2.0: Aviso de Cor para Cartografia</span><span class="sxs-lookup"><span data-stu-id="f0960-140">Color Brewer 2.0: Color Advice for Cartography</span></span>](http://colorbrewer2.org/)
- [<span data-ttu-id="f0960-141">Eu Quero Matiz</span><span class="sxs-lookup"><span data-stu-id="f0960-141">I Want Hue</span></span>](http://tools.medialab.sciences-po.fr/iwanthue/)

### <a name="gridlines"></a><span data-ttu-id="f0960-142">Linhas de grade</span><span class="sxs-lookup"><span data-stu-id="f0960-142">Gridlines</span></span>

<span data-ttu-id="f0960-p111">As linhas de grade geralmente são necessárias para a leitura precisa de um gráfico, mas elas devem ser apresentadas como um elemento visual secundário, aprimorando a tinta dos dados e não competindo com ela. Use linhas de grade estáticas finas e leves, a menos que elas tenham sido projetadas especificamente para alto contraste. Você também pode usar interação para criar linhas de grade dinâmicas, que aparecem no contexto quando um usuário interage com um gráfico.</span><span class="sxs-lookup"><span data-stu-id="f0960-p111">Gridlines are often necessary for accurately reading a chart, but should be presented as a secondary visual element, enhancing the data ink, not competing with it. Make static gridlines thin and light, unless they are designed specifically for high contrast. You can also use interaction to create dynamic, just-in-time gridlines that appear in context when a user interacts with a chart.</span></span>

<span data-ttu-id="f0960-p112">Cinza claro é mais eficaz para linhas de grade. Se você estiver usando o Fabric, consulte a [Paleta de cores neutras](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="f0960-p112">Light grays are most effective for gridlines. If you’re using Fabric, see the [Neutral Colors palette](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

<span data-ttu-id="f0960-148">A imagem a seguir mostra uma visualização de dados com linhas de grade.</span><span class="sxs-lookup"><span data-stu-id="f0960-148">The following image shows a data visualization with gridlines.</span></span>

![Imagem de uma visualização de dados com linhas de grade](../images/data-visualization-gridlines.png)

### <a name="legends"></a><span data-ttu-id="f0960-150">Legendas</span><span class="sxs-lookup"><span data-stu-id="f0960-150">Legends</span></span>

<span data-ttu-id="f0960-151">Adicione legendas, se for necessário:</span><span class="sxs-lookup"><span data-stu-id="f0960-151">Add legends if necessary to:</span></span>

- <span data-ttu-id="f0960-152">Diferenciar as séries</span><span class="sxs-lookup"><span data-stu-id="f0960-152">Distinguish between series</span></span>
- <span data-ttu-id="f0960-153">Apresentar mudanças de escala ou valor</span><span class="sxs-lookup"><span data-stu-id="f0960-153">Present scale or value changes</span></span>

<span data-ttu-id="f0960-p113">Confira se suas legendas aprimoram a tinta dos dados e não competem com ela. Coloque as legendas:</span><span class="sxs-lookup"><span data-stu-id="f0960-p113">Make sure that your legends enhance the data ink and do not compete with it. Place legends:</span></span>


- <span data-ttu-id="f0960-156">Recue à esquerda acima da área de plotagem por padrão, se todos os itens de legenda se ajustarem acima do gráfico.</span><span class="sxs-lookup"><span data-stu-id="f0960-156">Flush left above the plot area by default, if all legend items fit above the chart.</span></span>
- <span data-ttu-id="f0960-157">No lado superior direito da área de plotagem, se todos os itens de legenda não couberem acima do gráfico e use uma barra de rolagem, se necessário.</span><span class="sxs-lookup"><span data-stu-id="f0960-157">On the upper right side of the plot area, if all legend items do not fit above the chart, and make it scrollable, if necessary.</span></span>

<span data-ttu-id="f0960-p114">Para otimizar a legibilidade e a acessibilidade, mapeie os marcadores de legenda para a forma de gráfico relevante. Por exemplo, use marcadores de legenda em círculo para gráfico de bolhas e de plotagem de dispersão. Use marcadores de legenda de segmento de linha para gráficos de linhas.</span><span class="sxs-lookup"><span data-stu-id="f0960-p114">To optimize for readability and accessibility, map legend markers to the relevant chart shape. For example, use circle legend markers for scatter plot and bubble chart legends. Use line segment legend markers for line charts.</span></span>

### <a name="data-labels-and-tooltips"></a><span data-ttu-id="f0960-161">Dicas de ferramenta e rótulos de dados</span><span class="sxs-lookup"><span data-stu-id="f0960-161">Data labels and tooltips</span></span>

<span data-ttu-id="f0960-p115">Verifique se as dicas de ferramentas os e rótulos de dados têm espaço suficiente em branco e variação de tipos. Use algoritmos para minimizar oclusão e conflito. Por exemplo, uma dica de ferramenta pode ser exibida à direita de um ponto de dados por padrão, mas ser exibida à esquerda se forem detectadas bordas à direita.</span><span class="sxs-lookup"><span data-stu-id="f0960-p115">Ensure that data labels and tooltips have adequate white space and type variation. Use algorithms to minimize occlusion and collision. For example, a tooltip might surface to the right of a data point by default, but surface to the left if right edges are detected.</span></span>

## <a name="design-principles"></a><span data-ttu-id="f0960-165">Princípios de design</span><span class="sxs-lookup"><span data-stu-id="f0960-165">Design principles</span></span>

<span data-ttu-id="f0960-166">A equipe de Design do Office criou o conjunto de princípios de design a seguir, que usamos ao criar novas visualizações de dados para o pacote de produtos do Office.</span><span class="sxs-lookup"><span data-stu-id="f0960-166">The Office Design team created the following set of design principles, which we use when designing new data visualizations for the Office product suite.</span></span>

### <a name="visual-design-principles"></a><span data-ttu-id="f0960-167">Princípios de design visual</span><span class="sxs-lookup"><span data-stu-id="f0960-167">Visual design principles</span></span>

- <span data-ttu-id="f0960-p116">As visualizações devem honrar e aprimorar os dados, facilitando a compreensão. Realce os dados, adicionando elementos de suporte somente conforme o necessário para fornecer o contexto. Evite embelezamentos desnecessários (sombras, estruturas de tópicos etc), gráficos desnecessários ou distorção de dados.</span><span class="sxs-lookup"><span data-stu-id="f0960-p116">Visualizations should honor and enhance the data, making it easy to understand. Highlight the data, adding supporting elements only as needed to provide context. Avoid unnecessary embellishments (drop shadows, outlines, etc), chart junk, or data distortion.</span></span>
- <span data-ttu-id="f0960-p117">As visualizações devem incentivar a exploração fornecendo comentários visuais interessantes. Use padrões de interação bem estabelecidos, controles de interface e feedback claro do sistema.</span><span class="sxs-lookup"><span data-stu-id="f0960-p117">Visualizations should encourage exploration by providing rich visual feedback. Use well-established interaction patterns, interface controls, and clear system feedback.</span></span>
- <span data-ttu-id="f0960-p118">Incorpore princípios de design consagrados. Use princípios tipográficos e de design de comunicação visual estabelecidos para aprimorar a forma, a legibilidade e o significado.</span><span class="sxs-lookup"><span data-stu-id="f0960-p118">Embody time-honored design principles. Use established typographic and visual communication design principles to enhance form, readability, and meaning.</span></span>

### <a name="interaction-design-principles"></a><span data-ttu-id="f0960-175">Princípios de design de interação</span><span class="sxs-lookup"><span data-stu-id="f0960-175">Interaction design principles</span></span>

- <span data-ttu-id="f0960-176">Design para permitir a exploração.</span><span class="sxs-lookup"><span data-stu-id="f0960-176">Design to allow for exploration.</span></span>
- <span data-ttu-id="f0960-177">Permitir interações diretas com objetos que revelam novas informações (classificação ao arrastar, por exemplo).</span><span class="sxs-lookup"><span data-stu-id="f0960-177">Allow for direct interactions with objects that reveal new insights (sorting via drag, for example).</span></span>
- <span data-ttu-id="f0960-178">Use modelos de interação simples, diretos e familiares.</span><span class="sxs-lookup"><span data-stu-id="f0960-178">Use simple, direct, familiar interaction models.</span></span>

<span data-ttu-id="f0960-179">Para obter mais informações sobre como criar visualizações de dados interativas e amigáveis, confira [Princípios e armadilhas de interface do usuário](http://uitraps.com/).</span><span class="sxs-lookup"><span data-stu-id="f0960-179">For more information about how to design user-friendly interactive data visualizations, see [UI Tenets and Traps](http://uitraps.com/).</span></span>

### <a name="motion-design-principles"></a><span data-ttu-id="f0960-180">Princípios de design de animação</span><span class="sxs-lookup"><span data-stu-id="f0960-180">Motion design principles</span></span>

<span data-ttu-id="f0960-p119">A animação segue o estímulo. Os elementos visuais devem se mover na mesma direção e com a mesma velocidade. Isso se aplica a:</span><span class="sxs-lookup"><span data-stu-id="f0960-p119">Motion follows stimulus. Visual elements should move in the same direction at the same rate. This applies to:</span></span>

- <span data-ttu-id="f0960-184">Criação do gráfico</span><span class="sxs-lookup"><span data-stu-id="f0960-184">Chart creation</span></span>
- <span data-ttu-id="f0960-185">Transição de um tipo de gráfico para outro</span><span class="sxs-lookup"><span data-stu-id="f0960-185">Transition from one chart type to another chart type</span></span>
- <span data-ttu-id="f0960-186">Filtragem</span><span class="sxs-lookup"><span data-stu-id="f0960-186">Filtering</span></span>
- <span data-ttu-id="f0960-187">Classificação</span><span class="sxs-lookup"><span data-stu-id="f0960-187">Sorting</span></span>
- <span data-ttu-id="f0960-188">Adição ou subtração de dados</span><span class="sxs-lookup"><span data-stu-id="f0960-188">Adding or subtracting data</span></span>
- <span data-ttu-id="f0960-189">Revisão ou segmentação de dados</span><span class="sxs-lookup"><span data-stu-id="f0960-189">Brushing or slicing data</span></span>
- <span data-ttu-id="f0960-190">Redimensionamento de um gráfico</span><span class="sxs-lookup"><span data-stu-id="f0960-190">Resizing a chart</span></span>

<span data-ttu-id="f0960-p120">Crie uma percepção de causalidade. Quando preparar animações:</span><span class="sxs-lookup"><span data-stu-id="f0960-p120">Create a perception of causality. When staging animations:</span></span>

- <span data-ttu-id="f0960-193">Prepare uma coisa de cada vez.</span><span class="sxs-lookup"><span data-stu-id="f0960-193">Stage one thing at a time.</span></span> 
- <span data-ttu-id="f0960-194">Prepare as mudanças nos eixos antes da mudança na tinta dos dados.</span><span class="sxs-lookup"><span data-stu-id="f0960-194">Stage changes to axes before changes to data ink.</span></span>
- <span data-ttu-id="f0960-195">Prepare e anime objetos como um grupo se eles estiverem se movendo na mesma velocidade e na mesma direção.</span><span class="sxs-lookup"><span data-stu-id="f0960-195">Stage and animate objects as a group if they are moving at the same speed in the same direction.</span></span>
- <span data-ttu-id="f0960-p121">Prepare elementos de dados em grupos de não mais do que 4 a 5 objetos. Os espectadores têm dificuldade de acompanhar mais de 4 e 5 objetos independentemente.</span><span class="sxs-lookup"><span data-stu-id="f0960-p121">Stage data elements in groups of no more than 4-5 objects. Viewers have difficulty tracking more than 4-5 objects independently.</span></span>

<span data-ttu-id="f0960-198">A animação adiciona significado.</span><span class="sxs-lookup"><span data-stu-id="f0960-198">Motion adds meaning.</span></span>

- <span data-ttu-id="f0960-199">Animações aumentam a compreensão do usuário das alterações nos dados, fornecem contexto e atuam como uma camada de anotação não verbal.</span><span class="sxs-lookup"><span data-stu-id="f0960-199">Animations increase user comprehension of changes to the data, provide context, and act as a non-verbal annotation layer.</span></span>
- <span data-ttu-id="f0960-200">A animação deve ocorrer em um espaço de coordenadas significativo da visualização.</span><span class="sxs-lookup"><span data-stu-id="f0960-200">Motion should occur in a meaningful coordinate space of the visualization.</span></span>
- <span data-ttu-id="f0960-201">Personalize a animação para o visual.</span><span class="sxs-lookup"><span data-stu-id="f0960-201">Tailor the animation to the visual.</span></span> 
- <span data-ttu-id="f0960-202">Evite animações gratuitas.</span><span class="sxs-lookup"><span data-stu-id="f0960-202">Avoid gratuitous animations.</span></span>

<span data-ttu-id="f0960-203">A animação segue os dados.</span><span class="sxs-lookup"><span data-stu-id="f0960-203">Motion follows data.</span></span>

- <span data-ttu-id="f0960-p122">Preserve o mapeamentos de dados. Se uma área estiver vinculada a uma medida, mantenha essa área na transição.</span><span class="sxs-lookup"><span data-stu-id="f0960-p122">Preserve data mappings. If an area is tied to a measure, maintain that area in transition.</span></span>
- <span data-ttu-id="f0960-p123">Manter uma linguagem de design de animação consistente. Sempre que possível, mapeie a animação da visualização de dados para a linguagem de design de animação do Office. Use animações semelhantes para tipos de gráfico semelhantes.</span><span class="sxs-lookup"><span data-stu-id="f0960-p123">Maintain a consistent animation design language. Where possible, map data visualization animation to existing Office motion design language. Use similar animations for similar chart types.</span></span>

## <a name="accessibility-in-data-visualizations"></a><span data-ttu-id="f0960-209">Acessibilidade nas visualizações de dados</span><span class="sxs-lookup"><span data-stu-id="f0960-209">Accessibility in data visualizations</span></span>

- <span data-ttu-id="f0960-p124">Não use cor como a única maneira de comunicar informações. As pessoas que são daltônicas não serão capazes de interpretar os resultados. Use forma, tamanho e textura, além de cor quando possível para comunicar informações.</span><span class="sxs-lookup"><span data-stu-id="f0960-p124">Do not use color as the only way to communicate information. People who are color blind will not be able to interpret the results. Use shape, size and texture in addition to color when possible to communicate information.</span></span>
- <span data-ttu-id="f0960-213">Torne todos os elementos interativos, como botões de ação ou listas de escolha, acessíveis a partir de um teclado.</span><span class="sxs-lookup"><span data-stu-id="f0960-213">Make all interactive elements, such as push buttons or pick lists, accessible from a keyboard.</span></span>
- <span data-ttu-id="f0960-214">Envie eventos de acessibilidade para leitores de tela para anunciar alterações de foco, dicas de ferramentas e assim por diante.</span><span class="sxs-lookup"><span data-stu-id="f0960-214">Send accessibility events to screen readers to announce focus changes, tooltips, and so on.</span></span>

## <a name="see-also"></a><span data-ttu-id="f0960-215">Veja também</span><span class="sxs-lookup"><span data-stu-id="f0960-215">See also</span></span> 

- [<span data-ttu-id="f0960-216">Dados + Design: uma introdução simples para preparar e  visualizar as informações</span><span class="sxs-lookup"><span data-stu-id="f0960-216">Data + Design: A Simple Introduction to Preparing and Visualizing Information</span></span>](https://infoactive.co/data-design)
- [<span data-ttu-id="f0960-217">As cinco melhores bibliotecas para criar visualizações de dados</span><span class="sxs-lookup"><span data-stu-id="f0960-217">The Five Best Libraries for Building Data Visualizations</span></span>](http://www.fastcompany.com/3029760/the-five-best-libraries-for-building-data-vizualizations)
- [<span data-ttu-id="f0960-218">Exibição Visual de informações quantitativas</span><span class="sxs-lookup"><span data-stu-id="f0960-218">The Visual Display of Quantitative Information</span></span>](https://www.edwardtufte.com/tufte/books_vdqi)
