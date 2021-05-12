---
title: Diretrizes de estilo de visualização de dados para Suplementos do Office
description: Obter algumas práticas recomendadas para visualizar dados em um Office Desem.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: ac32d7f284850fc8daef1fb1588940844123550f
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330175"
---
# <a name="data-visualization-style-guidelines-for-office-add-ins"></a><span data-ttu-id="c9147-103">Diretrizes de estilo de visualização de dados para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="c9147-103">Data visualization style guidelines for Office Add-ins</span></span>

<span data-ttu-id="c9147-p101">Boas visualizações de dados ajudam os usuários a encontrarem informações em seus dados. Eles podem usar essas informações para contar histórias que informam e convencem. Este artigo fornece diretrizes para ajudá-lo a criar visualizações de dados eficazes em seus suplementos para o Excel e outros aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="c9147-p101">Good data visualizations help users find insights in their data. They can use those insights to tell stories that inform and persuade. This article provides guidelines to help you design effective data visualizations in your add-ins for Excel and other Office apps.</span></span>

<span data-ttu-id="c9147-107">Recomendamos que você use [a interface do usuário Fluent](../design/add-in-design.md) para criar o cromado para suas visualizações de dados.</span><span class="sxs-lookup"><span data-stu-id="c9147-107">We recommend that you use [Fluent UI](../design/add-in-design.md) to create the chrome for your data visualizations.</span></span> <span data-ttu-id="c9147-108">A interface do usuário fluente inclui estilos e componentes que se integram perfeitamente à aparência Office aparência.</span><span class="sxs-lookup"><span data-stu-id="c9147-108">Fluent UI includes styles and components that integrate seamlessly with the Office look and feel.</span></span>

## <a name="data-visualization-elements"></a><span data-ttu-id="c9147-109">Elementos de visualização de dados</span><span class="sxs-lookup"><span data-stu-id="c9147-109">Data visualization elements</span></span>

<span data-ttu-id="c9147-110">As visualizações de dados compartilham uma estrutura geral e elementos visuais e interativos comuns, incluindo títulos, rótulos e plotagem de dados, conforme mostrado na figura a seguir.</span><span class="sxs-lookup"><span data-stu-id="c9147-110">Data visualizations share a general framework and common visual and interactive elements, including titles, labels, and data plots, as shown in the following figure.</span></span>

![Gráfico de linha com título, eixos, legenda e área de plotagem rotulada](../images/excel-charts-visualization.png)

### <a name="chart-titles"></a><span data-ttu-id="c9147-112">Títulos de gráfico</span><span class="sxs-lookup"><span data-stu-id="c9147-112">Chart titles</span></span>

<span data-ttu-id="c9147-113">Siga estas diretrizes para títulos de gráfico:</span><span class="sxs-lookup"><span data-stu-id="c9147-113">Follow these guidelines for chart titles:</span></span>

- <span data-ttu-id="c9147-p103">Deixe seus títulos de gráfico bem legíveis. Posicione-os para criar uma hierarquia visual em relação ao restante do gráfico.</span><span class="sxs-lookup"><span data-stu-id="c9147-p103">Make your chart titles easily readable. Position them to create a clear visual hierarchy in relation to the rest of the chart.</span></span>
- <span data-ttu-id="c9147-p104">Em geral, use maiúsculas nas frases (a primeira letra da primeira palavra em letra maiúscula). Para criar o contraste ou reforçar hierarquias, você poderá usar todas em maiúsculas, mas use isso com moderação.</span><span class="sxs-lookup"><span data-stu-id="c9147-p104">In general, use sentence capitalization (capitalize the first word). To create contrast or to reinforce hierarchies, you can use all caps, but all caps should be used sparingly.</span></span>
- <span data-ttu-id="c9147-118">Incorpore [a rampa de tipo](https://developer.microsoft.com/fluentui#/styles/web/typography) de interface do usuário fluente para tornar seus gráficos consistentes com a interface do usuário Office, que usa o Segoe.</span><span class="sxs-lookup"><span data-stu-id="c9147-118">Incorporate the [Fluent UI type ramp](https://developer.microsoft.com/fluentui#/styles/web/typography) to make your charts consistent with the Office UI, which uses Segoe.</span></span> <span data-ttu-id="c9147-119">Você também pode usar outra fonte para diferenciar o conteúdo do gráfico da interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="c9147-119">You can also use a different typeface to differentiate chart content from the UI.</span></span>
- <span data-ttu-id="c9147-120">Use tipos sem serifa com contadores grandes.</span><span class="sxs-lookup"><span data-stu-id="c9147-120">Use sans-serif typefaces with large counters.</span></span>

### <a name="axis-labels"></a><span data-ttu-id="c9147-121">Rótulos dos eixos</span><span class="sxs-lookup"><span data-stu-id="c9147-121">Axis labels</span></span>

<span data-ttu-id="c9147-p106">Deixe os rótulos dos eixos escuros para serem lidos claramente, com um bom contraste entre as cores do plano de fundo e do texto. Verifique se não estão tão escuros que competem com a tinta dos dados.</span><span class="sxs-lookup"><span data-stu-id="c9147-p106">Make your axis labels dark enough to read clearly, with adequate contrast ratios between the text and background colors. Make sure that they are not so dark that they compete with data ink.</span></span>

<span data-ttu-id="c9147-124">Cinza claro é mais eficaz para rótulos dos eixos.</span><span class="sxs-lookup"><span data-stu-id="c9147-124">Light grays are most effective for axis labels.</span></span> <span data-ttu-id="c9147-125">Se você estiver usando a interface do usuário fluente, consulte a [paleta Cores Neutras](https://developer.microsoft.com/fluentui#/styles/web/colors/neutrals).</span><span class="sxs-lookup"><span data-stu-id="c9147-125">If you're using Fluent UI, see the [Neutral Colors palette](https://developer.microsoft.com/fluentui#/styles/web/colors/neutrals).</span></span>

### <a name="data-ink"></a><span data-ttu-id="c9147-126">Tinta dos dados</span><span class="sxs-lookup"><span data-stu-id="c9147-126">Data ink</span></span>

<span data-ttu-id="c9147-p108">Os pixels que representam os dados reais em um gráfico são chamados de tinta dos dados. Esse deve ser o foco central da visualização. Evite o uso de sombras, contornos pesados ou elementos de design desnecessários que distorcem ou competem com os dados. Use gradientes apenas quando os valores dos dados estiverem vinculados aos valores das cores. Evite gráficos tridimensionais, a menos que um valor mensurável e objetivo seja associado a uma terceira dimensão.</span><span class="sxs-lookup"><span data-stu-id="c9147-p108">The pixels that represent the actual data in a chart are referred to as data ink. This should be the central focus of the visualization. Avoid the use of drop shadows, heavy outlines, or unnecessary design elements that distort or compete with the data. Use gradients only when data values are tied to color values. Avoid three-dimensional charts unless a measurable, objective value is bound to a third dimension.</span></span>

### <a name="color"></a><span data-ttu-id="c9147-132">Cor</span><span class="sxs-lookup"><span data-stu-id="c9147-132">Color</span></span>

<span data-ttu-id="c9147-p109">Escolha cores que acompanham os temas do sistema operacional ou aplicativo em vez de cores codificadas. Ao mesmo tempo, não deixe que as cores que você aplica distorçam os dados. O uso incorreto de cores nas visualizações de dados pode resultar em distorção de dados e leitura incorreta de informações.</span><span class="sxs-lookup"><span data-stu-id="c9147-p109">Choose colors that follow operating system or application themes rather than hardcoded colors. At the same time, make sure that the colors you apply do not distort the data. Misuse of color in data visualizations can result in data distortion and incorrect reading of information.</span></span>

<span data-ttu-id="c9147-136">Para ver as práticas recomendadas para o uso de cores nas visualizações de dados, consulte o seguinte:</span><span class="sxs-lookup"><span data-stu-id="c9147-136">For best practices for use of color in data visualizations, see the following:</span></span>

- [<span data-ttu-id="c9147-137">Por que as cores do arco-íris não são a melhor opção para as visualizações de dados</span><span class="sxs-lookup"><span data-stu-id="c9147-137">Why rainbow colors aren't the best option for data visualizations</span></span>](https://www.poynter.org/2013/why-rainbow-colors-arent-always-the-best-options-for-data-visualizations/224413/)
- [<span data-ttu-id="c9147-138">Color Brewer 2.0: Aviso de Cor para Cartografia</span><span class="sxs-lookup"><span data-stu-id="c9147-138">Color Brewer 2.0: Color Advice for Cartography</span></span>](https://colorbrewer2.org/)
- [<span data-ttu-id="c9147-139">Eu Quero Matiz</span><span class="sxs-lookup"><span data-stu-id="c9147-139">I Want Hue</span></span>](https://tools.medialab.sciences-po.fr/iwanthue/)

### <a name="gridlines"></a><span data-ttu-id="c9147-140">Linhas de grade</span><span class="sxs-lookup"><span data-stu-id="c9147-140">Gridlines</span></span>

<span data-ttu-id="c9147-p110">As linhas de grade geralmente são necessárias para a leitura precisa de um gráfico, mas elas devem ser apresentadas como um elemento visual secundário, aprimorando a tinta dos dados e não competindo com ela. Use linhas de grade estáticas finas e leves, a menos que elas tenham sido projetadas especificamente para alto contraste. Você também pode usar interação para criar linhas de grade dinâmicas, que aparecem no contexto quando um usuário interage com um gráfico.</span><span class="sxs-lookup"><span data-stu-id="c9147-p110">Gridlines are often necessary for accurately reading a chart, but should be presented as a secondary visual element, enhancing the data ink, not competing with it. Make static gridlines thin and light, unless they are designed specifically for high contrast. You can also use interaction to create dynamic, just-in-time gridlines that appear in context when a user interacts with a chart.</span></span>

<span data-ttu-id="c9147-144">Cinza claro é mais eficaz para linhas de grade.</span><span class="sxs-lookup"><span data-stu-id="c9147-144">Light grays are most effective for gridlines.</span></span> <span data-ttu-id="c9147-145">Se você estiver usando a interface do usuário fluente, consulte a [paleta Cores Neutras](https://developer.microsoft.com/fluentui#/styles/web/colors/neutrals).</span><span class="sxs-lookup"><span data-stu-id="c9147-145">If you're using Fluent UI, see the [Neutral Colors palette](https://developer.microsoft.com/fluentui#/styles/web/colors/neutrals).</span></span>

<span data-ttu-id="c9147-146">A imagem a seguir mostra uma visualização de dados com linhas de grade.</span><span class="sxs-lookup"><span data-stu-id="c9147-146">The following image shows a data visualization with gridlines.</span></span>

![Visualização de dados do gráfico de linha com linhas de grade](../images/data-visualization.png)

### <a name="legends"></a><span data-ttu-id="c9147-148">Legendas</span><span class="sxs-lookup"><span data-stu-id="c9147-148">Legends</span></span>

<span data-ttu-id="c9147-149">Adicione legendas, se for necessário:</span><span class="sxs-lookup"><span data-stu-id="c9147-149">Add legends if necessary to:</span></span>

- <span data-ttu-id="c9147-150">Diferenciar as séries</span><span class="sxs-lookup"><span data-stu-id="c9147-150">Distinguish between series</span></span>
- <span data-ttu-id="c9147-151">Apresentar mudanças de escala ou valor</span><span class="sxs-lookup"><span data-stu-id="c9147-151">Present scale or value changes</span></span>

<span data-ttu-id="c9147-p112">Confira se suas legendas aprimoram a tinta dos dados e não competem com ela. Coloque as legendas:</span><span class="sxs-lookup"><span data-stu-id="c9147-p112">Make sure that your legends enhance the data ink and do not compete with it. Place legends:</span></span>


- <span data-ttu-id="c9147-154">Recue à esquerda acima da área de plotagem por padrão, se todos os itens de legenda se ajustarem acima do gráfico.</span><span class="sxs-lookup"><span data-stu-id="c9147-154">Flush left above the plot area by default, if all legend items fit above the chart.</span></span>
- <span data-ttu-id="c9147-155">No lado superior direito da área de plotagem, se todos os itens de legenda não couberem acima do gráfico e use uma barra de rolagem, se necessário.</span><span class="sxs-lookup"><span data-stu-id="c9147-155">On the upper right side of the plot area, if all legend items do not fit above the chart, and make it scrollable, if necessary.</span></span>

<span data-ttu-id="c9147-p113">Para otimizar a legibilidade e a acessibilidade, mapeie os marcadores de legenda para a forma de gráfico relevante. Por exemplo, use marcadores de legenda em círculo para gráfico de bolhas e de plotagem de dispersão. Use marcadores de legenda de segmento de linha para gráficos de linhas.</span><span class="sxs-lookup"><span data-stu-id="c9147-p113">To optimize for readability and accessibility, map legend markers to the relevant chart shape. For example, use circle legend markers for scatter plot and bubble chart legends. Use line segment legend markers for line charts.</span></span>

### <a name="data-labels-and-tooltips"></a><span data-ttu-id="c9147-159">Dicas de ferramenta e rótulos de dados</span><span class="sxs-lookup"><span data-stu-id="c9147-159">Data labels and tooltips</span></span>

<span data-ttu-id="c9147-p114">Verifique se as dicas de ferramentas os e rótulos de dados têm espaço suficiente em branco e variação de tipos. Use algoritmos para minimizar oclusão e conflito. Por exemplo, uma dica de ferramenta pode ser exibida à direita de um ponto de dados por padrão, mas ser exibida à esquerda se forem detectadas bordas à direita.</span><span class="sxs-lookup"><span data-stu-id="c9147-p114">Ensure that data labels and tooltips have adequate white space and type variation. Use algorithms to minimize occlusion and collision. For example, a tooltip might surface to the right of a data point by default, but surface to the left if right edges are detected.</span></span>

## <a name="design-principles"></a><span data-ttu-id="c9147-163">Princípios de design</span><span class="sxs-lookup"><span data-stu-id="c9147-163">Design principles</span></span>

<span data-ttu-id="c9147-164">A equipe de Design do Office criou o conjunto de princípios de design a seguir, que usamos ao criar novas visualizações de dados para o pacote de produtos do Office.</span><span class="sxs-lookup"><span data-stu-id="c9147-164">The Office Design team created the following set of design principles, which we use when designing new data visualizations for the Office product suite.</span></span>

### <a name="visual-design-principles"></a><span data-ttu-id="c9147-165">Princípios de design visual</span><span class="sxs-lookup"><span data-stu-id="c9147-165">Visual design principles</span></span>

- <span data-ttu-id="c9147-p115">As visualizações devem honrar e aprimorar os dados, facilitando a compreensão. Realce os dados, adicionando elementos de suporte somente conforme o necessário para fornecer o contexto. Evite embelezamentos desnecessários (sombras, estruturas de tópicos etc), gráficos desnecessários ou distorção de dados.</span><span class="sxs-lookup"><span data-stu-id="c9147-p115">Visualizations should honor and enhance the data, making it easy to understand. Highlight the data, adding supporting elements only as needed to provide context. Avoid unnecessary embellishments (drop shadows, outlines, etc), chart junk, or data distortion.</span></span>
- <span data-ttu-id="c9147-p116">As visualizações devem incentivar a exploração fornecendo comentários visuais interessantes. Use padrões de interação bem estabelecidos, controles de interface e feedback claro do sistema.</span><span class="sxs-lookup"><span data-stu-id="c9147-p116">Visualizations should encourage exploration by providing rich visual feedback. Use well-established interaction patterns, interface controls, and clear system feedback.</span></span>
- <span data-ttu-id="c9147-p117">Incorpore princípios de design consagrados. Use princípios tipográficos e de design de comunicação visual estabelecidos para aprimorar a forma, a legibilidade e o significado.</span><span class="sxs-lookup"><span data-stu-id="c9147-p117">Embody time-honored design principles. Use established typographic and visual communication design principles to enhance form, readability, and meaning.</span></span>

### <a name="interaction-design-principles"></a><span data-ttu-id="c9147-173">Princípios de design de interação</span><span class="sxs-lookup"><span data-stu-id="c9147-173">Interaction design principles</span></span>

- <span data-ttu-id="c9147-174">Design para permitir a exploração.</span><span class="sxs-lookup"><span data-stu-id="c9147-174">Design to allow for exploration.</span></span>
- <span data-ttu-id="c9147-175">Permitir interações diretas com objetos que revelam novas informações (classificação ao arrastar, por exemplo).</span><span class="sxs-lookup"><span data-stu-id="c9147-175">Allow for direct interactions with objects that reveal new insights (sorting via drag, for example).</span></span>
- <span data-ttu-id="c9147-176">Use modelos de interação simples, diretos e familiares.</span><span class="sxs-lookup"><span data-stu-id="c9147-176">Use simple, direct, familiar interaction models.</span></span>

<span data-ttu-id="c9147-177">Para obter mais informações sobre como criar visualizações de dados interativas e amigáveis, confira [Princípios e armadilhas de interface do usuário](https://uitraps.com/).</span><span class="sxs-lookup"><span data-stu-id="c9147-177">For more information about how to design user-friendly interactive data visualizations, see [UI Tenets and Traps](https://uitraps.com/).</span></span>

### <a name="motion-design-principles"></a><span data-ttu-id="c9147-178">Princípios de design de animação</span><span class="sxs-lookup"><span data-stu-id="c9147-178">Motion design principles</span></span>

<span data-ttu-id="c9147-p118">A animação segue o estímulo. Os elementos visuais devem se mover na mesma direção e com a mesma velocidade. Isso se aplica a:</span><span class="sxs-lookup"><span data-stu-id="c9147-p118">Motion follows stimulus. Visual elements should move in the same direction at the same rate. This applies to:</span></span>

- <span data-ttu-id="c9147-182">Criação do gráfico</span><span class="sxs-lookup"><span data-stu-id="c9147-182">Chart creation</span></span>
- <span data-ttu-id="c9147-183">Transição de um tipo de gráfico para outro</span><span class="sxs-lookup"><span data-stu-id="c9147-183">Transition from one chart type to another chart type</span></span>
- <span data-ttu-id="c9147-184">Filtragem</span><span class="sxs-lookup"><span data-stu-id="c9147-184">Filtering</span></span>
- <span data-ttu-id="c9147-185">Classificação</span><span class="sxs-lookup"><span data-stu-id="c9147-185">Sorting</span></span>
- <span data-ttu-id="c9147-186">Adição ou subtração de dados</span><span class="sxs-lookup"><span data-stu-id="c9147-186">Adding or subtracting data</span></span>
- <span data-ttu-id="c9147-187">Revisão ou segmentação de dados</span><span class="sxs-lookup"><span data-stu-id="c9147-187">Brushing or slicing data</span></span>
- <span data-ttu-id="c9147-188">Redimensionamento de um gráfico</span><span class="sxs-lookup"><span data-stu-id="c9147-188">Resizing a chart</span></span>

<span data-ttu-id="c9147-p119">Crie uma percepção de causalidade. Quando preparar animações:</span><span class="sxs-lookup"><span data-stu-id="c9147-p119">Create a perception of causality. When staging animations:</span></span>

- <span data-ttu-id="c9147-191">Prepare uma coisa de cada vez.</span><span class="sxs-lookup"><span data-stu-id="c9147-191">Stage one thing at a time.</span></span>
- <span data-ttu-id="c9147-192">Prepare as mudanças nos eixos antes da mudança na tinta dos dados.</span><span class="sxs-lookup"><span data-stu-id="c9147-192">Stage changes to axes before changes to data ink.</span></span>
- <span data-ttu-id="c9147-193">Prepare e anime objetos como um grupo se eles estiverem se movendo na mesma velocidade e na mesma direção.</span><span class="sxs-lookup"><span data-stu-id="c9147-193">Stage and animate objects as a group if they are moving at the same speed in the same direction.</span></span>
- <span data-ttu-id="c9147-p120">Prepare elementos de dados em grupos de não mais do que 4 a 5 objetos. Os espectadores têm dificuldade de acompanhar mais de 4 e 5 objetos independentemente.</span><span class="sxs-lookup"><span data-stu-id="c9147-p120">Stage data elements in groups of no more than 4-5 objects. Viewers have difficulty tracking more than 4-5 objects independently.</span></span>

<span data-ttu-id="c9147-196">A animação adiciona significado.</span><span class="sxs-lookup"><span data-stu-id="c9147-196">Motion adds meaning.</span></span>

- <span data-ttu-id="c9147-197">Animações aumentam a compreensão do usuário das alterações nos dados, fornecem contexto e atuam como uma camada de anotação não verbal.</span><span class="sxs-lookup"><span data-stu-id="c9147-197">Animations increase user comprehension of changes to the data, provide context, and act as a non-verbal annotation layer.</span></span>
- <span data-ttu-id="c9147-198">A animação deve ocorrer em um espaço de coordenadas significativo da visualização.</span><span class="sxs-lookup"><span data-stu-id="c9147-198">Motion should occur in a meaningful coordinate space of the visualization.</span></span>
- <span data-ttu-id="c9147-199">Personalize a animação para o visual.</span><span class="sxs-lookup"><span data-stu-id="c9147-199">Tailor the animation to the visual.</span></span>
- <span data-ttu-id="c9147-200">Evite animações gratuitas.</span><span class="sxs-lookup"><span data-stu-id="c9147-200">Avoid gratuitous animations.</span></span>

<span data-ttu-id="c9147-201">A animação segue os dados.</span><span class="sxs-lookup"><span data-stu-id="c9147-201">Motion follows data.</span></span>

- <span data-ttu-id="c9147-p121">Preserve o mapeamentos de dados. Se uma área estiver vinculada a uma medida, mantenha essa área na transição.</span><span class="sxs-lookup"><span data-stu-id="c9147-p121">Preserve data mappings. If an area is tied to a measure, maintain that area in transition.</span></span>
- <span data-ttu-id="c9147-p122">Manter uma linguagem de design de animação consistente. Sempre que possível, mapeie a animação da visualização de dados para a linguagem de design de animação do Office. Use animações semelhantes para tipos de gráfico semelhantes.</span><span class="sxs-lookup"><span data-stu-id="c9147-p122">Maintain a consistent animation design language. Where possible, map data visualization animation to existing Office motion design language. Use similar animations for similar chart types.</span></span>

## <a name="accessibility-in-data-visualizations"></a><span data-ttu-id="c9147-207">Acessibilidade nas visualizações de dados</span><span class="sxs-lookup"><span data-stu-id="c9147-207">Accessibility in data visualizations</span></span>

- <span data-ttu-id="c9147-p123">Não use cor como a única maneira de comunicar informações. As pessoas que são daltônicas não serão capazes de interpretar os resultados. Use forma, tamanho e textura, além de cor quando possível para comunicar informações.</span><span class="sxs-lookup"><span data-stu-id="c9147-p123">Do not use color as the only way to communicate information. People who are color blind will not be able to interpret the results. Use shape, size and texture in addition to color when possible to communicate information.</span></span>
- <span data-ttu-id="c9147-211">Torne todos os elementos interativos, como botões de ação ou listas de escolha, acessíveis a partir de um teclado.</span><span class="sxs-lookup"><span data-stu-id="c9147-211">Make all interactive elements, such as push buttons or pick lists, accessible from a keyboard.</span></span>
- <span data-ttu-id="c9147-212">Envie eventos de acessibilidade para leitores de tela para anunciar alterações de foco, dicas de ferramentas e assim por diante.</span><span class="sxs-lookup"><span data-stu-id="c9147-212">Send accessibility events to screen readers to announce focus changes, tooltips, and so on.</span></span>

## <a name="see-also"></a><span data-ttu-id="c9147-213">Confira também</span><span class="sxs-lookup"><span data-stu-id="c9147-213">See also</span></span>

- [<span data-ttu-id="c9147-214">As cinco melhores bibliotecas para criar visualizações de dados</span><span class="sxs-lookup"><span data-stu-id="c9147-214">The Five Best Libraries for Building Data Visualizations</span></span>](https://www.fastcompany.com/3029760/the-five-best-libraries-for-building-data-vizualizations)
- [<span data-ttu-id="c9147-215">Exibição Visual de informações quantitativas</span><span class="sxs-lookup"><span data-stu-id="c9147-215">The Visual Display of Quantitative Information</span></span>](https://www.edwardtufte.com/tufte/books_vdqi)
