---
title: Diretrizes de estilo de visualização de dados para Suplementos do Office
description: ''
ms.date: 12/04/2017
---



# <a name="data-visualization-style-guidelines-for-office-add-ins"></a>Diretrizes de estilo de visualização de dados para Suplementos do Office

Boas visualizações de dados ajudam os usuários a encontrarem informações em seus dados. Eles podem usar essas informações para contar histórias que informam e convencem. Este artigo fornece diretrizes para ajudá-lo a criar visualizações de dados eficazes em seus suplementos para o Excel e outros aplicativos do Office.

Recomendamos que você use o [Office UI Fabric](http://dev.office.com/fabric) para criar o cromado para suas visualizações de dados. O Office UI Fabric inclui estilos e componentes que se integram perfeitamente com a aparência do Office. 

<!--The following figure shows a data visualization in an add-in that uses Fabric.

![Image of a data visualization with Fabric elements applied**](../images/fabric-data-visualization.png) 

-->

## <a name="data-visualization-elements"></a>Elementos de visualização de dados

As visualizações de dados compartilham uma estrutura geral e elementos comuns visuais e interativos, incluindo títulos, rótulos e plotagens de dados, conforme mostrado nas figuras a seguir.

![Imagem de um gráfico de linhas com título, eixos, legenda e uma área de plotagem rotulada](../images/data-visualization-line-chart.png)
![Imagem de um gráfico de coluna com eixos, linhas de grade, legenda e plotagem de dados rotulada](../images/data-visualization-column-chart.png)

### <a name="chart-titles"></a>Títulos de gráfico

Siga estas diretrizes para títulos de gráfico:

- Deixe seus títulos de gráfico bem legíveis. Posicione-os para criar uma hierarquia visual em relação ao restante do gráfico.
- Em geral, use maiúsculas nas frases (a primeira letra da primeira palavra em letra maiúscula). Para criar o contraste ou reforçar hierarquias, você poderá usar todas em maiúsculas, mas use isso com moderação.
- Incorpore a [typeramp do Office UI Fabric](http://dev.office.com/fabric#/styles/typography) para deixar seus gráficos consistentes com a interface de usuário do Office, que usa o Segoe. Você também pode usar outra fonte para diferenciar o conteúdo do gráfico da interface do usuário.
- Use tipos sem serifa com contadores grandes.

Os exemplos a seguir mostram tipos com serifa e sem serifa usados em títulos de gráfico. Observe como o contraste de escala e o uso eficaz do espaço em branco criam uma hierarquia visual forte.

![Imagem de uma visualização de dados com fontes com serifa](../images/data-visualization-serif.png)
![Imagem de uma visualização de dados com fontes sem serifa](../images/data-visualization-sans-serif.png)

### <a name="axis-labels"></a>Rótulos dos eixos

Deixe os rótulos dos eixos escuros para serem lidos claramente, com um bom contraste entre as cores do plano de fundo e do texto. Verifique se não estão tão escuros que competem com a tinta dos dados.

Cinza claro é mais eficaz para rótulos dos eixos. Se você estiver usando o Fabric, consulte a [Paleta de cores neutras](http://dev.office.com/fabric#/styles/colors).

### <a name="data-ink"></a>Tinta dos dados

Os pixels que representam os dados reais em um gráfico são chamados de tinta dos dados. Esse deve ser o foco central da visualização. Evite o uso de sombras, contornos pesados ou elementos de design desnecessários que distorcem ou competem com os dados. Use gradientes apenas quando os valores dos dados estiverem vinculados aos valores das cores. Evite gráficos tridimensionais, a menos que um valor mensurável e objetivo seja associado a uma terceira dimensão.

### <a name="color"></a>Cor

Escolha cores que acompanham os temas do sistema operacional ou aplicativo em vez de cores codificadas. Ao mesmo tempo, não deixe que as cores que você aplica distorçam os dados. O uso incorreto de cores nas visualizações de dados pode resultar em distorção de dados e leitura incorreta de informações.

Para ver as práticas recomendadas para o uso de cores nas visualizações de dados, consulte o seguinte:


- [Por que as cores do arco-íris não são a melhor opção para as visualizações de dados](http://www.poynter.org/2013/why-rainbow-colors-arent-always-the-best-options-for-data-visualizations/224413/)
- [Color Brewer 2.0: Aviso de Cor para Cartografia](http://colorbrewer2.org/)
- [Eu Quero Matiz](http://tools.medialab.sciences-po.fr/iwanthue/)

### <a name="gridlines"></a>Linhas de grade

As linhas de grade geralmente são necessárias para a leitura precisa de um gráfico, mas elas devem ser apresentadas como um elemento visual secundário, aprimorando a tinta dos dados e não competindo com ela. Use linhas de grade estáticas finas e leves, a menos que elas tenham sido projetadas especificamente para alto contraste. Você também pode usar interação para criar linhas de grade dinâmicas, que aparecem no contexto quando um usuário interage com um gráfico.

Cinza claro é mais eficaz para linhas de grade. Se você estiver usando o Fabric, consulte a [Paleta de cores neutras](http://dev.office.com/fabric#/styles/colors).

A imagem a seguir mostra uma visualização de dados com linhas de grade.

![Imagem de uma visualização de dados com linhas de grade](../images/data-visualization-gridlines.png)

### <a name="legends"></a>Legendas

Adicione legendas, se for necessário:

- Diferenciar as séries
- Apresentar mudanças de escala ou valor

Confira se suas legendas aprimoram a tinta dos dados e não competem com ela. Coloque as legendas:


- Recue à esquerda acima da área de plotagem por padrão, se todos os itens de legenda se ajustarem acima do gráfico.
- No lado superior direito da área de plotagem, se todos os itens de legenda não couberem acima do gráfico e use uma barra de rolagem, se necessário.

Para otimizar a legibilidade e a acessibilidade, mapeie os marcadores de legenda para a forma de gráfico relevante. Por exemplo, use marcadores de legenda em círculo para gráfico de bolhas e de plotagem de dispersão. Use marcadores de legenda de segmento de linha para gráficos de linhas.

### <a name="data-labels-and-tooltips"></a>Dicas de ferramenta e rótulos de dados

Verifique se as dicas de ferramentas os e rótulos de dados têm espaço suficiente em branco e variação de tipos. Use algoritmos para minimizar oclusão e conflito. Por exemplo, uma dica de ferramenta pode ser exibida à direita de um ponto de dados por padrão, mas ser exibida à esquerda se forem detectadas bordas à direita.

## <a name="design-principles"></a>Princípios de design

A equipe de Design do Office criou o conjunto de princípios de design a seguir, que usamos ao criar novas visualizações de dados para o pacote de produtos do Office.

### <a name="visual-design-principles"></a>Princípios de design visual

- As visualizações devem honrar e aprimorar os dados, facilitando a compreensão. Realce os dados, adicionando elementos de suporte somente conforme o necessário para fornecer o contexto. Evite embelezamentos desnecessários (sombras, estruturas de tópicos etc), gráficos desnecessários ou distorção de dados.
- As visualizações devem incentivar a exploração fornecendo comentários visuais interessantes. Use padrões de interação bem estabelecidos, controles de interface e feedback claro do sistema.
- Incorpore princípios de design consagrados. Use princípios tipográficos e de design de comunicação visual estabelecidos para aprimorar a forma, a legibilidade e o significado.

### <a name="interaction-design-principles"></a>Princípios de design de interação

- Design para permitir a exploração.
- Permitir interações diretas com objetos que revelam novas informações (classificação ao arrastar, por exemplo).
- Use modelos de interação simples, diretos e familiares.

Para obter mais informações sobre como criar visualizações de dados interativas e amigáveis, confira [Princípios e armadilhas de interface do usuário](http://uitraps.com/).

### <a name="motion-design-principles"></a>Princípios de design de animação

A animação segue o estímulo. Os elementos visuais devem se mover na mesma direção e com a mesma velocidade. Isso se aplica a:

- Criação do gráfico
- Transição de um tipo de gráfico para outro
- Filtragem
- Classificação
- Adição ou subtração de dados
- Revisão ou segmentação de dados
- Redimensionamento de um gráfico

Crie uma percepção de causalidade. Quando preparar animações:

- Prepare uma coisa de cada vez. 
- Prepare as mudanças nos eixos antes da mudança na tinta dos dados.
- Prepare e anime objetos como um grupo se eles estiverem se movendo na mesma velocidade e na mesma direção.
- Prepare elementos de dados em grupos de não mais do que 4 a 5 objetos. Os espectadores têm dificuldade de acompanhar mais de 4 e 5 objetos independentemente.

A animação adiciona significado.

- Animações aumentam a compreensão do usuário das alterações nos dados, fornecem contexto e atuam como uma camada de anotação não verbal.
- A animação deve ocorrer em um espaço de coordenadas significativo da visualização.
- Personalize a animação para o visual. 
- Evite animações gratuitas.

A animação segue os dados.

- Preserve o mapeamentos de dados. Se uma área estiver vinculada a uma medida, mantenha essa área na transição.
- Manter uma linguagem de design de animação consistente. Sempre que possível, mapeie a animação da visualização de dados para a linguagem de design de animação do Office. Use animações semelhantes para tipos de gráfico semelhantes.

## <a name="accessibility-in-data-visualizations"></a>Acessibilidade nas visualizações de dados

- Não use cor como a única maneira de comunicar informações. As pessoas que são daltônicas não serão capazes de interpretar os resultados. Use forma, tamanho e textura, além de cor quando possível para comunicar informações.
- Torne todos os elementos interativos, como botões de ação ou listas de escolha, acessíveis a partir de um teclado.
- Envie eventos de acessibilidade para leitores de tela para anunciar alterações de foco, dicas de ferramentas e assim por diante.

## <a name="see-also"></a>Veja também 

- [Dados + Design: uma introdução simples para preparar e  visualizar as informações](https://infoactive.co/data-design)
- [As cinco melhores bibliotecas para criar visualizações de dados](http://www.fastcompany.com/3029760/the-five-best-libraries-for-building-data-vizualizations)
- [Exibição Visual de informações quantitativas](https://www.edwardtufte.com/tufte/books_vdqi)
