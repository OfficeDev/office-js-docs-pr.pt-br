---
title: Excel conjunto de requisitos da API JavaScript 1.7
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.7.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-17"></a>Quais são as novidades na API JavaScript do Excel 1.7

O conjunto de requisitos 1.7 da API JavaScript do Excel incluei APIs para gráficos, eventos, planilhas, intervalos, propriedades do documento, itens nomeados, opções de proteção e estilos.

## <a name="customize-charts"></a>Personalize gráficos

Com as novas APIs de gráficos, você pode criar tipos degráficos adicionais, adicionar uma série de dados a um gráfico, definir o título do gráfico, adicionar um título de eixo, adicionar unidade de exibição, adicionar uma linha de tendência com média móvel, alterar uma linha de tendência para linear e muito mais. A seguir estão alguns exemplos.

- Eixo gráfico - obtenha, defina, formate e remova unidade de eixo, etiqueta e título em um gráfico.
- Série de gráficos - adicione, defina e exclua uma série em um gráfico.  Alterar marcadores da série, pedidos de plotagem e dimensionamento.
- Gráfico de linhas de tendências: adicione, receba e formate linhas de tendências em um gráfico.
- Legenda do gráfico - formate a fonte de legenda de um gráfico.
- Ponto do gráfico - defina a cor do ponto do gráfico.
- Subtítulo do título do gráfico - obtenha e defina a subseqüência do título para um gráfico.
- Tipo de gráfico - opção para criar mais tipos de gráfico.

## <a name="events"></a>Eventos

As APIs de eventos JavaScript do Excel fornecem diversos,  manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Você pode criar essa função para executar as ações que seu cenário exige. Para obter uma lista de eventos que estão disponíveis, confira [trabalhar com eventos usando as API JavaScript do Excel](../../excel/excel-add-ins-events.md).

## <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Personalizar a aparência de planilhas e intervalos

Nas novas APIs você pode personalizar a aparência das planilhas de várias maneiras:

- Congele painéis para manter linhas ou colunas específicas visíveis durante a rolagem na planilha. Por exemplo, se a primeira linha da planilha inclui cabeçalhos, você pode congelá-la para que os cabeçalhos das colunas permaneçam visíveis enquanto rola para baixo na planilha.
- Modificar a cor da guia de planilha.
- Adicione títulos de planilha.

Você pode personalizar a aparência de intervalos de várias maneiras:

- Defina o estilo de célula para um intervalo para garantir que todas as células no intervalo tenham formatação consistente. Um estilo de célula é um conjunto definido de características de formatação, como fontes e tamanhos de fonte, formatos numéricos, bordas de célula e sombreamento de célula. Use qualquer um dos estilos de célula internas do Excel ou crie seu próprio estilo de célula personalizado.
- Defina a orientação de texto para um intervalo.
- Adicione ou modifique um hiperlink em um intervalo vinculado a outro local na pasta de trabalho ou a um local externo.

## <a name="manage-document-properties"></a>Gerenciar propriedades dos documentos

Usando as APIs de propriedades do documento, você pode acessar as propriedades do documento interno e também criar e gerenciar propriedades personalizadas do documento para armazenar o estado da pasta de trabalho e direcionar o fluxo de trabalho e a lógica comercial.

## <a name="copy-worksheets"></a>Copiar planilhas

Usando a cópia da planilha APIs, você pode copiar os dados e o formato de uma planilha para uma nova planilha na mesma pasta de trabalho e reduzir a quantidade de transferência de dados necessária.

## <a name="handle-ranges-with-ease"></a>Lidar com intervalos com facilidade

Usando várias APIs de intervalo, você pode fazer coisas como obter região ao redor, obter um intervalo redimensionado e muito mais. Essas APIs devem tornar as tarefas, como manipulação de intervalo e endereçamento, muito mais eficientes.

Além disso:

- Opções de proteção de pasta de trabalho e planilha - use estas APIs para proteger dados em uma planilha e a estrutura da pasta de trabalho.
- Atualizar um item nomeado - usar esta API para atualizar um item nomeado.
- Obter a célula ativa - usar esta API para acessar a célula ativa da pasta de trabalho.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.7. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.7 ou anterior, consulte Excel APIs no conjunto de requisitos [1.7](/javascript/api/excel?view=excel-js-1.7&preserve-view=true) ou anterior.

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#excel-excel-chart-charttype-member)|Especifica o tipo do gráfico.|
||[id](/javascript/api/excel/excel.chart#excel-excel-chart-id-member)|Id exclusiva do gráfico.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#excel-excel-chart-showallfieldbuttons-member)|Especifica se todos os botões de campo serão exibidos em um Gráfico Dinâmico.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[borda](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-border-member)|Representa o formato de borda da área do gráfico, que inclui cor, estilo de linha e peso.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem(type: Excel. ChartAxisType, group?: Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-getitem-member(1))|Retorna o eixo específico identificado por tipo e grupo.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[axisGroup](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-axisgroup-member)|Especifica o grupo do eixo especificado.|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-basetimeunit-member)|Especifica a unidade base do eixo de categoria especificado.|
||[categoryType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-categorytype-member)|Especifica o tipo de eixo de categoria.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-customdisplayunit-member)|Especifica o valor da unidade de exibição do eixo personalizado.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-displayunit-member)|Representa a unidade de exibição de eixo.|
||[height](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-height-member)|Especifica a altura, em pontos, do eixo do gráfico.|
||[left](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-left-member)|Especifica a distância, em pontos, da borda esquerda do eixo até a esquerda da área do gráfico.|
||[logBase](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-logbase-member)|Especifica a base do logaritmo ao usar escalas logarítmicas.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majortickmark-member)|Especifica o tipo de marca de escala principal para o eixo especificado.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majortimeunitscale-member)|Especifica o valor de escala de unidade principal para o eixo de categoria quando a `categoryType` propriedade é definida como `dateAxis`.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minortickmark-member)|Especifica o tipo de marca de escala secundária para o eixo especificado.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minortimeunitscale-member)|Especifica o valor de escala de unidade secundária para o eixo de categoria quando a `categoryType` propriedade é definida como `dateAxis`.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-reverseplotorder-member)|Especifica se Excel plota pontos de dados do último para o primeiro.|
||[scaleType](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-scaletype-member)|Especifica o tipo de escala do eixo do valor.|
||[setCategoryNames(sourceData: Range)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setcategorynames-member(1))|Define todos os nomes de categoria para o eixo especificado.|
||[setCustomDisplayUnit(value: number)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setcustomdisplayunit-member(1))|Definirá a unidade de exibição de eixo a um valor personalizado.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-showdisplayunitlabel-member)|Especifica se o rótulo da unidade de exibição do eixo está visível.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-ticklabelposition-member)|Especifica a posição dos rótulos de marcas de escala no eixo especificado.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-ticklabelspacing-member)|Especifica o número de categorias ou séries entre rótulos de marca de escala.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-tickmarkspacing-member)|Especifica o número de categorias ou séries entre marcas de escala.|
||[top](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-top-member)|Especifica a distância, em pontos, da borda superior do eixo até a parte superior da área do gráfico.|
||[tipo](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-type-member)|Especifica o tipo de eixo.|
||[visible](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-visible-member)|Especifica se o eixo está visível.|
||[width](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-width-member)|Especifica a largura, em pontos, do eixo do gráfico.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-color-member)|Código de cor HTML que representa a cor das bordas no gráfico.|
||[lineStyle](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-linestyle-member)|Representa o estilo de linha da borda.|
||[peso](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-weight-member)|Representa a espessura da borda, em pontos.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-position-member)|Valor que representa a posição do rótulo de dados.|
||[separador](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-separator-member)|Cadeia de caracteres que representa o separador usado para o rótulo de dados em um gráfico.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showbubblesize-member)|Especifica se o tamanho da bolha do rótulo de dados está visível.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showcategoryname-member)|Especifica se o nome da categoria do rótulo de dados está visível.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showlegendkey-member)|Especifica se a chave de legenda do rótulo de dados está visível.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showpercentage-member)|Especifica se a porcentagem do rótulo de dados está visível.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showseriesname-member)|Especifica se o nome da série de rótulos de dados está visível.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-showvalue-member)|Especifica se o valor do rótulo de dados está visível.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#excel-excel-chartformatstring-font-member)|Representa os atributos de fonte, como nome da fonte, tamanho da fonte e cor de um objeto de caracteres de gráfico.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-height-member)|Especifica a altura, em pontos, da legenda no gráfico.|
||[left](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-left-member)|Especifica o valor esquerdo, em pontos, da legenda no gráfico.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-legendentries-member)|Representa uma coleção de legendEntries na legenda.|
||[showShadow](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-showshadow-member)|Especifica se a legenda tem uma sombra no gráfico.|
||[top](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-top-member)|Especifica a parte superior de uma legenda de gráfico.|
||[width](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-width-member)|Especifica a largura, em pontos, da legenda no gráfico.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-visible-member)|Representa a visibilidade de uma entrada de legenda de gráfico.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getcount-member(1))|Retorna o número de entradas de legenda na coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-getitemat-member(1))|Retorna uma entrada de legenda no índice determinado.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#excel-excel-chartlegendentrycollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-linestyle-member)|Representa o estilo de linha.|
||[weight](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-weight-member)|Representa a espessura da linha, em pontos.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[dataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-datalabel-member)|Retorna o rótulo de dados de um ponto de gráfico.|
||[hasDataLabel](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-hasdatalabel-member)|Representa se um ponto de dados tem um rótulo de dados.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerbackgroundcolor-member)|Representação de código de cor HTML da cor de plano de fundo do marcador de um ponto de dados (por exemplo, #FF0000 representa Vermelho).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerforegroundcolor-member)|Representação de código de cor HTML da cor de primeiro plano do marcador de um ponto de dados (por exemplo, #FF0000 representa Vermelho).|
||[markerSize](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markersize-member)|Representa o tamanho do marcador de um ponto de dados.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-markerstyle-member)|Representa estilo do marcador de um ponto de dados do gráfico.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[borda](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-border-member)|Representa o formato de borda de um ponto de dados do gráfico, que inclui informações de cor, estilo e peso.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-charttype-member)|Representa o tipo de gráfico de uma série.|
||[delete()](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-delete-member(1))|Exclui a série de gráfico.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-doughnutholesize-member)|Representa o tamanho do furo de rosca de uma série de gráficos.|
||[filtrado](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-filtered-member)|Especifica se a série é filtrada.|
||[gapWidth](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gapwidth-member)|Representa a largura do espaçamento de uma série de gráfico.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-hasdatalabels-member)|Especifica se a série tem rótulos de dados.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerbackgroundcolor-member)|Especifica a cor de plano de fundo do marcador de uma série de gráficos.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerforegroundcolor-member)|Especifica a cor do marcador em primeiro plano de uma série de gráficos.|
||[markerSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markersize-member)|Especifica o tamanho do marcador de uma série de gráficos.|
||[markerStyle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-markerstyle-member)|Especifica o estilo de marcador de uma série de gráficos.|
||[plotOrder](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-plotorder-member)|Especifica a ordem de plotagem de uma série de gráficos dentro do grupo de gráficos.|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setbubblesizes-member(1))|Define os tamanhos de bolha para uma série de gráficos.|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setvalues-member(1))|Define os valores de uma série de gráficos.|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-setxaxisvalues-member(1))|Define os valores do eixo x para uma série de gráficos.|
||[showShadow](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showshadow-member)|Especifica se a série tem uma sombra.|
||[smooth](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-smooth-member)|Especifica se a série é suave.|
||[trendlines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-trendlines-member)|A coleção de linhas de tendência na série.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-add-member(1))|Adiciona uma nova série para o conjunto.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-getsubstring-member(1))|Obter a subdistragem de um título de gráfico.|
||[height](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-height-member)|Representa a altura, em pontos, do título do gráfico.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-horizontalalignment-member)|Especifica o alinhamento horizontal para o título do gráfico.|
||[left](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-left-member)|Especifica a distância, em pontos, da borda esquerda do título do gráfico até a borda esquerda da área do gráfico.|
||[position](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-position-member)|Representa a posição de título do gráfico.|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-setformula-member(1))|Define um valor de cadeia de caracteres que representa a fórmula do título do eixo do gráfico usando a notação no estilo A1.|
||[showShadow](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-showshadow-member)|Representa um valor booliano que determina se o título do gráfico tiver uma sombra.|
||[textOrientation](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-textorientation-member)|Especifica o ângulo para o qual o texto é orientado para o título do gráfico.|
||[top](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-top-member)|Especifica a distância, em pontos, da borda superior do título do gráfico até a parte superior da área do gráfico.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-verticalalignment-member)|Especifica o alinhamento vertical do título do gráfico.|
||[width](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-width-member)|Especifica a largura, em pontos, do título do gráfico.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[borda](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-border-member)|Representa o formato de borda do título do gráfico, que inclui cor, estilo de linha e peso.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-delete-member(1))|Deleta o objeto Trendline.|
||[format](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-format-member)|Representa a formatação de uma linha de tendência do gráfico.|
||[intercept](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-intercept-member)|Representa o valor de intercepção da linha de tendência.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-movingaverageperiod-member)|Representa o período de uma linha de tendência de gráfico.|
||[name](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-name-member)|Representa o nome da linha de tendência.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-polynomialorder-member)|Representa a ordem de uma linha de tendência de gráfico.|
||[tipo](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-type-member)|Representa o tipo da linha de tendência de um gráfico.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-add-member(1))|Adiciona uma nova linha de tendência ao conjunto de linha de tendência.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getcount-member(1))|Retorna o número de linha de tendência na coleção.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-getitem-member(1))|Obtém um objeto trendline por índice, que é a ordem de inserção na matriz de itens.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#excel-excel-charttrendlinecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#excel-excel-charttrendlineformat-line-member)|Representa a formatação de linha do gráfico.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-delete-member(1))|Exclui a propriedade personalizada.|
||[key](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-key-member)|A chave da propriedade personalizada.|
||[tipo](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-type-member)|O tipo do valor usado para a propriedade personalizada.|
||[value](/javascript/api/excel/excel.customproperty#excel-excel-customproperty-value-member)|O valor da propriedade personalizada.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-add-member(1))|Cria uma nova propriedade personalizada ou define uma existente.|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-deleteall-member(1))|Exclui todas as propriedades personalizadas nesta coleção.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getcount-member(1))|Obtém a contagem das propriedades personalizadas.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitem-member(1))|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-getitemornullobject-member(1))|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[items](/javascript/api/excel/excel.custompropertycollection#excel-excel-custompropertycollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#excel-excel-dataconnectioncollection-refreshall-member(1))|Atualiza todas as conexões de dados na coleção.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[author](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-author-member)|O autor da workbook.|
||[category](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-category-member)|A categoria da guia de trabalho.|
||[comments](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-comments-member)|Os comentários da workbook.|
||[company](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-company-member)|A empresa da workbook.|
||[creationDate](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-creationdate-member)|Obtém a data de criação da pasta de trabalho.|
||[custom](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-custom-member)|Obtém a coleção de propriedades personalizadas da pasta de trabalho.|
||[keywords](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-keywords-member)|As palavras-chave da workbook.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-lastauthor-member)|Obtém o último autor da pasta de trabalho.|
||[manager](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-manager-member)|O gerente da workbook.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-revisionnumber-member)|Obtém o número de revisão da pasta de trabalho.|
||[subject](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-subject-member)|O assunto da workbook.|
||[title](/javascript/api/excel/excel.documentproperties#excel-excel-documentproperties-title-member)|O título da guia de trabalho.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[arrayValues](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-arrayvalues-member)|Retorna um objeto que contém valores e tipos do item nomeado.|
||[formula](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-formula-member)|A fórmula do item nomeado.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-types-member)|Representa os tipos de cada item na matriz de itens nomeados|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#excel-excel-nameditemarrayvalues-values-member)|Representa os valores de cada item na matriz de itens nomeados.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange(numRows: number, numColumns: number)](/javascript/api/excel/excel.range#excel-excel-range-getabsoluteresizedrange-member(1))|Obtém `Range` um objeto com a mesma célula superior esquerda que `Range` o objeto atual, mas com os números especificados de linhas e colunas.|
||[getImage()](/javascript/api/excel/excel.range#excel-excel-range-getimage-member(1))|Renderiza o intervalo como uma imagem png codificada com base64.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#excel-excel-range-getsurroundingregion-member(1))|Retorna um `Range` objeto que representa a região ao redor da célula superior esquerda neste intervalo.|
||[hiperlink](/javascript/api/excel/excel.range#excel-excel-range-hyperlink-member)|Representa o hiperlink do intervalo atual.|
||[isEntireColumn](/javascript/api/excel/excel.range#excel-excel-range-isentirecolumn-member)|Representa se o intervalo atual está em uma coluna inteira.|
||[isEntireRow](/javascript/api/excel/excel.range#excel-excel-range-isentirerow-member)|Representa se o intervalo atual está em uma linha inteira.|
||[numberFormatLocal](/javascript/api/excel/excel.range#excel-excel-range-numberformatlocal-member)|Representa Excel código de formato de número do usuário para o intervalo determinado, com base nas configurações de idioma do usuário.|
||[showCard()](/javascript/api/excel/excel.range#excel-excel-range-showcard-member(1))|Exibe o cartão para uma célula ativa se ele tiver um conteúdo valioso.|
||[style](/javascript/api/excel/excel.range#excel-excel-range-style-member)|Representa o estilo de intervalo atual.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-textorientation-member)|A orientação de texto de todas as células dentro do intervalo.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-usestandardheight-member)|Determina se a altura da linha do `Range` objeto é igual à altura padrão da planilha.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-usestandardwidth-member)|Especifica se a largura da coluna do `Range` objeto é igual à largura padrão da planilha.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-address-member)|Representa o destino de URL do hiperlink.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-documentreference-member)|Representa o destino de referência do documento para o hiperlink.|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-screentip-member)|Representa a cadeia exibida ao passar o mouse sobre o hiperlink.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#excel-excel-rangehyperlink-texttodisplay-member)|Representa a cadeia de caracteres exibida na parte superior esquerda da maioria das células no intervalo.|
|[Style](/javascript/api/excel/excel.style)|[Borders](/javascript/api/excel/excel.style#excel-excel-style-borders-member)|Uma coleção de quatro objetos de borda que representam o estilo das quatro bordas.|
||[builtIn](/javascript/api/excel/excel.style#excel-excel-style-builtin-member)|Especifica se o estilo é um estilo integrado.|
||[delete()](/javascript/api/excel/excel.style#excel-excel-style-delete-member(1))|Exclui este estilo.|
||[fill](/javascript/api/excel/excel.style#excel-excel-style-fill-member)|O preenchimento do estilo.|
||[font](/javascript/api/excel/excel.style#excel-excel-style-font-member)|Um `Font` objeto que representa a fonte do estilo.|
||[formulaHidden](/javascript/api/excel/excel.style#excel-excel-style-formulahidden-member)|Especifica se a fórmula ficará oculta quando a planilha estiver protegida.|
||[horizontalAlignment](/javascript/api/excel/excel.style#excel-excel-style-horizontalalignment-member)|Representa o alinhamento horizontal para o estilo.|
||[includeAlignment](/javascript/api/excel/excel.style#excel-excel-style-includealignment-member)|Especifica se o estilo inclui o recuo automático, o alinhamento horizontal, o alinhamento vertical, o texto de quebra, o nível de recuo e as propriedades de orientação de texto.|
||[includeBorder](/javascript/api/excel/excel.style#excel-excel-style-includeborder-member)|Especifica se o estilo inclui as propriedades de cor, índice de cor, estilo de linha e borda de peso.|
||[includeFont](/javascript/api/excel/excel.style#excel-excel-style-includefont-member)|Especifica se o estilo inclui as propriedades de fonte de plano de fundo, negrito, cor, índice de cores, estilo de fonte, itálico, nome, tamanho, tachado, subscrito, sobrescrito e sublinhado.|
||[includeNumber](/javascript/api/excel/excel.style#excel-excel-style-includenumber-member)|Especifica se o estilo inclui a propriedade de formato de número.|
||[includePatterns](/javascript/api/excel/excel.style#excel-excel-style-includepatterns-member)|Especifica se o estilo inclui a cor, o índice de cores, inverte se negativo, padrão, cor do padrão e propriedades internas do índice de cores padrão.|
||[includeProtection](/javascript/api/excel/excel.style#excel-excel-style-includeprotection-member)|Especifica se o estilo inclui a fórmula oculta e as propriedades de proteção bloqueadas.|
||[indentLevel](/javascript/api/excel/excel.style#excel-excel-style-indentlevel-member)|Um número inteiro entre 0 e 250 que indica o nível de recuo do estilo.|
||[bloqueado](/javascript/api/excel/excel.style#excel-excel-style-locked-member)|Especifica se o objeto está bloqueado quando a planilha está protegida.|
||[name](/javascript/api/excel/excel.style#excel-excel-style-name-member)|O nome do estilo.|
||[numberFormat](/javascript/api/excel/excel.style#excel-excel-style-numberformat-member)|O código de formatação de formato de número para o estilo.|
||[numberFormatLocal](/javascript/api/excel/excel.style#excel-excel-style-numberformatlocal-member)|O código de formato localizado do formato numérico para o estilo.|
||[readingOrder](/javascript/api/excel/excel.style#excel-excel-style-readingorder-member)|A ordem de leitura para o estilo.|
||[shrinkToFit](/javascript/api/excel/excel.style#excel-excel-style-shrinktofit-member)|Especifica se o texto reduz automaticamente para caber na largura da coluna disponível.|
||[verticalAlignment](/javascript/api/excel/excel.style#excel-excel-style-verticalalignment-member)|Especifica o alinhamento vertical do estilo.|
||[wrapText](/javascript/api/excel/excel.style#excel-excel-style-wraptext-member)|Especifica se Excel quebra o texto no objeto.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-add-member(1))|Adiciona um novo estilo para o conjunto.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitem-member(1))|Obtém `Style` um pelo nome.|
||[items](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#excel-excel-table-onchanged-member)|Ocorre quando os dados nas células mudam em uma tabela específica.|
||[onSelectionChanged](/javascript/api/excel/excel.table#excel-excel-table-onselectionchanged-member)|Ocorre quando a seleção muda em uma tabela específica.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-address-member)|Obtém o endereço que representa a área alterada de uma tabela em uma planilha específica.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-changetype-member)|Obtém o tipo de alteração que representa como o evento alterado é disparado.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-source-member)|Obtém a origem do evento.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-tableid-member)|Obtém a ID da tabela na qual os dados foram alterados.|
||[tipo](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-worksheetid-member)|Obtém a ID da planilha na qual os dados foram alterados.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onchanged-member)|Ocorre quando os dados mudam em qualquer tabela em uma pasta de trabalho ou em uma planilha.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-address-member)|Obtém o endereço do intervalo que representa a área selecionada da tabela em uma planilha específica.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-isinsidetable-member)|Especifica se a seleção está dentro de uma tabela.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-tableid-member)|Obtém a ID da tabela na qual a seleção foi alterada.|
||[tipo](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#excel-excel-tableselectionchangedeventargs-worksheetid-member)|Obtém a ID da planilha na qual a seleção foi alterada.|
|[Workbook](/javascript/api/excel/excel.workbook)|[dataConnections](/javascript/api/excel/excel.workbook#excel-excel-workbook-dataconnections-member)|Representa todas as conexões de dados na workbook.|
||[getActiveCell()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivecell-member(1))|Obtém a célula ativa no momento da pasta de trabalho.|
||[name](/javascript/api/excel/excel.workbook#excel-excel-workbook-name-member)|Obtém o nome da pasta de trabalho.|
||[properties](/javascript/api/excel/excel.workbook#excel-excel-workbook-properties-member)|Obtém as propriedades da pasta de trabalho.|
||[protection](/javascript/api/excel/excel.workbook#excel-excel-workbook-protection-member)|Retorna o objeto de proteção de uma workbook.|
||[styles](/javascript/api/excel/excel.workbook#excel-excel-workbook-styles-member)|Representa uma coleção de estilos associados à pasta de trabalho.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protect-member(1))|Protege uma pasta de trabalho.|
||[protegido](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-protected-member)|Especifica se a workbook está protegida.|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#excel-excel-workbookprotection-unprotect-member(1))|Desprotege uma pasta de trabalho.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[copy(positionType?: Excel. WorksheetPositionType, relativeTo?: Excel. Planilha)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-copy-member(1))|Copia uma planilha e a coloca na posição especificada.|
||[freezePanes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-freezepanes-member)|Obtém um objeto que pode ser usado para manipular painéis congelados na planilha.|
||[getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrangebyindexes-member(1))|Obtém `Range` o objeto começando em um índice de linha específico e índice de coluna e abrangendo um determinado número de linhas e colunas.|
||[onActivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onactivated-member)|Ocorre quando a planilha é ativada.|
||[onChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onchanged-member)|Ocorre quando os dados mudam em uma planilha específica.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-ondeactivated-member)|Ocorre quando a planilha é desativada.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onselectionchanged-member)|Ocorre quando a seleção é mudada em uma planilha específica.|
||[standardHeight](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardheight-member)|Retorna a altura padrão de todas as linhas na planilha, em pontos.|
||[standardWidth](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-standardwidth-member)|Especifica a largura padrão (padrão) de todas as colunas na planilha.|
||[tabColor](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tabcolor-member)|A cor da guia da planilha.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[tipo](/javascript/api/excel/excel.worksheetactivatedeventargs#excel-excel-worksheetactivatedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#excel-excel-worksheetactivatedeventargs-worksheetid-member)|Obtém a ID da planilha ativada.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-source-member)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#excel-excel-worksheetaddedeventargs-worksheetid-member)|Obtém a ID da planilha adicionada à pasta de trabalho.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-address-member)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changetype-member)|Obtém o tipo de alteração que representa como o evento alterado é disparado.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-source-member)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-worksheetid-member)|Obtém a ID da planilha na qual os dados foram alterados.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onactivated-member)|Ocorre quando qualquer planilha na pasta de trabalho é ativada.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onadded-member)|Ocorre quando uma nova planilha é adicionada à pasta de trabalho.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeactivated-member)|Ocorre quando qualquer planilha na pasta de trabalho é desativada.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-ondeleted-member)|Ocorre quando uma planilha é excluída da pasta de trabalho.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[tipo](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#excel-excel-worksheetdeactivatedeventargs-worksheetid-member)|Obtém a ID da planilha desativada.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-source-member)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#excel-excel-worksheetdeletedeventargs-worksheetid-member)|Obtém a ID da planilha excluída da pasta de trabalho.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt(frozenRange: Range \| string)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezeat-member(1))|Define as células congeladas no modo de exibição da planilha ativa.|
||[freezeColumns(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezecolumns-member(1))|Congelar a primeira coluna ou colunas da planilha no local.|
||[freezeRows(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-freezerows-member(1))|Congelar a linha superior ou as linhas da planilha no local.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-getlocation-member(1))|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-getlocationornullobject-member(1))|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|
||[unfreeze()](/javascript/api/excel/excel.worksheetfreezepanes#excel-excel-worksheetfreezepanes-unfreeze-member(1))|Remove todos os painéis congelados na planilha.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#excel-excel-worksheetprotection-unprotect-member(1))|Desprotege uma planilha.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditobjects-member)|Representa a opção de proteção de planilha que permite a edição de objetos.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-alloweditscenarios-member)|Representa a opção de proteção de planilha que permite a edição de cenários.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#excel-excel-worksheetprotectionoptions-selectionmode-member)|Representa a opção de proteção da planilha do modo de seleção.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-address-member)|Obtém o endereço do intervalo que representa a área selecionada de uma planilha específica.|
||[tipo](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#excel-excel-worksheetselectionchangedeventargs-worksheetid-member)|Obtém a ID da planilha na qual a seleção foi alterada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
