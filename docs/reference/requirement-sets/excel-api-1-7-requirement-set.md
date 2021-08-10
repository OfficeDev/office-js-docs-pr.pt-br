---
title: Excel Conjunto de requisitos da API JavaScript 1.7
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.7.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1012efa0143710a6646c5288ef8c4adfd49a6006e62eb850f25d6d12e9028c51
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57088957"
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

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.7. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.7 ou anterior, consulte Excel APIs no conjunto de requisitos [1.7](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)ou anterior .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Gráfico](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#chartType)|Especifica o tipo do gráfico.|
||[id](/javascript/api/excel/excel.chart#id)|Id exclusiva do gráfico.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showAllFieldButtons)|Especifica se todos os botões de campo serão exibidos em um Gráfico Dinâmico.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[border](/javascript/api/excel/excel.chartareaformat#border)|Representa o formato de borda da área do gráfico, que inclui cor, estilo de linha e peso.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem(type: Excel. ChartAxisType, group?: Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getItem_type__group_)|Retorna o eixo específico identificado por tipo e grupo.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#baseTimeUnit)|Especifica a unidade base do eixo de categoria especificado.|
||[categoryType](/javascript/api/excel/excel.chartaxis#categoryType)|Especifica o tipo de eixo de categoria.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayUnit)|Representa a unidade de exibição de eixo.|
||[logBase](/javascript/api/excel/excel.chartaxis#logBase)|Especifica a base do logaritmo ao usar escalas logarítmicas.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majorTickMark)|Especifica o tipo de marca de escala principal para o eixo especificado.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majorTimeUnitScale)|Especifica o valor de escala de unidade principal para o eixo de categoria quando a `categoryType` propriedade é definida como `dateAxis` .|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minorTickMark)|Especifica o tipo de marca de escala secundária para o eixo especificado.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minorTimeUnitScale)|Especifica o valor de escala de unidade secundária para o eixo de categoria quando a `categoryType` propriedade é definida como `dateAxis` .|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisGroup)|Especifica o grupo do eixo especificado.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customDisplayUnit)|Especifica o valor da unidade de exibição do eixo personalizado.|
||[height](/javascript/api/excel/excel.chartaxis#height)|Especifica a altura, em pontos, do eixo do gráfico.|
||[left](/javascript/api/excel/excel.chartaxis#left)|Especifica a distância, em pontos, da borda esquerda do eixo até a esquerda da área do gráfico.|
||[top](/javascript/api/excel/excel.chartaxis#top)|Especifica a distância, em pontos, da borda superior do eixo até a parte superior da área do gráfico.|
||[type](/javascript/api/excel/excel.chartaxis#type)|Especifica o tipo de eixo.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Especifica a largura, em pontos, do eixo do gráfico.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reversePlotOrder)|Especifica se Excel plota pontos de dados do último para o primeiro.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaleType)|Especifica o tipo de escala do eixo do valor.|
||[setCategoryNames(sourceData: Range)](/javascript/api/excel/excel.chartaxis#setCategoryNames_sourceData_)|Define todos os nomes de categoria para o eixo especificado.|
||[setCustomDisplayUnit(value: number)](/javascript/api/excel/excel.chartaxis#setCustomDisplayUnit_value_)|Definirá a unidade de exibição de eixo a um valor personalizado.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showDisplayUnitLabel)|Especifica se o rótulo da unidade de exibição do eixo está visível.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#tickLabelPosition)|Especifica a posição dos rótulos de marcas de escala no eixo especificado.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#tickLabelSpacing)|Especifica o número de categorias ou séries entre rótulos de marca de escala.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickMarkSpacing)|Especifica o número de categorias ou séries entre marcas de escala.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Especifica se o eixo está visível.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|Código de cor HTML que representa a cor das bordas no gráfico.|
||[lineStyle](/javascript/api/excel/excel.chartborder#lineStyle)|Representa o estilo de linha da borda.|
||[peso](/javascript/api/excel/excel.chartborder#weight)|Representa a espessura da borda, em pontos.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|Valor que representa a posição do rótulo de dados.|
||[separador](/javascript/api/excel/excel.chartdatalabel#separator)|Cadeia de caracteres que representa o separador usado para o rótulo de dados em um gráfico.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showBubbleSize)|Especifica se o tamanho da bolha do rótulo de dados está visível.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showCategoryName)|Especifica se o nome da categoria do rótulo de dados está visível.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showLegendKey)|Especifica se a chave de legenda do rótulo de dados está visível.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showPercentage)|Especifica se a porcentagem do rótulo de dados está visível.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showSeriesName)|Especifica se o nome da série de rótulos de dados está visível.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showValue)|Especifica se o valor do rótulo de dados está visível.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|Representa os atributos de fonte, como nome da fonte, tamanho da fonte e cor de um objeto de caracteres de gráfico.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Especifica a altura, em pontos, da legenda no gráfico.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Especifica o valor esquerdo, em pontos, da legenda no gráfico.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendEntries)|Representa uma coleção de legendEntries na legenda.|
||[showShadow](/javascript/api/excel/excel.chartlegend#showShadow)|Especifica se a legenda tem uma sombra no gráfico.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Especifica a parte superior de uma legenda de gráfico.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Especifica a largura, em pontos, da legenda no gráfico.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Representa a visibilidade de uma entrada de legenda de gráfico.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getCount__)|Retorna o número de entradas de legenda na coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getItemAt_index_)|Retorna uma entrada de legenda no índice determinado.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#lineStyle)|Representa o estilo de linha.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Representa a espessura da linha, em pontos.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasDataLabel)|Representa se um ponto de dados tem um rótulo de dados.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerBackgroundColor)|Representação de código de cor HTML da cor de plano de fundo do marcador de um ponto de dados (por exemplo, #FF0000 representa Vermelho).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerForegroundColor)|Representação de código de cor HTML da cor de primeiro plano do marcador de um ponto de dados (por exemplo, #FF0000 representa Vermelho).|
||[markerSize](/javascript/api/excel/excel.chartpoint#markerSize)|Representa o tamanho do marcador de um ponto de dados.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerStyle)|Representa estilo do marcador de um ponto de dados do gráfico.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#dataLabel)|Retorna o rótulo de dados de um ponto de gráfico.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[border](/javascript/api/excel/excel.chartpointformat#border)|Representa o formato de borda de um ponto de dados do gráfico, que inclui informações de cor, estilo e peso.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#chartType)|Representa o tipo de gráfico de uma série.|
||[delete()](/javascript/api/excel/excel.chartseries#delete__)|Exclui a série de gráfico.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutHoleSize)|Representa o tamanho do furo de rosca de uma série de gráficos.|
||[filtrado](/javascript/api/excel/excel.chartseries#filtered)|Especifica se a série é filtrada.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapWidth)|Representa a largura do espaçamento de uma série de gráfico.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasDataLabels)|Especifica se a série tem rótulos de dados.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerBackgroundColor)|Especifica a cor de plano de fundo do marcador de uma série de gráficos.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerForegroundColor)|Especifica a cor do marcador em primeiro plano de uma série de gráficos.|
||[markerSize](/javascript/api/excel/excel.chartseries#markerSize)|Especifica o tamanho do marcador de uma série de gráficos.|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerStyle)|Especifica o estilo de marcador de uma série de gráficos.|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotOrder)|Especifica a ordem de plotagem de uma série de gráficos dentro do grupo de gráficos.|
||[trendlines](/javascript/api/excel/excel.chartseries#trendlines)|A coleção de linhas de tendência na série.|
||[setBubbleSizes(sourceData: Range)](/javascript/api/excel/excel.chartseries#setBubbleSizes_sourceData_)|Define os tamanhos de bolha para uma série de gráficos.|
||[setValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setValues_sourceData_)|Define os valores de uma série de gráficos.|
||[setXAxisValues(sourceData: Range)](/javascript/api/excel/excel.chartseries#setXAxisValues_sourceData_)|Define os valores do eixo x para uma série de gráficos.|
||[showShadow](/javascript/api/excel/excel.chartseries#showShadow)|Especifica se a série tem uma sombra.|
||[smooth](/javascript/api/excel/excel.chartseries#smooth)|Especifica se a série é suave.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[add(name?: string, index?: number)](/javascript/api/excel/excel.chartseriescollection#add_name__index_)|Adiciona uma nova série para o conjunto.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring(start: number, length: number)](/javascript/api/excel/excel.charttitle#getSubstring_start__length_)|Obter a subdistragem de um título de gráfico.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalAlignment)|Especifica o alinhamento horizontal para o título do gráfico.|
||[left](/javascript/api/excel/excel.charttitle#left)|Especifica a distância, em pontos, da borda esquerda do título do gráfico até a borda esquerda da área do gráfico.|
||[position](/javascript/api/excel/excel.charttitle#position)|Representa a posição de título do gráfico.|
||[height](/javascript/api/excel/excel.charttitle#height)|Representa a altura, em pontos, do título do gráfico.|
||[width](/javascript/api/excel/excel.charttitle#width)|Especifica a largura, em pontos, do título do gráfico.|
||[setFormula(formula: string)](/javascript/api/excel/excel.charttitle#setFormula_formula_)|Define um valor de cadeia de caracteres que representa a fórmula do título do eixo do gráfico usando a notação no estilo A1.|
||[showShadow](/javascript/api/excel/excel.charttitle#showShadow)|Representa um valor booliano que determina se o título do gráfico tiver uma sombra.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textOrientation)|Especifica o ângulo para o qual o texto é orientado para o título do gráfico.|
||[top](/javascript/api/excel/excel.charttitle#top)|Especifica a distância, em pontos, da borda superior do título do gráfico até a parte superior da área do gráfico.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalAlignment)|Especifica o alinhamento vertical do título do gráfico.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[border](/javascript/api/excel/excel.charttitleformat#border)|Representa o formato de borda do título do gráfico, que inclui cor, estilo de linha e peso.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete__)|Deleta o objeto Trendline.|
||[intercept](/javascript/api/excel/excel.charttrendline#intercept)|Representa o valor de intercepção da linha de tendência.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingAveragePeriod)|Representa o período de uma linha de tendência de gráfico.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Representa o nome da linha de tendência.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialOrder)|Representa a ordem de uma linha de tendência de gráfico.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Representa a formatação de uma linha de tendência do gráfico.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Representa o tipo da linha de tendência de um gráfico.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[add(type?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add_type_)|Adiciona uma nova linha de tendência ao conjunto de linha de tendência.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getCount__)|Retorna o número de linha de tendência na coleção.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getItem_index_)|Obtém um objeto trendline por índice, que é a ordem de inserção na matriz de itens.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Representa a formatação de linha do gráfico.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete__)|Exclui a propriedade personalizada.|
||[key](/javascript/api/excel/excel.customproperty#key)|A chave da propriedade personalizada.|
||[type](/javascript/api/excel/excel.customproperty#type)|O tipo do valor usado para a propriedade personalizada.|
||[value](/javascript/api/excel/excel.customproperty#value)|O valor da propriedade personalizada.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[add(key: string, value: any)](/javascript/api/excel/excel.custompropertycollection#add_key__value_)|Cria uma nova propriedade personalizada ou define uma existente.|
||[deleteAll()](/javascript/api/excel/excel.custompropertycollection#deleteAll__)|Exclui todas as propriedades personalizadas nesta coleção.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getCount__)|Obtém a contagem das propriedades personalizadas.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getItem_key_)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getItemOrNullObject_key_)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll()](/javascript/api/excel/excel.dataconnectioncollection#refreshAll__)|Atualiza todas as conexões de dados na coleção.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[author](/javascript/api/excel/excel.documentproperties#author)|O autor da workbook.|
||[category](/javascript/api/excel/excel.documentproperties#category)|A categoria da guia de trabalho.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Os comentários da workbook.|
||[company](/javascript/api/excel/excel.documentproperties#company)|A empresa da workbook.|
||[keywords](/javascript/api/excel/excel.documentproperties#keywords)|As palavras-chave da workbook.|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|O gerente da workbook.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationDate)|Obtém a data de criação da pasta de trabalho.|
||[custom](/javascript/api/excel/excel.documentproperties#custom)|Obtém a coleção de propriedades personalizadas da pasta de trabalho.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastAuthor)|Obtém o último autor da pasta de trabalho.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionNumber)|Obtém o número de revisão da pasta de trabalho.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|O assunto da workbook.|
||[title](/javascript/api/excel/excel.documentproperties#title)|O título da guia de trabalho.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|A fórmula do item nomeado.|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayValues)|Retorna um objeto que contém valores e tipos do item nomeado.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Representa os tipos de cada item na matriz de itens nomeados|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Representa os valores de cada item na matriz de itens nomeados.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange(numRows: number, numColumns: number)](/javascript/api/excel/excel.range#getAbsoluteResizedRange_numRows__numColumns_)|Obtém um objeto com a mesma célula superior esquerda que o objeto atual, mas com os números especificados de `Range` `Range` linhas e colunas.|
||[getImage()](/javascript/api/excel/excel.range#getImage__)|Renderiza o intervalo como uma imagem png codificada com base64.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getSurroundingRegion__)|Retorna um `Range` objeto que representa a região ao redor da célula superior esquerda neste intervalo.|
||[hiperlink](/javascript/api/excel/excel.range#hyperlink)|Representa o hiperlink do intervalo atual.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberFormatLocal)|Representa Excel código de formato de número do usuário para o intervalo determinado, com base nas configurações de idioma do usuário.|
||[isEntireColumn](/javascript/api/excel/excel.range#isEntireColumn)|Representa se o intervalo atual está em uma coluna inteira.|
||[isEntireRow](/javascript/api/excel/excel.range#isEntireRow)|Representa se o intervalo atual está em uma linha inteira.|
||[showCard()](/javascript/api/excel/excel.range#showCard__)|Exibe o cartão para uma célula ativa se ele tiver um conteúdo valioso.|
||[style](/javascript/api/excel/excel.range#style)|Representa o estilo de intervalo atual.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textOrientation)|A orientação de texto de todas as células dentro do intervalo.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#useStandardHeight)|Determina se a altura da linha do objeto é igual à altura `Range` padrão da planilha.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#useStandardWidth)|Especifica se a largura da coluna do `Range` objeto é igual à largura padrão da planilha.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|Representa o destino de URL do hiperlink.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentReference)|Representa o destino de referência do documento para o hiperlink.|
||[screenTip](/javascript/api/excel/excel.rangehyperlink#screenTip)|Representa a cadeia exibida ao passar o mouse sobre o hiperlink.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#textToDisplay)|Representa a cadeia de caracteres exibida na parte superior esquerda da maioria das células no intervalo.|
|[Estilo](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete__)|Exclui este estilo.|
||[formulaHidden](/javascript/api/excel/excel.style#formulaHidden)|Especifica se a fórmula ficará oculta quando a planilha estiver protegida.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalAlignment)|Representa o alinhamento horizontal para o estilo.|
||[includeAlignment](/javascript/api/excel/excel.style#includeAlignment)|Especifica se o estilo inclui o recuo automático, o alinhamento horizontal, o alinhamento vertical, o texto de quebra, o nível de recuo e as propriedades de orientação de texto.|
||[includeBorder](/javascript/api/excel/excel.style#includeBorder)|Especifica se o estilo inclui as propriedades de cor, índice de cor, estilo de linha e borda de peso.|
||[includeFont](/javascript/api/excel/excel.style#includeFont)|Especifica se o estilo inclui as propriedades de fonte de plano de fundo, negrito, cor, índice de cores, estilo de fonte, itálico, nome, tamanho, tachado, subscrito, sobrescrito e sublinhado.|
||[includeNumber](/javascript/api/excel/excel.style#includeNumber)|Especifica se o estilo inclui a propriedade de formato de número.|
||[includePatterns](/javascript/api/excel/excel.style#includePatterns)|Especifica se o estilo inclui a cor, o índice de cores, inverte se negativo, padrão, cor do padrão e propriedades internas do índice de cores padrão.|
||[includeProtection](/javascript/api/excel/excel.style#includeProtection)|Especifica se o estilo inclui a fórmula oculta e as propriedades de proteção bloqueadas.|
||[indentLevel](/javascript/api/excel/excel.style#indentLevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo do estilo.|
||[bloqueado](/javascript/api/excel/excel.style#locked)|Especifica se o objeto está bloqueado quando a planilha está protegida.|
||[numberFormat](/javascript/api/excel/excel.style#numberFormat)|O código de formatação de formato de número para o estilo.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberFormatLocal)|O código de formato localizado do formato numérico para o estilo.|
||[readingOrder](/javascript/api/excel/excel.style#readingOrder)|A ordem de leitura para o estilo.|
||[Borders](/javascript/api/excel/excel.style#borders)|Uma coleção de quatro objetos de borda que representam o estilo das quatro bordas.|
||[builtIn](/javascript/api/excel/excel.style#builtIn)|Especifica se o estilo é um estilo integrado.|
||[fill](/javascript/api/excel/excel.style#fill)|O preenchimento do estilo.|
||[font](/javascript/api/excel/excel.style#font)|Um `Font` objeto que representa a fonte do estilo.|
||[name](/javascript/api/excel/excel.style#name)|O nome do estilo.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinkToFit)|Especifica se o texto reduz automaticamente para caber na largura da coluna disponível.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalAlignment)|Especifica o alinhamento vertical do estilo.|
||[wrapText](/javascript/api/excel/excel.style#wrapText)|Especifica se Excel quebra o texto no objeto.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add_name_)|Adiciona um novo estilo para o conjunto.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getItem_name_)|Obtém `Style` um pelo nome.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onChanged)|Ocorre quando os dados nas células mudam em uma tabela específica.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onSelectionChanged)|Ocorre quando a seleção muda em uma tabela específica.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|Obtém o endereço que representa a área alterada de uma tabela em uma planilha específica.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changeType)|Obtém o tipo de alteração que representa como o evento alterado é disparado.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Obtém a origem do evento.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableId)|Obtém a ID da tabela na qual os dados foram alterados.|
||[tipo](/javascript/api/excel/excel.tablechangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetId)|Obtém a ID da planilha na qual os dados foram alterados.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onChanged)|Ocorre quando os dados mudam em qualquer tabela em uma pasta de trabalho ou em uma planilha.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Obtém o endereço do intervalo que representa a área selecionada da tabela em uma planilha específica.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isInsideTable)|Especifica se a seleção está dentro de uma tabela.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableId)|Obtém a ID da tabela na qual a seleção foi alterada.|
||[tipo](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetId)|Obtém a ID da planilha na qual a seleção foi alterada.|
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getActiveCell__)|Obtém a célula ativa no momento da pasta de trabalho.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataConnections)|Representa todas as conexões de dados na workbook.|
||[name](/javascript/api/excel/excel.workbook#name)|Obtém o nome da pasta de trabalho.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Obtém as propriedades da pasta de trabalho.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Retorna o objeto de proteção de uma workbook.|
||[styles](/javascript/api/excel/excel.workbook#styles)|Representa uma coleção de estilos associados à pasta de trabalho.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[protect(password?: string)](/javascript/api/excel/excel.workbookprotection#protect_password_)|Protege uma pasta de trabalho.|
||[protegido](/javascript/api/excel/excel.workbookprotection#protected)|Especifica se a workbook está protegida.|
||[unprotect(password?: string)](/javascript/api/excel/excel.workbookprotection#unprotect_password_)|Desprotege uma pasta de trabalho.|
|[Planilha](/javascript/api/excel/excel.worksheet)|[copy(positionType?: Excel. WorksheetPositionType, relativeTo?: Excel. Planilha)](/javascript/api/excel/excel.worksheet#copy_positionType__relativeTo_)|Copia uma planilha e a coloca na posição especificada.|
||[getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](/javascript/api/excel/excel.worksheet#getRangeByIndexes_startRow__startColumn__rowCount__columnCount_)|Obtém o objeto começando em um índice de linha específico e índice de coluna e abrangendo um determinado número de `Range` linhas e colunas.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezePanes)|Obtém um objeto que pode ser usado para manipular painéis congelados na planilha.|
||[onActivated](/javascript/api/excel/excel.worksheet#onActivated)|Ocorre quando a planilha é ativada.|
||[onChanged](/javascript/api/excel/excel.worksheet#onChanged)|Ocorre quando os dados mudam em uma planilha específica.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#onDeactivated)|Ocorre quando a planilha é desativada.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onSelectionChanged)|Ocorre quando a seleção é mudada em uma planilha específica.|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardHeight)|Retorna a altura padrão de todas as linhas na planilha, em pontos.|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardWidth)|Especifica a largura padrão (padrão) de todas as colunas na planilha.|
||[tabColor](/javascript/api/excel/excel.worksheet#tabColor)|A cor da guia da planilha.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[tipo](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetId)|Obtém a ID da planilha ativada.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetaddedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetId)|Obtém a ID da planilha adicionada à pasta de trabalho.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changeType)|Obtém o tipo de alteração que representa como o evento alterado é disparado.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetId)|Obtém a ID da planilha na qual os dados foram alterados.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onActivated)|Ocorre quando qualquer planilha na pasta de trabalho é ativada.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onAdded)|Ocorre quando uma nova planilha é adicionada à pasta de trabalho.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#onDeactivated)|Ocorre quando qualquer planilha na pasta de trabalho é desativada.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#onDeleted)|Ocorre quando uma planilha é excluída da pasta de trabalho.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[tipo](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetId)|Obtém a ID da planilha desativada.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetId)|Obtém a ID da planilha excluída da pasta de trabalho.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt(frozenRange: Range \| string)](/javascript/api/excel/excel.worksheetfreezepanes#freezeAt_frozenRange_)|Define as células congeladas no modo de exibição da planilha ativa.|
||[freezeColumns(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezeColumns_count_)|Congelar a primeira coluna ou colunas da planilha no local.|
||[freezeRows(count?: number)](/javascript/api/excel/excel.worksheetfreezepanes#freezeRows_count_)|Congelar a linha superior ou as linhas da planilha no local.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getLocation__)|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getLocationOrNullObject__)|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|
||[unfreeze()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze__)|Remove todos os painéis congelados na planilha.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[unprotect(password?: string)](/javascript/api/excel/excel.worksheetprotection#unprotect_password_)|Desprotege uma planilha.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#allowEditObjects)|Representa a opção de proteção de planilha que permite a edição de objetos.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#allowEditScenarios)|Representa a opção de proteção de planilha que permite a edição de cenários.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionMode)|Representa a opção de proteção da planilha do modo de seleção.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Obtém o endereço do intervalo que representa a área selecionada de uma planilha específica.|
||[tipo](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetId)|Obtém a ID da planilha na qual a seleção foi alterada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
