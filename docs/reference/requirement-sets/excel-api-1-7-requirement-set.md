---
title: Conjunto de requisitos de API JavaScript do Excel 1,7
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1,7.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ea1fe7a3d28acce2d1f4e9ff33f7b2bd31758fbd
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996232"
---
# <a name="whats-new-in-excel-javascript-api-17"></a>Quais são as novidades na API JavaScript do Excel 1.7

O conjunto de requisitos 1.7 da API JavaScript do Excel incluei APIs para gráficos, eventos, planilhas, intervalos, propriedades do documento, itens nomeados, opções de proteção e estilos.

## <a name="customize-charts"></a>Personalize gráficos

Com as novas APIs de gráficos, você pode criar tipos degráficos adicionais, adicionar uma série de dados a um gráfico, definir o título do gráfico, adicionar um título de eixo, adicionar unidade de exibição, adicionar uma linha de tendência com média móvel, alterar uma linha de tendência para linear e muito mais. Estes são alguns exemplos:

* Eixo gráfico - obtenha, defina, formate e remova unidade de eixo, etiqueta e título em um gráfico.
* Série de gráficos - adicione, defina e exclua uma série em um gráfico.  Alterar marcadores da série, pedidos de plotagem e dimensionamento.
* Gráfico de linhas de tendências: adicione, receba e formate linhas de tendências em um gráfico.
* Legenda do gráfico - formate a fonte de legenda de um gráfico.
* Ponto do gráfico - defina a cor do ponto do gráfico.
* Subtítulo do título do gráfico - obtenha e defina a subseqüência do título para um gráfico.
* Tipo de gráfico - opção para criar mais tipos de gráfico.

## <a name="events"></a>Eventos

As APIs de eventos JavaScript do Excel fornecem diversos,  manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Você pode criar essa função para executar as ações que seu cenário exige. Para obter uma lista de eventos que estão disponíveis, confira [trabalhar com eventos usando as API JavaScript do Excel](../../excel/excel-add-ins-events.md).

## <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Personalizar a aparência de planilhas e intervalos

Nas novas APIs você pode personalizar a aparência das planilhas de várias maneiras:

* Congele painéis para manter linhas ou colunas específicas visíveis durante a rolagem na planilha. Por exemplo, se a primeira linha da planilha inclui cabeçalhos, você pode congelá-la para que os cabeçalhos das colunas permaneçam visíveis enquanto rola para baixo na planilha.
* Modificar a cor da guia de planilha.
* Adicione títulos de planilha.

Você pode personalizar a aparência de intervalos de várias maneiras:

* Defina o estilo de célula para um intervalo para garantir que todas as células no intervalo tenham formatação consistente. Um estilo de célula é um conjunto definido de características de formatação, como fontes e tamanhos de fonte, formatos numéricos, bordas de célula e sombreamento de célula. Use qualquer um dos estilos de célula internas do Excel ou crie seu próprio estilo de célula personalizado.
* Defina a orientação de texto para um intervalo.
* Adicione ou modifique um hiperlink em um intervalo vinculado a outro local na pasta de trabalho ou a um local externo.

## <a name="manage-document-properties"></a>Gerenciar propriedades dos documentos

Usando as APIs de propriedades do documento, você pode acessar as propriedades do documento interno e também criar e gerenciar propriedades personalizadas do documento para armazenar o estado da pasta de trabalho e direcionar o fluxo de trabalho e a lógica comercial.

## <a name="copy-worksheets"></a>Copiar planilhas

Usando a cópia da planilha APIs, você pode copiar os dados e o formato de uma planilha para uma nova planilha na mesma pasta de trabalho e reduzir a quantidade de transferência de dados necessária.

## <a name="handle-ranges-with-ease"></a>Lidar com intervalos com facilidade

Usando várias APIs de intervalo, você pode fazer coisas como obter região ao redor, obter um intervalo redimensionado e muito mais. Essas APIs devem tornar as tarefas, como manipulação de intervalo e endereçamento, muito mais eficientes.

Além disso:

* Opções de proteção de pasta de trabalho e planilha - use estas APIs para proteger dados em uma planilha e a estrutura da pasta de trabalho.
* Atualizar um item nomeado - usar esta API para atualizar um item nomeado.
* Obter a célula ativa - usar esta API para acessar a célula ativa da pasta de trabalho.

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Excel 1,7. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Excel 1,7 ou anterior, confira [APIs do Excel no conjunto de requisitos 1,7 ou anterior](/javascript/api/excel?view=excel-js-1.7&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|Especifica o tipo do gráfico.|
||[id](/javascript/api/excel/excel.chart#id)|Id exclusiva do gráfico.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|Especifica se todos os botões de campo devem ser exibidos em um gráfico dinâmico.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[borda](/javascript/api/excel/excel.chartareaformat#border)|Representa o formato da borda da área do gráfico, que inclui cores, LineStyle e Weight.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (tipo: Excel. ChartAxisType, Group?: Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Retorna o eixo específico identificado por tipo e grupo.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|Especifica a unidade base do eixo de categoria especificado.|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|Especifica o tipo de eixo das categorias.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|Representa a unidade de exibição de eixo.|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|Especifica a base do logaritmo ao usar escalas logarítmicas.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|Especifica o tipo de marca de escala principal para o eixo especificado.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|Especifica o valor de escala de unidades principal para o eixo de categoria quando a Propriedade CategoryType estiver definida como escala de valores.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|Especifica o tipo de marca de escala secundária do eixo especificado.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|Especifica o valor de escala de unidades secundária para o eixo de categoria quando a Propriedade CategoryType estiver definida como escala de valores.|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|Especifica o grupo do eixo especificado.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|Especifica o valor da unidade de exibição do eixo personalizado.|
||[height](/javascript/api/excel/excel.chartaxis#height)|Especifica a altura, em pontos, do eixo do gráfico.|
||[left](/javascript/api/excel/excel.chartaxis#left)|Especifica a distância, em pontos, da borda esquerda do eixo à esquerda da área do gráfico.|
||[top](/javascript/api/excel/excel.chartaxis#top)|Especifica a distância, em pontos, da borda superior do eixo até a parte superior da área do gráfico.|
||[type](/javascript/api/excel/excel.chartaxis#type)|Especifica o tipo de eixo.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Especifica a largura, em pontos, do eixo do gráfico.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|Especifica se o Excel plota os pontos de dados do último ao primeiro.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|Especifica o tipo de escala do eixo dos valores.|
||[setcategorynames (sourceData: intervalo)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|Define todos os nomes de categoria para o eixo especificado.|
||[setCustomDisplayUnit (valor: número)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|Definirá a unidade de exibição de eixo a um valor personalizado.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|Especifica se o rótulo da unidade de exibição do eixo estará visível.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|Especifica a posição dos rótulos de marcas de escala no eixo especificado.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|Especifica o número de categorias ou séries entre os rótulos de marca de escala.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|Especifica o número de categorias ou séries entre marcas de escala.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Especifica se o eixo está visível.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|Código de cor HTML que representa a cor das bordas no gráfico.|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|Representa o estilo de linha da borda.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Representa a espessura da borda, em pontos.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|Valor de DataLabelPosition que representa a posição do rótulo de dados.|
||[divisória](/javascript/api/excel/excel.chartdatalabel#separator)|Cadeia de caracteres que representa o separador usado para o rótulo de dados em um gráfico.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Especifica se o tamanho da bolha do rótulo de dados é visível.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Especifica se o nome da categoria do rótulo de dados está visível.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Especifica se a tecla de legenda do rótulo de dados está visível.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Especifica se o percentual do rótulo de dados está visível.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Especifica se o nome da série do rótulo de dados é visível.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Especifica se o valor do rótulo de dados é visível.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|Representa os atributos de fonte, como nome da fonte, tamanho da fonte, cor, etc.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Especifica a altura, em pontos, da legenda no gráfico.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Especifica a esquerda, em pontos, da legenda no gráfico.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Representa uma coleção de legendEntries na legenda.|
||[Ocultar sombra](/javascript/api/excel/excel.chartlegend#showshadow)|Especifica se a legenda tem uma sombra no gráfico.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Especifica a parte superior de uma legenda de gráfico.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Especifica a largura, em pontos, da legenda no gráfico.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Representa o visível de uma entrada de legenda do gráfico.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Retorna o número de legendEntry da coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Retorna legendEntry no índice fornecido.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Representa o estilo da linha.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Representa a espessura da linha, em pontos.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Indica se um ponto de dados tem um rótulo de dados.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|Representação do código de cor HTML da cor de plano de fundo do marcador do ponto de dados (por exemplo, #FF0000 representa vermelho).|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|Representação do código de cor HTML da cor de primeiro plano do marcador do ponto de dados (por exemplo, #FF0000 representa vermelho).|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|Representa o tamanho do marcador do ponto de dados.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|Representa estilo do marcador de um ponto de dados do gráfico.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Retorna o rótulo de dados de um ponto de gráfico.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[borda](/javascript/api/excel/excel.chartpointformat#border)|Representa o formato da borda de um ponto de dados do gráfico, que inclui informações de cor, estilo e peso.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|Representa o tipo de gráfico de uma série.|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|Exclui a série de gráfico.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|Representa o tamanho do furo de rosca de uma série de gráficos.|
||[último](/javascript/api/excel/excel.chartseries#filtered)|Especifica se a série é filtrada.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|Representa a largura do espaçamento de uma série de gráfico.|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|Especifica se a série tem rótulos de dados.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|Especifica a cor de plano de fundo dos marcadores de uma série de gráficos.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|Especifica a cor de primeiro plano de marcadores de uma série de gráficos.|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|Especifica o tamanho do marcador de uma série de gráficos.|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerstyle)|Especifica o estilo de marcador de uma série de gráficos.|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|Especifica a ordem de plotagem de uma série de gráfico dentro do grupo de gráficos.|
||[Trendlines](/javascript/api/excel/excel.chartseries#trendlines)|A coleção de linhas de tendência na série.|
||[setBubbleSizes (sourceData: Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|Define os tamanhos de bolha para uma série de gráficos.|
||[SetValues (sourceData: Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|Define os valores de uma série de gráficos.|
||[setXAxisValues (sourceData: Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|Define os valores do eixo X para uma série de gráficos.|
||[Ocultar sombra](/javascript/api/excel/excel.chartseries#showshadow)|Especifica se a série tem uma sombra.|
||[suave](/javascript/api/excel/excel.chartseries#smooth)|Especifica se a série é suave.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[Add (Name?: String, index?: Number)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|Adiciona uma nova série para o conjunto.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring (início: número, comprimento: número)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|Obter a subcadeia de caracteres de um título de gráfico.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|Especifica o alinhamento horizontal do título do gráfico.|
||[left](/javascript/api/excel/excel.charttitle#left)|Especifica a distância, em pontos, da borda esquerda do título do gráfico até a borda esquerda da área do gráfico.|
||[position](/javascript/api/excel/excel.charttitle#position)|Representa a posição de título do gráfico.|
||[height](/javascript/api/excel/excel.charttitle#height)|Representa a altura, em pontos, do título do gráfico.|
||[width](/javascript/api/excel/excel.charttitle#width)|Especifica a largura, em pontos, do título do gráfico.|
||[setformula (fórmula: cadeia de caracteres)](/javascript/api/excel/excel.charttitle#setformula-formula-)|Define um valor de cadeia de caracteres que representa a fórmula do título do eixo do gráfico usando a notação no estilo A1.|
||[Ocultar sombra](/javascript/api/excel/excel.charttitle#showshadow)|Representa um valor booliano que determina se o título do gráfico tiver uma sombra.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|Especifica o ângulo no qual o texto é orientado para o título do gráfico.|
||[top](/javascript/api/excel/excel.charttitle#top)|Especifica a distância, em pontos, da borda superior do título do gráfico até a parte superior da área do gráfico.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|Especifica o alinhamento vertical do título do gráfico.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[borda](/javascript/api/excel/excel.charttitleformat#border)|Representa o formato da borda do título do gráfico, que inclui cores, LineStyle e Weight.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Deleta o objeto Trendline.|
||[detecta](/javascript/api/excel/excel.charttrendline#intercept)|Representa o valor de intercepção da linha de tendência.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Representa o período de uma tendência de gráfico.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Representa o nome da linha de tendência.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|Representa a ordem de uma tendência de gráfico.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Representa a formatação de uma linha de tendência do gráfico.|
||[type](/javascript/api/excel/excel.charttrendline#type)|Representa o tipo da linha de tendência de um gráfico.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[Add (tipo?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Adiciona uma nova linha de tendência ao conjunto de linha de tendência.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Retorna o número de linha de tendência na coleção.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Obtém o objeto da linha de tendência por índice, que é a ordem de inserção na matriz de itens.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Representa a formatação de linha do gráfico.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Exclui a propriedade personalizada.|
||[key](/javascript/api/excel/excel.customproperty#key)|A chave da propriedade personalizada.|
||[type](/javascript/api/excel/excel.customproperty#type)|O tipo de valor usado para a propriedade personalizada.|
||[value](/javascript/api/excel/excel.customproperty#value)|O valor da propriedade personalizada.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[Add (Key: String, value: any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Cria uma nova propriedade personalizada ou define uma existente.|
||[deleteAll ()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Exclui todas as propriedades personalizadas nesta coleção.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Obtém a contagem das propriedades personalizadas.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[DataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll ()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|Atualiza todas as conexões de dados da coleção.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[autor](/javascript/api/excel/excel.documentproperties#author)|O autor da pasta de trabalho.|
||[Categorias](/javascript/api/excel/excel.documentproperties#category)|A categoria da pasta de trabalho.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Os comentários da pasta de trabalho.|
||[company](/javascript/api/excel/excel.documentproperties#company)|A empresa da pasta de trabalho.|
||[Palavras-chave](/javascript/api/excel/excel.documentproperties#keywords)|As palavras-chave da pasta de trabalho.|
||[manager](/javascript/api/excel/excel.documentproperties#manager)|O gerente da pasta de trabalho.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|Obtém a data de criação da pasta de trabalho.|
||[cliente](/javascript/api/excel/excel.documentproperties#custom)|Obtém a coleção de propriedades personalizadas da pasta de trabalho.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|Obtém o último autor da pasta de trabalho.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|Obtém o número de revisão da pasta de trabalho.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|O assunto da pasta de trabalho.|
||[title](/javascript/api/excel/excel.documentproperties#title)|O título da pasta de trabalho.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|A fórmula do item nomeado.|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|Retorna um objeto que contém valores e tipos do item nomeado.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Representa os tipos de cada item na matriz de itens nomeados|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Representa os valores de cada item na matriz de itens nomeados.|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows: Number, numColumns: Number)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Obtém um objeto Range com a mesma célula superior esquerda do objeto Range atual, mas com os números especificados de linhas e colunas.|
||[GetImage ()](/javascript/api/excel/excel.range#getimage--)|Renderiza o intervalo como uma imagem png codificada em base64.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|Retorna um objeto Range que representa a região circundante da célula superior esquerda nesse intervalo.|
||[hiperlink](/javascript/api/excel/excel.range#hyperlink)|Representa o hiperlink do intervalo atual.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Representa o código de formato de número do Excel para o intervalo determinado, com base nas configurações de idioma do usuário.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Representa se o intervalo atual está em uma coluna inteira.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Representa se o intervalo atual está em uma linha inteira.|
||[Cartão ()](/javascript/api/excel/excel.range#showcard--)|Exibe o cartão para uma célula ativa se ele tiver um conteúdo valioso.|
||[style](/javascript/api/excel/excel.range#style)|Representa o estilo de intervalo atual.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|A orientação do texto de todas as células dentro do intervalo.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Determina se a altura da linha do objeto Range é igual a altura padrão da planilha.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Especifica se a largura da coluna do objeto Range é igual à largura padrão da planilha.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|Representa o destino da url do hiperlink.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|Representa o destino de referência de documento para o hiperlink.|
||[Dica](/javascript/api/excel/excel.rangehyperlink#screentip)|Representa a cadeia exibida ao passar o mouse sobre o hiperlink.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Representa a cadeia de caracteres exibida na parte superior esquerda da maioria das células no intervalo.|
|[Estilo](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|Exclui este estilo.|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|Especifica se a fórmula ficará oculta quando a planilha estiver protegida.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|Representa o alinhamento horizontal para o estilo.|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|Especifica se o estilo inclui as propriedades autoindent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel e TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|Especifica se o estilo inclui as propriedades de borda Color, ColorIndex, LineStyle e Weight.|
||[includeFont](/javascript/api/excel/excel.style#includefont)|Especifica se o estilo inclui as propriedades de plano de fundo, negrito, cor, ColorIndex, FontStyle, itálico, nome, tamanho, tachado, subscrito, sobrescrito e sublinhado.|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|Especifica se o estilo inclui a propriedade NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|Especifica se o estilo inclui as propriedades internas Color, ColorIndex, InvertIfNegative, Pattern, PatternColor e PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|Especifica se o estilo inclui as propriedades de proteção FormulaHidden e Locked.|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo do estilo.|
||[bloqueado](/javascript/api/excel/excel.style#locked)|Especifica se o objeto será bloqueado quando a planilha estiver protegida.|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|O código de formatação de formato de número para o estilo.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|O código de formato localizado do formato numérico para o estilo.|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|A ordem de leitura para o estilo.|
||[Borders](/javascript/api/excel/excel.style#borders)|Uma coleção Border de quatro objetos Border que representam o estilo das quatro bordas.|
||[Interna](/javascript/api/excel/excel.style#builtin)|Especifica se o estilo é um estilo interno.|
||[fill](/javascript/api/excel/excel.style#fill)|O preenchimento do estilo.|
||[font](/javascript/api/excel/excel.style#font)|Objeto de fonte que representa a fonte do estilo.|
||[name](/javascript/api/excel/excel.style#name)|O nome do estilo.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Especifica se o texto é automaticamente reduzido para se ajustar à largura de coluna disponível.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Especifica o alinhamento vertical para o estilo.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Especifica se o Excel quebra o texto no objeto.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Adiciona um novo estilo para o conjunto.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Obtém um estilo por nome.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Ocorre quando os dados nas células são alterados em uma tabela específica.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Ocorre quando a seleção é alterada em uma tabela específica.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|Obtém o endereço que representa a área alterada de uma tabela em uma planilha específica.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Obtém o tipo de mudança que representa como o evento Changed é acionado.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Obtém a origem do evento.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|Obtém o id da tabela na qual os dados foram alterados.|
||[tipo](/javascript/api/excel/excel.tablechangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|Obtém o id da planilha na qual os dados são alterados.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|Ocorre quando os dados são alterados em qualquer tabela em uma pasta de trabalho ou em uma planilha.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Obtém o endereço do intervalo que representa a área selecionada da tabela em uma planilha específica.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|Especifica se a seleção está dentro de uma tabela, o endereço será inútil se isinternatable for false.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|Obtém o id da tabela na qual a seleção foi alterada.|
||[tipo](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|Obtém o id da planilha na qual a seleção foi alterada.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getactivecell--)|Obtém a célula ativa no momento da pasta de trabalho.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|Representa todas as conexões de dados na pasta de trabalho.|
||[name](/javascript/api/excel/excel.workbook#name)|Obtém o nome da pasta de trabalho.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Obtém as propriedades da pasta de trabalho.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Retorna o objeto de proteção de uma pasta de trabalho.|
||[estilos](/javascript/api/excel/excel.workbook#styles)|Representa uma coleção de estilos associados à pasta de trabalho.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[proteger (senha?: cadeia de caracteres)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Protege uma pasta de trabalho.|
||[protegido](/javascript/api/excel/excel.workbookprotection#protected)|Especifica se a pasta de trabalho está protegida.|
||[desproteger (senha?: cadeia de caracteres)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Desprotege uma pasta de trabalho.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Copy (PositionType?: Excel. WorksheetPositionType, relativeTo?: Excel. Worksheet)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Copia uma planilha e a coloca na posição especificada.|
||[getRangeByIndexes (startRow: Number, startColumn: Number, rowCount: Number, columnCount: Number)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|Obtém o objeto Range que começa em um determinado índice de linha e índice de coluna e que abrange um determinado número de linhas e colunas.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|Obtém um objeto que pode ser usado para manipular painéis congelados na planilha.|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|Ocorre quando a planilha é ativada.|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|Ocorre quando os dados são alterados em uma planilha específica.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|Ocorre quando a planilha é desativada.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|Ocorre quando a seleção é alterada em uma planilha específica.|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|Retorna a altura padrão de todas as linhas na planilha, em pontos.|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|Especifica a largura padrão de todas as colunas da planilha.|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|A cor da guia da planilha.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[tipo](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|Obtém o id da planilha que está ativada.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetaddedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|Obtém o id da planilha que é adicionada à pasta de trabalho.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|Obtém o tipo de mudança que representa como o evento Changed é acionado.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|Obtém o id da planilha na qual os dados são alterados.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Ocorre quando qualquer planilha na pasta de trabalho é ativada.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Ocorre quando uma nova planilha é adicionada à pasta de trabalho.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Ocorre quando qualquer planilha na pasta de trabalho é desativada.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Ocorre quando uma planilha é excluída da pasta de trabalho.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[tipo](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|Obtém o id da planilha que está desativada.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|Obtém o id do gráfico que é excluído da pasta de trabalho.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt (frozenRange: cadeia de caracteres de intervalo \| )](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|Define as células congeladas no modo de exibição da planilha ativa.|
||[freezeColumns (contagem?: número)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|Congela a primeira colunas da planilha no local.|
||[freezeRows (contagem?: número)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|Congela as linhas superiores da planilha no local.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|
||[descongelar ()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|Remove todos os painéis congelados na planilha.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[desproteger (senha?: cadeia de caracteres)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|Desprotege uma planilha.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|Indica a opção de proteção de planilha para permitir a edição de objetos.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|Indica a opção de proteção de planilha para permitir a edição de cenários.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|Representa a opção de proteção da planilha do modo de seleção.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Obtém o endereço do intervalo que representa a área selecionada de uma planilha específica.|
||[tipo](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Obtém o id da planilha na qual a seleção foi alterada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.7&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
