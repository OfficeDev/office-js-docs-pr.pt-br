---
title: Excel conjunto de requisitos da API JavaScript 1.1
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.1.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 45061afc7e401e18a67377bf88fa1670bb7a8ece
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745954"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Excel conjunto de requisitos da API JavaScript 1.1

A API JavaScript do Excel 1.1 é a primeira versão da API. É o único conjunto de requisitos Excel específico com suporte Excel 2016.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.1. Para exibir a documentação de referência da API para todas as APIs com suporte Excel conjunto de requisitos da API JavaScript 1.1, consulte Excel APIs no conjunto de requisitos [1.1](/javascript/api/excel?view=excel-js-1.1&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[calculate(calculationType: Excel. CalculationType)](/javascript/api/excel/excel.application#excel-excel-application-calculate-member(1))|Recalcula todas as pastas de trabalho abertas no Excel no momento.|
||[calculationMode](/javascript/api/excel/excel.application#excel-excel-application-calculationmode-member)|Retorna o modo de cálculo usado na manual de trabalho, conforme definido pelas constantes em `Excel.CalculationMode`.|
|[Associação](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#excel-excel-binding-getrange-member(1))|Retorna o intervalo representado pela associação.|
||[getTable()](/javascript/api/excel/excel.binding#excel-excel-binding-gettable-member(1))|Retorna a tabela representada pela associação.|
||[getText()](/javascript/api/excel/excel.binding#excel-excel-binding-gettext-member(1))|Retorna o texto representado pela associação.|
||[id](/javascript/api/excel/excel.binding#excel-excel-binding-id-member)|Representa o identificador de associação.|
||[tipo](/javascript/api/excel/excel.binding#excel-excel-binding-type-member)|Retorna o tipo da associação.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[Count](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-count-member)|Retorna o número de associações da coleção.|
||[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitem-member(1))|Obtém um objeto de associação pela ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemat-member(1))|Obtém um objeto de associação com base em sua posição na matriz dos itens.|
||[items](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Chart](/javascript/api/excel/excel.chart)|[axes](/javascript/api/excel/excel.chart#excel-excel-chart-axes-member)|Representa os eixos de um gráfico.|
||[dataLabels](/javascript/api/excel/excel.chart#excel-excel-chart-datalabels-member)|Representa os rótulos de dados no gráfico.|
||[delete()](/javascript/api/excel/excel.chart#excel-excel-chart-delete-member(1))|Exclui o objeto de gráfico.|
||[format](/javascript/api/excel/excel.chart#excel-excel-chart-format-member)|Encapsula as propriedades de formato da área do gráfico.|
||[height](/javascript/api/excel/excel.chart#excel-excel-chart-height-member)|Especifica a altura, em pontos, do objeto chart.|
||[left](/javascript/api/excel/excel.chart#excel-excel-chart-left-member)|A distância, em pontos, da esquerda do gráfico à origem da planilha.|
||[legend](/javascript/api/excel/excel.chart#excel-excel-chart-legend-member)|Representa a legenda do gráfico.|
||[name](/javascript/api/excel/excel.chart#excel-excel-chart-name-member)|Especifica o nome de um objeto chart.|
||[series](/javascript/api/excel/excel.chart#excel-excel-chart-series-member)|Representa uma única série ou uma coleção de séries no gráfico.|
||[setData(sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chart#excel-excel-chart-setdata-member(1))|Redefine os dados de origem do gráfico.|
||[setPosition(startCell: Range \| string, endCell?: Range \| string)](/javascript/api/excel/excel.chart#excel-excel-chart-setposition-member(1))|Posiciona o gráfico em relação às células na planilha.|
||[title](/javascript/api/excel/excel.chart#excel-excel-chart-title-member)|Representa o título do gráfico especificado, incluindo o respectivo texto, a visibilidade, a posição e a formatação.|
||[top](/javascript/api/excel/excel.chart#excel-excel-chart-top-member)|Especifica a distância, em pontos, da borda superior do objeto até a parte superior da linha 1 (em uma planilha) ou a parte superior da área do gráfico (em um gráfico).|
||[width](/javascript/api/excel/excel.chart#excel-excel-chart-width-member)|Especifica a largura, em pontos, do objeto chart.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-fill-member)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-font-member)|Representa os atributos de fonte do objeto atual, como nome, tamanho, cor, dentre outros.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-categoryaxis-member)|Representa o eixo de categoria em um gráfico.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-seriesaxis-member)|Representa o eixo de série de um gráfico 3D.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#excel-excel-chartaxes-valueaxis-member)|Representa o eixo dos valores em um eixo.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[format](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-format-member)|Representa a formatação de um objeto Chart, que inclui formatação de linha e de fonte.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorgridlines-member)|Retorna um objeto que representa as linhas de grade principais do eixo especificado.|
||[majorUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-majorunit-member)|Representa o intervalo entre as duas principais marcas de escala.|
||[maximum](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-maximum-member)|Representa o valor máximo no eixo dos valores.|
||[minimum](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minimum-member)|Representa o valor mínimo no eixo dos valores.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorgridlines-member)|Retorna um objeto que representa as linhas de grade secundárias do eixo especificado.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-minorunit-member)|Representa o intervalo entre as duas marcas de escala secundárias.|
||[title](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-title-member)|Representa o título do eixo.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-font-member)|Especifica os atributos de fonte (nome da fonte, tamanho da fonte, cor etc.) para um elemento de eixo do gráfico.|
||[line](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-line-member)|Especifica a formatação de linha de gráfico.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-format-member)|Especifica a formatação do título do eixo do gráfico.|
||[text](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-text-member)|Especifica o título do eixo.|
||[visible](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-visible-member)|Especifica se o título do eixo é visibile.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-font-member)|Especifica os atributos de fonte do título do eixo do gráfico, como nome da fonte, tamanho da fonte ou cor, do objeto title do eixo do gráfico.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: Excel. ChartType, sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-add-member(1))|Cria um novo gráfico.|
||[Count](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-count-member)|Retorna o número de gráficos da planilha.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitem-member(1))|Obtém um gráfico usando o respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemat-member(1))|Obtém um gráfico com base em sua posição no conjunto.|
||[items](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-fill-member)|Representa o formato de preenchimento do rótulo de dados atual do gráfico.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-font-member)|Representa os atributos de fonte (como nome da fonte, tamanho da fonte e cor) para um rótulo de dados de gráfico.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[format](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-format-member)|Especifica o formato dos rótulos de dados do gráfico, que inclui a formatação de preenchimento e fonte.|
||[position](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-position-member)|Valor que representa a posição do rótulo de dados.|
||[separador](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-separator-member)|Cadeia de caracteres que representa o separador usado para os rótulos de dados em um gráfico.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showbubblesize-member)|Especifica se o tamanho da bolha do rótulo de dados está visível.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showcategoryname-member)|Especifica se o nome da categoria do rótulo de dados está visível.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showlegendkey-member)|Especifica se a chave de legenda do rótulo de dados está visível.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showpercentage-member)|Especifica se a porcentagem do rótulo de dados está visível.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showseriesname-member)|Especifica se o nome da série de rótulos de dados está visível.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-showvalue-member)|Especifica se o valor do rótulo de dados está visível.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-clear-member(1))|Limpa a cor de preenchimento de um elemento gráfico.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#excel-excel-chartfill-setsolidcolor-member(1))|Define a formatação de preenchimento de um elemento do gráfico com uma cor uniforme.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-bold-member)|Representa o status da fonte em negrito.|
||[color](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-color-member)|Representação de código de cor HTML da cor do texto (por exemplo, #FF0000 representa Vermelho).|
||[italic](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-italic-member)|Representa o status da fonte em itálico.|
||[name](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-name-member)|Nome da fonte (por exemplo, "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-size-member)|Tamanho da fonte (por exemplo, 11)|
||[underline](/javascript/api/excel/excel.chartfont#excel-excel-chartfont-underline-member)|Tipo de sublinhado aplicado à fonte.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-format-member)|Representa a formatação de linhas de grade do gráfico.|
||[visible](/javascript/api/excel/excel.chartgridlines#excel-excel-chartgridlines-visible-member)|Especifica se as linhas de grade do eixo estão visíveis.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#excel-excel-chartgridlinesformat-line-member)|Representa a formatação de linha do gráfico.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[format](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-format-member)|Representa a formatação de uma legenda de gráfico, que inclui a formatação de fonte e de preenchimento.|
||[overlay](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-overlay-member)|Especifica se a legenda do gráfico deve se sobrepor ao corpo principal do gráfico.|
||[position](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-position-member)|Especifica a posição da legenda no gráfico.|
||[visible](/javascript/api/excel/excel.chartlegend#excel-excel-chartlegend-visible-member)|Especifica se a legenda do gráfico está visível.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-fill-member)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-font-member)|Representa os atributos de fonte, como nome da fonte, tamanho da fonte e cor de uma legenda de gráfico.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-clear-member(1))|Limpa o formato de linha de um elemento gráfico.|
||[color](/javascript/api/excel/excel.chartlineformat#excel-excel-chartlineformat-color-member)|Código de cores HTML que representa a cor das linhas no gráfico.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-format-member)|Encapsula as propriedades de formato de um ponto do gráfico.|
||[value](/javascript/api/excel/excel.chartpoint#excel-excel-chartpoint-value-member)|Retorna o valor de um ponto do gráfico.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#excel-excel-chartpointformat-fill-member)|Representa o formato de preenchimento de um gráfico, que inclui informações de formatação em segundo plano.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[Count](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-count-member)|Retorna o número de pontos do gráfico da série.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getitemat-member(1))|Recupera um ponto com base na respectiva posição dentro da série.|
||[items](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[format](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-format-member)|Representa a formatação de uma série do gráfico, que inclui a formatação de linha e de preenchimento.|
||[name](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-name-member)|Especifica o nome de uma série em um gráfico.|
||[points](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-points-member)|Retorna uma coleção de todos os pontos da série.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[Count](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-count-member)|Retorna o número de série da coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getitemat-member(1))|Recupera uma série com base na respectiva posição na coleção.|
||[items](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#excel-excel-chartseriesformat-fill-member)|Representa o formato de preenchimento de uma série do gráfico, que inclui informações sobre a formatação da tela de fundo.|
||[line](/javascript/api/excel/excel.chartseriesformat#excel-excel-chartseriesformat-line-member)|Representa a formatação de linha.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[format](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-format-member)|Representa a formatação de um título do gráfico, que inclui a formatação de fonte e de preenchimento.|
||[overlay](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-overlay-member)|Especifica se o título do gráfico sobrepõe o gráfico.|
||[text](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-text-member)|Especifica o texto do título do gráfico.|
||[visible](/javascript/api/excel/excel.charttitle#excel-excel-charttitle-visible-member)|Especifica se o título do gráfico é visibile.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-fill-member)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.charttitleformat#excel-excel-charttitleformat-font-member)|Representa os atributos de fonte (como nome da fonte, tamanho da fonte e cor) de um objeto.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrange-member(1))|Retorna o objeto Range associado ao nome.|
||[name](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-name-member)|O nome do objeto.|
||[tipo](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-type-member)|Especifica o tipo do valor retornado pela fórmula do nome.|
||[value](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-value-member)|Representa o valor calculado pela fórmula do nome.|
||[visible](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-visible-member)|Especifica se o objeto está visível.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitem-member(1))|Obtém um `NamedItem` objeto usando seu nome.|
||[items](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/excel/excel.range)|[address](/javascript/api/excel/excel.range#excel-excel-range-address-member)|Especifica a referência de intervalo no estilo A1.|
||[addressLocal](/javascript/api/excel/excel.range#excel-excel-range-addresslocal-member)|Representa a referência de intervalo para o intervalo especificado no idioma do usuário.|
||[cellCount](/javascript/api/excel/excel.range#excel-excel-range-cellcount-member)|Especifica o número de células no intervalo.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#excel-excel-range-clear-member(1))|Limpe valores de intervalo, formatação, preenchimento, bordas, etc.|
||[columnCount](/javascript/api/excel/excel.range#excel-excel-range-columncount-member)|Especifica o número total de colunas no intervalo.|
||[columnIndex](/javascript/api/excel/excel.range#excel-excel-range-columnindex-member)|Especifica o número da coluna da primeira célula no intervalo.|
||[delete(shift: Excel. DeleteShiftDirection)](/javascript/api/excel/excel.range#excel-excel-range-delete-member(1))|Exclui as células associadas ao intervalo.|
||[format](/javascript/api/excel/excel.range#excel-excel-range-format-member)|Retorna um objeto de formato que encapsula a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades do intervalo.|
||[fórmulas](/javascript/api/excel/excel.range#excel-excel-range-formulas-member)|Representa a fórmula em notação A1.|
||[formulasLocal](/javascript/api/excel/excel.range#excel-excel-range-formulaslocal-member)|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.|
||[getBoundingRect(anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getboundingrect-member(1))|Obtém o menor objeto de intervalo que abrange os intervalos determinados.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#excel-excel-range-getcell-member(1))|Obtém o objeto de intervalo que contém a célula única com base nos números de linha e de coluna.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#excel-excel-range-getcolumn-member(1))|Obtém uma coluna incluída no intervalo.|
||[getEntireColumn()](/javascript/api/excel/excel.range#excel-excel-range-getentirecolumn-member(1))|Obtém um objeto que representa a coluna inteira do intervalo (por exemplo, se o intervalo atual representa as células "B4:E11", `getEntireColumn` é um intervalo que representa colunas "B:E").|
||[getEntireRow()](/javascript/api/excel/excel.range#excel-excel-range-getentirerow-member(1))|Obtém um objeto que representa a linha inteira do intervalo (por exemplo, se o intervalo atual representa as células "B4:E11", `GetEntireRow` é um intervalo que representa linhas "4:11").|
||[getIntersection(anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getintersection-member(1))|Obtém o objeto Range que representa a interseção retangular dos intervalos determinados.|
||[getLastCell()](/javascript/api/excel/excel.range#excel-excel-range-getlastcell-member(1))|Obtém a última célula do intervalo.|
||[getLastColumn()](/javascript/api/excel/excel.range#excel-excel-range-getlastcolumn-member(1))|Obtém a última coluna do intervalo.|
||[getLastRow()](/javascript/api/excel/excel.range#excel-excel-range-getlastrow-member(1))|Obtém a última linha do intervalo.|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#excel-excel-range-getoffsetrange-member(1))|Obtém um objeto que representa um intervalo deslocado do intervalo especificado.|
||[getRow(row: number)](/javascript/api/excel/excel.range#excel-excel-range-getrow-member(1))|Obtém uma linha contida no intervalo.|
||[insert(shift: Excel. InsertShiftDirection)](/javascript/api/excel/excel.range#excel-excel-range-insert-member(1))|Insere uma célula ou um intervalo de células na planilha, no lugar desse intervalo, e desloca as outras células para liberar espaço.|
||[numberFormat](/javascript/api/excel/excel.range#excel-excel-range-numberformat-member)|Representa Excel código de formato de número para o intervalo determinado.|
||[rowCount](/javascript/api/excel/excel.range#excel-excel-range-rowcount-member)|Retorna o número total de linhas no intervalo.|
||[rowIndex](/javascript/api/excel/excel.range#excel-excel-range-rowindex-member)|Representa o número de linhas da primeira célula no intervalo.|
||[Seleciona.](/javascript/api/excel/excel.range#excel-excel-range-select-member(1))|Seleciona o intervalo especificado na interface do usuário do Excel.|
||[text](/javascript/api/excel/excel.range#excel-excel-range-text-member)|Valores de texto do intervalo especificado.|
||[valueTypes](/javascript/api/excel/excel.range#excel-excel-range-valuetypes-member)|Especifica o tipo de dados em cada célula.|
||[values](/javascript/api/excel/excel.range#excel-excel-range-values-member)|Representa os valores brutos do intervalo especificado.|
||[worksheet](/javascript/api/excel/excel.range#excel-excel-range-worksheet-member)|A planilha que contém o intervalo atual.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-color-member)|Código de cor HTML que representa a cor da linha de borda, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[sideIndex](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-sideindex-member)|Valor constante que indica o lado específico da borda.|
||[style](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-style-member)|Uma das constantes de estilo de linha especificando o estilo de linha da borda.|
||[peso](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-weight-member)|Especifica o peso da borda em torno de um intervalo.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[Count](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-count-member)|Número de objetos de borda da coleção.|
||[getItem(index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-getitem-member(1))|Obtém um objeto Border usando o respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-getitemat-member(1))|Obtém um objeto Border usando o respectivo índice.|
||[items](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-clear-member(1))|Redefine a tela de fundo do intervalo.|
||[color](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-color-member)|Código de cor HTML que representa a cor do plano de fundo, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-bold-member)|Representa o status em negrito da fonte.|
||[color](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-color-member)|Representação de código de cor HTML da cor do texto (por exemplo, #FF0000 representa Vermelho).|
||[italic](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-italic-member)|Especifica o status itálico da fonte.|
||[name](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-name-member)|Nome da fonte (por exemplo, "Calibri").|
||[size](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-size-member)|Font Size|
||[underline](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-underline-member)|Tipo de sublinhado aplicado à fonte.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[Borders](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-borders-member)|Coleção de objetos border que se aplicam a todo o intervalo.|
||[fill](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-fill-member)|Retorna o objeto de preenchimento definido em todo o intervalo.|
||[font](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-font-member)|Retorna o objeto font definido em todo o intervalo.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-horizontalalignment-member)|Representa o alinhamento horizontal do objeto especificado.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-verticalalignment-member)|Representa o alinhamento vertical do objeto especificado.|
||[wrapText](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-wraptext-member)|Especifica se Excel quebra o texto no objeto.|
|[Table](/javascript/api/excel/excel.table)|[columns](/javascript/api/excel/excel.table#excel-excel-table-columns-member)|Representa uma coleção de todas as colunas na tabela.|
||[delete()](/javascript/api/excel/excel.table#excel-excel-table-delete-member(1))|Exclui a tabela.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#excel-excel-table-getdatabodyrange-member(1))|Obtém o objeto de intervalo associado ao corpo de dados da tabela.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#excel-excel-table-getheaderrowrange-member(1))|Obtém o objeto de intervalo associado à linha de cabeçalho da tabela.|
||[getRange()](/javascript/api/excel/excel.table#excel-excel-table-getrange-member(1))|Obtém o objeto de intervalo associado a toda a tabela.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#excel-excel-table-gettotalrowrange-member(1))|Obtém o objeto de intervalo associado à linha de totais da tabela.|
||[id](/javascript/api/excel/excel.table#excel-excel-table-id-member)|Retorna um valor que identifica de forma exclusiva a tabela em uma determinada pasta de trabalho.|
||[name](/javascript/api/excel/excel.table#excel-excel-table-name-member)|Nome da tabela.|
||[rows](/javascript/api/excel/excel.table#excel-excel-table-rows-member)|Representa uma coleção de todas as linhas na tabela.|
||[showHeaders](/javascript/api/excel/excel.table#excel-excel-table-showheaders-member)|Especifica se a linha de header está visível.|
||[showTotals](/javascript/api/excel/excel.table#excel-excel-table-showtotals-member)|Especifica se a linha total está visível.|
||[style](/javascript/api/excel/excel.table#excel-excel-table-style-member)|Valor constante que representa o estilo da tabela.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-add-member(1))|Cria uma nova tabela.|
||[Count](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-count-member)|Retorna o número de tabelas na pasta de trabalho.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitem-member(1))|Obtém uma tabela pelo nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemat-member(1))|Obtém uma tabela com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-delete-member(1))|Exclui a coluna da tabela.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getdatabodyrange-member(1))|Obtém o objeto de intervalo associado ao corpo de dados da coluna.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getheaderrowrange-member(1))|Obtém o objeto de intervalo associado à linha de cabeçalho da coluna.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-getrange-member(1))|Obtém o objeto de intervalo associado a toda a coluna.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-gettotalrowrange-member(1))|Obtém o objeto de intervalo associado à linha de totais da coluna.|
||[id](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-id-member)|Retorna uma chave exclusiva que identifica a coluna na tabela.|
||[índice](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-index-member)|Retorna o número de índice da coluna na coleção de colunas da tabela.|
||[name](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-name-member)|Especifica o nome da coluna da tabela.|
||[values](/javascript/api/excel/excel.tablecolumn#excel-excel-tablecolumn-values-member)|Representa os valores brutos do intervalo especificado.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<Array<boolean string number>> \| boolean \| \| \| string \| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-add-member(1))|Adiciona uma nova coluna à tabela.|
||[Count](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-count-member)|Retorna o número de colunas na tabela.|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitem-member(1))|Obtém um objeto de coluna por nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemat-member(1))|Obtém uma coluna com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-delete-member(1))|Exclui a linha da tabela.|
||[getRange()](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-getrange-member(1))|Retorna o objeto de intervalo associado a toda a linha.|
||[índice](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-index-member)|Retorna o número de índice da linha na coleção de linhas da tabela.|
||[values](/javascript/api/excel/excel.tablerow#excel-excel-tablerow-values-member)|Representa os valores brutos do intervalo especificado.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<Array<boolean string number>> \| boolean \| \| \| string \| number, alwaysInsert?: boolean)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-add-member(1))|Adiciona uma ou mais linhas à tabela.|
||[Count](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-count-member)|Retorna o número de linhas na tabela.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getitemat-member(1))|Obtém uma linha com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[Workbook](/javascript/api/excel/excel.workbook)|[application](/javascript/api/excel/excel.workbook#excel-excel-workbook-application-member)|Representa a Excel de aplicativo que contém essa workbook.|
||[bindings](/javascript/api/excel/excel.workbook#excel-excel-workbook-bindings-member)|Representa uma coleção de ligações que fazem parte da pasta de trabalho.|
||[getSelectedRange()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1))|Obtém o intervalo único selecionado no momento da guia de trabalho.|
||[names](/javascript/api/excel/excel.workbook#excel-excel-workbook-names-member)|Representa uma coleção de itens nomeados com escopo de lista de trabalho (intervalos e constantes nomeados).|
||[tables](/javascript/api/excel/excel.workbook#excel-excel-workbook-tables-member)|Representa uma coleção de tabelas associadas à pasta de trabalho.|
||[planilhas](/javascript/api/excel/excel.workbook#excel-excel-workbook-worksheets-member)|Representa uma coleção de planilhas associadas à pasta de trabalho.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-activate-member(1))|Ative a planilha na interface do usuário do Excel.|
||[charts](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-charts-member)|Retorna uma coleção de gráficos que fazem parte da planilha.|
||[delete()](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-delete-member(1))|Exclui a planilha da pasta de trabalho.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getcell-member(1))|Obtém `Range` o objeto que contém a única célula com base nos números de linha e coluna.|
||[getRange(address?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1))|Obtém `Range` o objeto, representando um único bloco retangular de células, especificado pelo endereço ou nome.|
||[id](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-id-member)|Retorna um valor que identifica de forma exclusiva a planilha em uma determinada pasta de trabalho.|
||[name](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-name-member)|O nome de exibição da planilha.|
||[position](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-position-member)|A posição baseada em zero da planilha na pasta de trabalho.|
||[tables](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tables-member)|Coleção de tabelas que fazem parte da planilha.|
||[visibility](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-visibility-member)|A visibilidade da planilha.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-add-member(1))|Adiciona uma nova planilha à pasta de trabalho.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getactiveworksheet-member(1))|Obtém a planilha ativa no momento na pasta de trabalho.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitem-member(1))|Obtém um objeto de planilha usando o nome ou ID dele.|
||[items](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
