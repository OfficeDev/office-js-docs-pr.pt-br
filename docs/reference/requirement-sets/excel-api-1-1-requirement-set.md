---
title: Excel Conjunto de requisitos da API JavaScript 1.1
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.1.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7d0e4bf298afe697e919c2aa557dbf10c233c807
ms.sourcegitcommit: 3b187769e86530334ca83cfdb03c1ecfac2ad9a8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/15/2021
ms.locfileid: "60367289"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Excel Conjunto de requisitos da API JavaScript 1.1

A API JavaScript do Excel 1.1 é a primeira versão da API. É o único conjunto de requisitos Excel específico com suporte Excel 2016.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.1. Para exibir a documentação de referência da API para todas as APIs com suporte Excel conjunto de requisitos da API JavaScript 1.1, consulte Excel APIs no conjunto de requisitos [1.1](/javascript/api/excel?view=excel-js-1.1&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[calculate(calculationType: Excel. CalculationType)](/javascript/api/excel/excel.application#calculate_calculationType_)|Recalcula todas as pastas de trabalho abertas no Excel no momento.|
||[calculationMode](/javascript/api/excel/excel.application#calculationMode)|Retorna o modo de cálculo usado na manual de trabalho, conforme definido pelas constantes em `Excel.CalculationMode` .|
|[Associação](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getRange__)|Retorna o intervalo representado pela associação.|
||[getTable()](/javascript/api/excel/excel.binding#getTable__)|Retorna a tabela representada pela associação.|
||[getText()](/javascript/api/excel/excel.binding#getText__)|Retorna o texto representado pela associação.|
||[id](/javascript/api/excel/excel.binding#id)|Representa o identificador de associação.|
||[tipo](/javascript/api/excel/excel.binding#type)|Retorna o tipo da associação.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[Count](/javascript/api/excel/excel.bindingcollection#count)|Retorna o número de associações da coleção.|
||[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getItem_id_)|Obtém um objeto de associação pela ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getItemAt_index_)|Obtém um objeto de associação com base em sua posição na matriz dos itens.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Chart](/javascript/api/excel/excel.chart)|[axes](/javascript/api/excel/excel.chart#axes)|Representa os eixos de um gráfico.|
||[dataLabels](/javascript/api/excel/excel.chart#dataLabels)|Representa os rótulos de dados no gráfico.|
||[delete()](/javascript/api/excel/excel.chart#delete__)|Exclui o objeto de gráfico.|
||[format](/javascript/api/excel/excel.chart#format)|Encapsula as propriedades de formato da área do gráfico.|
||[height](/javascript/api/excel/excel.chart#height)|Especifica a altura, em pontos, do objeto chart.|
||[left](/javascript/api/excel/excel.chart#left)|A distância, em pontos, da esquerda do gráfico à origem da planilha.|
||[legend](/javascript/api/excel/excel.chart#legend)|Representa a legenda do gráfico.|
||[name](/javascript/api/excel/excel.chart#name)|Especifica o nome de um objeto chart.|
||[series](/javascript/api/excel/excel.chart#series)|Representa uma única série ou uma coleção de séries no gráfico.|
||[setData(sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chart#setData_sourceData__seriesBy_)|Redefine os dados de origem do gráfico.|
||[setPosition(startCell: Range \| string, endCell?: Range \| string)](/javascript/api/excel/excel.chart#setPosition_startCell__endCell_)|Posiciona o gráfico em relação às células na planilha.|
||[title](/javascript/api/excel/excel.chart#title)|Representa o título do gráfico especificado, incluindo o respectivo texto, a visibilidade, a posição e a formatação.|
||[top](/javascript/api/excel/excel.chart#top)|Especifica a distância, em pontos, da borda superior do objeto até a parte superior da linha 1 (em uma planilha) ou a parte superior da área do gráfico (em um gráfico).|
||[width](/javascript/api/excel/excel.chart#width)|Especifica a largura, em pontos, do objeto chart.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.chartareaformat#font)|Representa os atributos de fonte do objeto atual, como nome, tamanho, cor, dentre outros.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryAxis)|Representa o eixo de categoria em um gráfico.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesAxis)|Representa o eixo de série de um gráfico 3D.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueAxis)|Representa o eixo dos valores em um eixo.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[format](/javascript/api/excel/excel.chartaxis#format)|Representa a formatação de um objeto Chart, que inclui formatação de linha e de fonte.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorGridlines)|Retorna um objeto que representa as linhas de grade principais do eixo especificado.|
||[majorUnit](/javascript/api/excel/excel.chartaxis#majorUnit)|Representa o intervalo entre as duas principais marcas de escala.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Representa o valor máximo no eixo dos valores.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Representa o valor mínimo no eixo dos valores.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorGridlines)|Retorna um objeto que representa as linhas de grade secundárias do eixo especificado.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorUnit)|Representa o intervalo entre as duas marcas de escala secundárias.|
||[title](/javascript/api/excel/excel.chartaxis#title)|Representa o título do eixo.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|Especifica os atributos de fonte (nome da fonte, tamanho da fonte, cor etc.) para um elemento de eixo do gráfico.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Especifica a formatação de linha de gráfico.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Especifica a formatação do título do eixo do gráfico.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Especifica o título do eixo.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|Especifica se o título do eixo é visibile.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|Especifica os atributos de fonte do título do eixo do gráfico, como nome da fonte, tamanho da fonte ou cor, do objeto title do eixo do gráfico.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[add(type: Excel. ChartType, sourceData: Range, seriesBy?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add_type__sourceData__seriesBy_)|Cria um novo gráfico.|
||[Count](/javascript/api/excel/excel.chartcollection#count)|Retorna o número de gráficos da planilha.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getItem_name_)|Obtém um gráfico usando o respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getItemAt_index_)|Obtém um gráfico com base em sua posição no conjunto.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Representa o formato de preenchimento do rótulo de dados atual do gráfico.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|Representa os atributos de fonte (como nome da fonte, tamanho da fonte e cor) para um rótulo de dados de gráfico.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[format](/javascript/api/excel/excel.chartdatalabels#format)|Especifica o formato dos rótulos de dados do gráfico, que inclui a formatação de preenchimento e fonte.|
||[position](/javascript/api/excel/excel.chartdatalabels#position)|Valor que representa a posição do rótulo de dados.|
||[separador](/javascript/api/excel/excel.chartdatalabels#separator)|Cadeia de caracteres que representa o separador usado para os rótulos de dados em um gráfico.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showBubbleSize)|Especifica se o tamanho da bolha do rótulo de dados está visível.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showCategoryName)|Especifica se o nome da categoria do rótulo de dados está visível.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showLegendKey)|Especifica se a chave de legenda do rótulo de dados está visível.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showPercentage)|Especifica se a porcentagem do rótulo de dados está visível.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showSeriesName)|Especifica se o nome da série de rótulos de dados está visível.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showValue)|Especifica se o valor do rótulo de dados está visível.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear__)|Limpa a cor de preenchimento de um elemento gráfico.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setSolidColor_color_)|Define a formatação de preenchimento de um elemento do gráfico com uma cor uniforme.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Representa o status da fonte em negrito.|
||[color](/javascript/api/excel/excel.chartfont#color)|Representação de código de cor HTML da cor do texto (por exemplo, #FF0000 representa Vermelho).|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Representa o status da fonte em itálico.|
||[name](/javascript/api/excel/excel.chartfont#name)|Nome da fonte (por exemplo, "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#size)|Tamanho da fonte (por exemplo, 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Tipo de sublinhado aplicado à fonte.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Representa a formatação de linhas de grade do gráfico.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Especifica se as linhas de grade do eixo estão visíveis.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Representa a formatação de linha do gráfico.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[format](/javascript/api/excel/excel.chartlegend#format)|Representa a formatação de uma legenda de gráfico, que inclui a formatação de fonte e de preenchimento.|
||[overlay](/javascript/api/excel/excel.chartlegend#overlay)|Especifica se a legenda do gráfico deve se sobrepor ao corpo principal do gráfico.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Especifica a posição da legenda no gráfico.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|Especifica se a legenda do gráfico está visível.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|Representa os atributos de fonte, como nome da fonte, tamanho da fonte e cor de uma legenda de gráfico.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear__)|Limpa o formato de linha de um elemento gráfico.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|Código de cores HTML que representa a cor das linhas no gráfico.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Encapsula as propriedades de formato de um ponto do gráfico.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Retorna o valor de um ponto do gráfico.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Representa o formato de preenchimento de um gráfico, que inclui informações de formatação em segundo plano.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[Count](/javascript/api/excel/excel.chartpointscollection#count)|Retorna o número de pontos do gráfico da série.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getItemAt_index_)|Recupera um ponto com base na respectiva posição dentro da série.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[format](/javascript/api/excel/excel.chartseries#format)|Representa a formatação de uma série do gráfico, que inclui a formatação de linha e de preenchimento.|
||[name](/javascript/api/excel/excel.chartseries#name)|Especifica o nome de uma série em um gráfico.|
||[points](/javascript/api/excel/excel.chartseries#points)|Retorna uma coleção de todos os pontos da série.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[Count](/javascript/api/excel/excel.chartseriescollection#count)|Retorna o número de série da coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getItemAt_index_)|Recupera uma série com base na respectiva posição na coleção.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Representa o formato de preenchimento de uma série do gráfico, que inclui informações sobre a formatação da tela de fundo.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Representa a formatação de linha.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[format](/javascript/api/excel/excel.charttitle#format)|Representa a formatação de um título do gráfico, que inclui a formatação de fonte e de preenchimento.|
||[overlay](/javascript/api/excel/excel.charttitle#overlay)|Especifica se o título do gráfico sobrepõe o gráfico.|
||[text](/javascript/api/excel/excel.charttitle#text)|Especifica o texto do título do gráfico.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|Especifica se o título do gráfico é visibile.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.charttitleformat#font)|Representa os atributos de fonte (como nome da fonte, tamanho da fonte e cor) de um objeto.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getRange__)|Retorna o objeto Range associado ao nome.|
||[name](/javascript/api/excel/excel.nameditem#name)|O nome do objeto.|
||[tipo](/javascript/api/excel/excel.nameditem#type)|Especifica o tipo do valor retornado pela fórmula do nome.|
||[value](/javascript/api/excel/excel.nameditem#value)|Representa o valor calculado pela fórmula do nome.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Especifica se o objeto está visível.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getItem_name_)|Obtém `NamedItem` um objeto usando seu nome.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/excel/excel.range)|[address](/javascript/api/excel/excel.range#address)|Especifica a referência de intervalo no estilo A1.|
||[addressLocal](/javascript/api/excel/excel.range#addressLocal)|Representa a referência de intervalo para o intervalo especificado no idioma do usuário.|
||[cellCount](/javascript/api/excel/excel.range#cellCount)|Especifica o número de células no intervalo.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear_applyTo_)|Limpe valores de intervalo, formatação, preenchimento, bordas, etc.|
||[columnCount](/javascript/api/excel/excel.range#columnCount)|Especifica o número total de colunas no intervalo.|
||[columnIndex](/javascript/api/excel/excel.range#columnIndex)|Especifica o número da coluna da primeira célula no intervalo.|
||[delete(shift: Excel. DeleteShiftDirection)](/javascript/api/excel/excel.range#delete_shift_)|Exclui as células associadas ao intervalo.|
||[format](/javascript/api/excel/excel.range#format)|Retorna um objeto de formato que encapsula a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades do intervalo.|
||[fórmulas](/javascript/api/excel/excel.range#formulas)|Representa a fórmula em notação A1.|
||[formulasLocal](/javascript/api/excel/excel.range#formulasLocal)|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.|
||[getBoundingRect(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getBoundingRect_anotherRange_)|Obtém o menor objeto de intervalo que abrange os intervalos determinados.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getCell_row__column_)|Obtém o objeto de intervalo que contém a célula única com base nos números de linha e de coluna.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getColumn_column_)|Obtém uma coluna incluída no intervalo.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getEntireColumn__)|Obtém um objeto que representa a coluna inteira do intervalo (por exemplo, se o intervalo atual representa as células "B4:E11", é um intervalo que representa `getEntireColumn` colunas "B:E").|
||[getEntireRow()](/javascript/api/excel/excel.range#getEntireRow__)|Obtém um objeto que representa a linha inteira do intervalo (por exemplo, se o intervalo atual representa as células "B4:E11", é um intervalo que representa linhas `GetEntireRow` "4:11").|
||[getIntersection(anotherRange: Range \| string)](/javascript/api/excel/excel.range#getIntersection_anotherRange_)|Obtém o objeto Range que representa a interseção retangular dos intervalos determinados.|
||[getLastCell()](/javascript/api/excel/excel.range#getLastCell__)|Obtém a última célula do intervalo.|
||[getLastColumn()](/javascript/api/excel/excel.range#getLastColumn__)|Obtém a última coluna do intervalo.|
||[getLastRow()](/javascript/api/excel/excel.range#getLastRow__)|Obtém a última linha do intervalo.|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getOffsetRange_rowOffset__columnOffset_)|Obtém um objeto que representa um intervalo deslocado do intervalo especificado.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getRow_row_)|Obtém uma linha contida no intervalo.|
||[insert(shift: Excel. InsertShiftDirection)](/javascript/api/excel/excel.range#insert_shift_)|Insere uma célula ou um intervalo de células na planilha, no lugar desse intervalo, e desloca as outras células para liberar espaço.|
||[numberFormat](/javascript/api/excel/excel.range#numberFormat)|Representa Excel código de formato de número para o intervalo determinado.|
||[rowCount](/javascript/api/excel/excel.range#rowCount)|Retorna o número total de linhas no intervalo.|
||[rowIndex](/javascript/api/excel/excel.range#rowIndex)|Representa o número de linhas da primeira célula no intervalo.|
||[Seleciona.](/javascript/api/excel/excel.range#select__)|Seleciona o intervalo especificado na interface do usuário do Excel.|
||[text](/javascript/api/excel/excel.range#text)|Valores de texto do intervalo especificado.|
||[valueTypes](/javascript/api/excel/excel.range#valueTypes)|Especifica o tipo de dados em cada célula.|
||[values](/javascript/api/excel/excel.range#values)|Representa os valores brutos do intervalo especificado.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|A planilha que contém o intervalo atual.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|Código de cor HTML que representa a cor da linha de borda, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideIndex)|Valor constante que indica o lado específico da borda.|
||[style](/javascript/api/excel/excel.rangeborder#style)|Uma das constantes de estilo de linha especificando o estilo de linha da borda.|
||[peso](/javascript/api/excel/excel.rangeborder#weight)|Especifica o peso da borda em torno de um intervalo.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[Count](/javascript/api/excel/excel.rangebordercollection#count)|Número de objetos de borda da coleção.|
||[getItem(index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getItem_index_)|Obtém um objeto Border usando o respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getItemAt_index_)|Obtém um objeto Border usando o respectivo índice.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear__)|Redefine a tela de fundo do intervalo.|
||[color](/javascript/api/excel/excel.rangefill#color)|Código de cor HTML que representa a cor do plano de fundo, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Representa o status em negrito da fonte.|
||[color](/javascript/api/excel/excel.rangefont#color)|Representação de código de cor HTML da cor do texto (por exemplo, #FF0000 representa Vermelho).|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Especifica o status itálico da fonte.|
||[name](/javascript/api/excel/excel.rangefont#name)|Nome da fonte (por exemplo, "Calibri").|
||[size](/javascript/api/excel/excel.rangefont#size)|Font Size|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Tipo de sublinhado aplicado à fonte.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[Borders](/javascript/api/excel/excel.rangeformat#borders)|Coleção de objetos border que se aplicam a todo o intervalo.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Retorna o objeto de preenchimento definido em todo o intervalo.|
||[font](/javascript/api/excel/excel.rangeformat#font)|Retorna o objeto font definido em todo o intervalo.|
||[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalAlignment)|Representa o alinhamento horizontal do objeto especificado.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalAlignment)|Representa o alinhamento vertical do objeto especificado.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wrapText)|Especifica se Excel quebra o texto no objeto.|
|[Table](/javascript/api/excel/excel.table)|[colunas](/javascript/api/excel/excel.table#columns)|Representa uma coleção de todas as colunas na tabela.|
||[delete()](/javascript/api/excel/excel.table#delete__)|Exclui a tabela.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getDataBodyRange__)|Obtém o objeto de intervalo associado ao corpo de dados da tabela.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getHeaderRowRange__)|Obtém o objeto de intervalo associado à linha de cabeçalho da tabela.|
||[getRange()](/javascript/api/excel/excel.table#getRange__)|Obtém o objeto de intervalo associado a toda a tabela.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#getTotalRowRange__)|Obtém o objeto de intervalo associado à linha de totais da tabela.|
||[id](/javascript/api/excel/excel.table#id)|Retorna um valor que identifica de forma exclusiva a tabela em uma determinada pasta de trabalho.|
||[name](/javascript/api/excel/excel.table#name)|Nome da tabela.|
||[rows](/javascript/api/excel/excel.table#rows)|Representa uma coleção de todas as linhas na tabela.|
||[showHeaders](/javascript/api/excel/excel.table#showHeaders)|Especifica se a linha de header está visível.|
||[showTotals](/javascript/api/excel/excel.table#showTotals)|Especifica se a linha total está visível.|
||[style](/javascript/api/excel/excel.table#style)|Valor constante que representa o estilo da tabela.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[add(address: Range \| string, hasHeaders: boolean)](/javascript/api/excel/excel.tablecollection#add_address__hasHeaders_)|Cria uma nova tabela.|
||[Count](/javascript/api/excel/excel.tablecollection#count)|Retorna o número de tabelas na pasta de trabalho.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getItem_key_)|Obtém uma tabela pelo nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getItemAt_index_)|Obtém uma tabela com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete__)|Exclui a coluna da tabela.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getDataBodyRange__)|Obtém o objeto de intervalo associado ao corpo de dados da coluna.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getHeaderRowRange__)|Obtém o objeto de intervalo associado à linha de cabeçalho da coluna.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getRange__)|Obtém o objeto de intervalo associado a toda a coluna.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#getTotalRowRange__)|Obtém o objeto de intervalo associado à linha de totais da coluna.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Retorna uma chave exclusiva que identifica a coluna na tabela.|
||[índice](/javascript/api/excel/excel.tablecolumn#index)|Retorna o número de índice da coluna na coleção de colunas da tabela.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Especifica o nome da coluna da tabela.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Representa os valores brutos do intervalo especificado.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[add(index?: number, values?: Array<Array<boolean string \| \| number>> \| boolean string \| \| number, name?: string)](/javascript/api/excel/excel.tablecolumncollection#add_index__values__name_)|Adiciona uma nova coluna à tabela.|
||[Count](/javascript/api/excel/excel.tablecolumncollection#count)|Retorna o número de colunas na tabela.|
||[getItem(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getItem_key_)|Obtém um objeto de coluna por nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getItemAt_index_)|Obtém uma coluna com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete__)|Exclui a linha da tabela.|
||[getRange()](/javascript/api/excel/excel.tablerow#getRange__)|Retorna o objeto de intervalo associado a toda a linha.|
||[índice](/javascript/api/excel/excel.tablerow#index)|Retorna o número de índice da linha na coleção de linhas da tabela.|
||[values](/javascript/api/excel/excel.tablerow#values)|Representa os valores brutos do intervalo especificado.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[add(index?: number, values?: Array<Array<boolean string \| \| number>> \| boolean string \| \| number, alwaysInsert?: boolean)](/javascript/api/excel/excel.tablerowcollection#add_index__values__alwaysInsert_)|Adiciona uma ou mais linhas à tabela.|
||[Count](/javascript/api/excel/excel.tablerowcollection#count)|Retorna o número de linhas na tabela.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getItemAt_index_)|Obtém uma linha com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Workbook](/javascript/api/excel/excel.workbook)|[aplicativo](/javascript/api/excel/excel.workbook#application)|Representa a Excel de aplicativo que contém essa workbook.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Representa uma coleção de ligações que fazem parte da pasta de trabalho.|
||[getSelectedRange()](/javascript/api/excel/excel.workbook#getSelectedRange__)|Obtém o intervalo único selecionado no momento da guia de trabalho.|
||[names](/javascript/api/excel/excel.workbook#names)|Representa uma coleção de itens nomeados com escopo de lista de trabalho (intervalos e constantes nomeados).|
||[tables](/javascript/api/excel/excel.workbook#tables)|Representa uma coleção de tabelas associadas à pasta de trabalho.|
||[planilhas](/javascript/api/excel/excel.workbook#worksheets)|Representa uma coleção de planilhas associadas à pasta de trabalho.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate__)|Ative a planilha na interface do usuário do Excel.|
||[charts](/javascript/api/excel/excel.worksheet#charts)|Retorna uma coleção de gráficos que fazem parte da planilha.|
||[delete()](/javascript/api/excel/excel.worksheet#delete__)|Exclui a planilha da pasta de trabalho.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getCell_row__column_)|Obtém `Range` o objeto que contém a única célula com base nos números de linha e coluna.|
||[getRange(address?: string)](/javascript/api/excel/excel.worksheet#getRange_address_)|Obtém `Range` o objeto, representando um único bloco retangular de células, especificado pelo endereço ou nome.|
||[id](/javascript/api/excel/excel.worksheet#id)|Retorna um valor que identifica de forma exclusiva a planilha em uma determinada pasta de trabalho.|
||[name](/javascript/api/excel/excel.worksheet#name)|O nome de exibição da planilha.|
||[position](/javascript/api/excel/excel.worksheet#position)|A posição baseada em zero da planilha na pasta de trabalho.|
||[tables](/javascript/api/excel/excel.worksheet#tables)|Coleção de tabelas que fazem parte da planilha.|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|A visibilidade da planilha.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[add(name?: string)](/javascript/api/excel/excel.worksheetcollection#add_name_)|Adiciona uma nova planilha à pasta de trabalho.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getActiveWorksheet__)|Obtém a planilha ativa no momento na pasta de trabalho.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getItem_key_)|Obtém um objeto de planilha usando o nome ou ID dele.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
