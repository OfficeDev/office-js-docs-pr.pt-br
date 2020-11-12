---
title: Conjunto de requisitos de API JavaScript do Excel 1,1
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1,1.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 648013738729961a2d36897534f500dd025cab75
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996241"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Conjunto de requisitos de API JavaScript do Excel 1,1

A API JavaScript do Excel 1.1 é a primeira versão da API. É o único conjunto de requisitos específico do Excel suportado pelo Excel 2016.

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Excel 1,1. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Excel 1,1, confira [APIs do Excel no conjunto de requisitos 1,1](/javascript/api/excel?view=excel-js-1.1&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[Calculate (calculatype: Excel. Calculatype)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Recalcula todas as pastas de trabalho abertas no Excel no momento.|
||[cálculomode](/javascript/api/excel/excel.application#calculationmode)|Retorna o modo de cálculo usado na pasta de trabalho, conforme definido pelas constantes no Excel. Calculation.|
|[Associação](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|Retorna o intervalo representado pela associação.|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|Retorna a tabela representada pela associação.|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|Retorna o texto representado pela associação.|
||[id](/javascript/api/excel/excel.binding#id)|Representa um identificador de associação.|
||[type](/javascript/api/excel/excel.binding#type)|Retorna o tipo da associação.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|Obtém um objeto de associação pela ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|Obtém um objeto de associação com base em sua posição na matriz dos itens.|
||[Count](/javascript/api/excel/excel.bindingcollection#count)|Retorna o número de associações da coleção.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|Exclui o objeto de gráfico.|
||[height](/javascript/api/excel/excel.chart#height)|Especifica a altura, em pontos, do objeto de gráfico.|
||[left](/javascript/api/excel/excel.chart#left)|A distância, em pontos, da esquerda do gráfico à origem da planilha.|
||[name](/javascript/api/excel/excel.chart#name)|Especifica o nome de um objeto de gráfico.|
||[Axes](/javascript/api/excel/excel.chart#axes)|Representa os eixos de um gráfico.|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|Representa os rótulos de dados no gráfico.|
||[format](/javascript/api/excel/excel.chart#format)|Encapsula as propriedades de formato da área do gráfico.|
||[Legenda](/javascript/api/excel/excel.chart#legend)|Representa a legenda do gráfico.|
||[série](/javascript/api/excel/excel.chart#series)|Representa uma única série ou uma coleção de séries no gráfico.|
||[title](/javascript/api/excel/excel.chart#title)|Especifica o título do gráfico especificado, incluindo o texto, a visibilidade, a posição e a formatação do título.|
||[setData (sourceData: Range, Seriesby como?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Redefine os dados de origem do gráfico.|
||[SETPOSITION (startCell: \| String de intervalo, endcell?: \| cadeia de caracteres de intervalo)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|Posiciona o gráfico em relação às células na planilha.|
||[top](/javascript/api/excel/excel.chart#top)|Especifica a distância, em pontos, da borda superior do objeto até a parte superior da linha 1 (em uma planilha) ou a parte superior da área do gráfico (em um gráfico).|
||[width](/javascript/api/excel/excel.chart#width)|Especifica a largura, em pontos, do objeto de gráfico.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.chartareaformat#font)|Representa os atributos de fonte do objeto atual, como nome, tamanho, cor, dentre outros.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|Representa o eixo de categoria em um gráfico.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesaxis)|Representa o eixo das séries de um gráfico 3D.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueaxis)|Representa o eixo dos valores em um eixo.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|Representa o intervalo entre as duas principais marcas de escala.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Representa o valor máximo no eixo dos valores.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Representa o valor mínimo no eixo dos valores.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|Representa o intervalo entre as duas marcas de escala secundárias.|
||[format](/javascript/api/excel/excel.chartaxis#format)|Representa a formatação de um objeto Chart, que inclui formatação de linha e de fonte.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|Retorna um objeto Gridlines que representa as principais linhas de grade do eixo especificado.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|Retorna um objeto Gridlines que representa as linhas de grade secundárias do eixo especificado.|
||[title](/javascript/api/excel/excel.chartaxis#title)|Representa o título do eixo.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|Especifica os atributos de fonte (nome da fonte, tamanho da fonte, cor, etc.) para um elemento de eixo do gráfico.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Especifica a formatação da linha do gráfico.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Especifica a formatação do título do eixo do gráfico.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Especifica o título do eixo.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|Especifica se o título do eixo é visibile.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|Especifica os atributos de fonte do título do eixo do gráfico, como nome da fonte, tamanho da fonte, cor, etc.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[Add (tipo: Excel. ChartType, sourceData: Range, Seriesby como?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Cria um novo gráfico.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|Obtém um gráfico usando o respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|Obtém um gráfico com base em sua posição no conjunto.|
||[Count](/javascript/api/excel/excel.chartcollection#count)|Retorna o número de gráficos da planilha.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Representa o formato de preenchimento do rótulo de dados atual do gráfico.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|Representa os atributos de fonte do rótulo de dados do gráfico, como nome, tamanho, cor, dentre outros.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|Valor de DataLabelPosition que representa a posição do rótulo de dados.|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|Especifica o formato dos rótulos de dados do gráfico, que inclui a formatação de fonte e preenchimento.|
||[divisória](/javascript/api/excel/excel.chartdatalabels#separator)|Cadeia de caracteres que representa o separador usado para os rótulos de dados em um gráfico.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|Especifica se o tamanho da bolha do rótulo de dados é visível.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|Especifica se o nome da categoria do rótulo de dados está visível.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|Especifica se a tecla de legenda do rótulo de dados está visível.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|Especifica se o percentual do rótulo de dados está visível.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|Especifica se o nome da série do rótulo de dados é visível.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|Especifica se o valor do rótulo de dados é visível.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|Limpa a cor de preenchimento de um elemento do gráfico.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|Define a formatação de preenchimento de um elemento do gráfico com uma cor uniforme.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Representa o status da fonte em negrito.|
||[color](/javascript/api/excel/excel.chartfont#color)|Representação do código de cor HTML da cor do texto (por exemplo, #FF0000 representa vermelho).|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Representa o status da fonte em itálico.|
||[name](/javascript/api/excel/excel.chartfont#name)|Nome da fonte (por exemplo, "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#size)|Tamanho da fonte (por exemplo, 11)|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Tipo de sublinhado aplicado à fonte.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Representa a formatação de linhas de grade do gráfico.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Especifica se as linhas de grade do eixo estão visíveis.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Representa a formatação de linha do gráfico.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[overlay](/javascript/api/excel/excel.chartlegend#overlay)|Especifica se a legenda do gráfico deve se sobrepor ao corpo principal do gráfico.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Especifica a posição da legenda no gráfico.|
||[format](/javascript/api/excel/excel.chartlegend#format)|Representa a formatação de uma legenda de gráfico, que inclui a formatação de fonte e de preenchimento.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|Especifica se o ChartLegend está visível.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|Representa os atributos de fonte, como nome da fonte, tamanho da fonte, cor, etc.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|Limpar o formato da linha de um elemento do gráfico.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|Código de cores HTML que representa a cor das linhas no gráfico.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Encapsula as propriedades de formato de um ponto do gráfico.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Retorna o valor de um ponto do gráfico.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Representa o formato de preenchimento de um gráfico, que inclui informações de formatação de plano de fundo.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|Recupera um ponto com base na respectiva posição dentro da série.|
||[Count](/javascript/api/excel/excel.chartpointscollection#count)|Retorna o número de pontos do gráfico da série.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|Especifica o nome de uma série em um gráfico.|
||[format](/javascript/api/excel/excel.chartseries#format)|Representa a formatação de uma série do gráfico, que inclui a formatação de linha e de preenchimento.|
||[pontos](/javascript/api/excel/excel.chartseries#points)|Retorna uma coleção de todos os pontos da série.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|Recupera uma série com base na respectiva posição na coleção.|
||[Count](/javascript/api/excel/excel.chartseriescollection#count)|Retorna o número de série da coleção.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Representa o formato de preenchimento de uma série do gráfico, que inclui informações sobre a formatação da tela de fundo.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Representa a formatação de linha.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[overlay](/javascript/api/excel/excel.charttitle#overlay)|Especifica se o título do gráfico será sobreposto ao gráfico.|
||[format](/javascript/api/excel/excel.charttitle#format)|Representa a formatação de um título do gráfico, que inclui a formatação de fonte e de preenchimento.|
||[text](/javascript/api/excel/excel.charttitle#text)|Especifica o texto do título do gráfico.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|Especifica se o título do gráfico é visibile.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[font](/javascript/api/excel/excel.charttitleformat#font)|Representa os atributos de fonte de um objeto, como nome, tamanho, cor, dentre outros.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|Retorna o objeto Range associado ao nome.|
||[name](/javascript/api/excel/excel.nameditem#name)|O nome do objeto.|
||[type](/javascript/api/excel/excel.nameditem#type)|Especifica o tipo do valor retornado pela fórmula do nome.|
||[value](/javascript/api/excel/excel.nameditem#value)|Representa o valor calculado pela fórmula do nome.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Especifica se o objeto está visível.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|Obtém um objeto NamedItem usando seu nome.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|Limpe valores de intervalo, formatação, preenchimento, bordas, etc.|
||[excluir (Shift: Excel. DeleteShiftDirection)](/javascript/api/excel/excel.range#delete-shift-)|Exclui as células associadas ao intervalo.|
||[fórmulas](/javascript/api/excel/excel.range#formulas)|Representa a fórmula em notação A1.|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.|
||[getBoundingRect (anotherRange: cadeia de caracteres de intervalo \| )](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|Obtém o menor objeto de intervalo que abrange os intervalos determinados.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|Obtém o objeto de intervalo que contém a célula única com base nos números de linha e de coluna.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|Obtém uma coluna incluída no intervalo.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|Obtém um objeto que representa a coluna inteira do intervalo (por exemplo, se o intervalo atual representa as células "B4: E11", seu `getEntireColumn` é um intervalo que representa as colunas "B:E").|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|Obtém um objeto que representa a linha inteira do intervalo (por exemplo, se o intervalo atual representa as células "B4: E11", seu `GetEntireRow` é um intervalo que representa as linhas "4:11").|
||[getintersection (anotherRange: cadeia de caracteres de intervalo \| )](/javascript/api/excel/excel.range#getintersection-anotherrange-)|Obtém o objeto Range que representa a interseção retangular dos intervalos determinados.|
||[getLastCell()](/javascript/api/excel/excel.range#getlastcell--)|Obtém a última célula do intervalo.|
||[getLastColumn()](/javascript/api/excel/excel.range#getlastcolumn--)|Obtém a última coluna do intervalo.|
||[getLastRow()](/javascript/api/excel/excel.range#getlastrow--)|Obtém a última linha do intervalo.|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|Obtém um objeto que representa um intervalo deslocado do intervalo especificado.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|Obtém uma linha contida no intervalo.|
||[Inserir (Shift: Excel. InsertShiftDirection)](/javascript/api/excel/excel.range#insert-shift-)|Insere uma célula ou um intervalo de células na planilha, no lugar desse intervalo, e desloca as outras células para liberar espaço.|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|Representa o código de formato de número do Excel para o intervalo especificado.|
||[address](/javascript/api/excel/excel.range#address)|Especifica a referência de intervalo no estilo a1.|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|Especifica a referência de intervalo para o intervalo especificado no idioma do usuário.|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|Especifica o número de células no intervalo.|
||[columnCount](/javascript/api/excel/excel.range#columncount)|Especifica o número total de colunas no intervalo.|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|Especifica o número de coluna da primeira célula do intervalo.|
||[format](/javascript/api/excel/excel.range#format)|Retorna um objeto de formato que encapsula a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades do intervalo.|
||[Validação](/javascript/api/excel/excel.range#rowcount)|Retorna o número total de linhas no intervalo.|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|Representa o número de linhas da primeira célula no intervalo.|
||[text](/javascript/api/excel/excel.range#text)|Valores de texto do intervalo especificado.|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|Especifica o tipo de dados em cada célula.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|A planilha que contém o intervalo atual.|
||[Seleciona.](/javascript/api/excel/excel.range#select--)|Seleciona o intervalo especificado na interface do usuário do Excel.|
||[values](/javascript/api/excel/excel.range#values)|Representa os valores brutos do intervalo especificado.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|Código de cor HTML que representa a cor da linha de borda do formulário #RRGGBB (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|Valor constante que indica o lado específico da borda.|
||[style](/javascript/api/excel/excel.rangeborder#style)|Uma das constantes de estilo de linha especificando o estilo de linha da borda.|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|Especifica o peso da borda em torno de um intervalo.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem (index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Obtém um objeto Border usando o respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|Obtém um objeto Border usando o respectivo índice.|
||[Count](/javascript/api/excel/excel.rangebordercollection#count)|Número de objetos de borda da coleção.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|Redefine a tela de fundo do intervalo.|
||[color](/javascript/api/excel/excel.rangefill#color)|Código de cor HTML que representa a cor do plano de fundo, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Representa o status da fonte em negrito.|
||[color](/javascript/api/excel/excel.rangefont#color)|Representação do código de cor HTML da cor do texto (por exemplo, #FF0000 representa vermelho).|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Especifica o status de itálico da fonte.|
||[name](/javascript/api/excel/excel.rangefont#name)|Nome da fonte (por exemplo, "Calibri").|
||[size](/javascript/api/excel/excel.rangefont#size)|Font Size|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Tipo de sublinhado aplicado à fonte.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|Representa o alinhamento horizontal do objeto especificado.|
||[Borders](/javascript/api/excel/excel.rangeformat#borders)|Coleção de objetos border que se aplicam a todo o intervalo.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Retorna o objeto de preenchimento definido em todo o intervalo.|
||[font](/javascript/api/excel/excel.rangeformat#font)|Retorna o objeto font definido em todo o intervalo.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|Representa o alinhamento vertical do objeto especificado.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|Especifica se o Excel quebra o texto no objeto.|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|Exclui a tabela.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getdatabodyrange--)|Obtém o objeto de intervalo associado ao corpo de dados da tabela.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getheaderrowrange--)|Obtém o objeto de intervalo associado à linha de cabeçalho da tabela.|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|Obtém o objeto de intervalo associado a toda a tabela.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#gettotalrowrange--)|Obtém o objeto de intervalo associado à linha de totais da tabela.|
||[name](/javascript/api/excel/excel.table#name)|Nome da tabela.|
||[colunas](/javascript/api/excel/excel.table#columns)|Representa uma coleção de todas as colunas na tabela.|
||[id](/javascript/api/excel/excel.table#id)|Retorna um valor que identifica de forma exclusiva a tabela em uma determinada pasta de trabalho.|
||[rows](/javascript/api/excel/excel.table#rows)|Representa uma coleção de todas as linhas na tabela.|
||[showHeaders](/javascript/api/excel/excel.table#showheaders)|Especifica se a linha de cabeçalho está visível.|
||[showTotals](/javascript/api/excel/excel.table#showtotals)|Especifica se a linha de total está visível.|
||[style](/javascript/api/excel/excel.table#style)|Valor da constante que representa o estilo de Tabela.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[Add (endereço: \| cadeia de caracteres de intervalo, hasHeaders: Boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|Crie uma nova tabela.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|Obtém uma tabela pelo nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|Obtém uma tabela com base em sua posição na coleção.|
||[Count](/javascript/api/excel/excel.tablecollection#count)|Retorna o número de tabelas na pasta de trabalho.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|Exclui a coluna da tabela.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|Obtém o objeto de intervalo associado ao corpo de dados da coluna.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|Obtém o objeto de intervalo associado à linha de cabeçalho da coluna.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|Obtém o objeto de intervalo associado a toda a coluna.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|Obtém o objeto de intervalo associado à linha de totais da coluna.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Especifica o nome da coluna da tabela.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Retorna uma chave exclusiva que identifica a coluna na tabela.|
||[índice](/javascript/api/excel/excel.tablecolumn#index)|Retorna o número de índice da coluna na coleção de colunas da tabela.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Representa os valores brutos do intervalo especificado.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[Add (index?: Number, Values?: matriz<matriz<\| número da cadeia de caracteres boolean \|>> \| número da \| cadeia de caracteres booleana \| , Name?: String)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|Adiciona uma nova coluna à tabela.|
||[getItem (Key: String de número \| )](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|Obtém um objeto de coluna por nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|Obtém uma coluna com base em sua posição na coleção.|
||[Count](/javascript/api/excel/excel.tablecolumncollection#count)|Retorna o número de colunas na tabela.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|Exclui a linha da tabela.|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|Retorna o objeto de intervalo associado a toda a linha.|
||[índice](/javascript/api/excel/excel.tablerow#index)|Retorna o número de índice da linha na coleção de linhas da tabela.|
||[values](/javascript/api/excel/excel.tablerow#values)|Representa os valores brutos do intervalo especificado.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[Add (index?: Number, Values?: matriz<matriz<\| número da cadeia de caracteres boolean \|>> \| número da cadeia de caracteres booleana \| \| )](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|Adiciona uma ou mais linhas à tabela.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|Obtém uma linha com base em sua posição na coleção.|
||[Count](/javascript/api/excel/excel.tablerowcollection#count)|Retorna o número de linhas na tabela.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange()](/javascript/api/excel/excel.workbook#getselectedrange--)|Obtém o intervalo único selecionado atualmente da pasta de trabalho.|
||[aplicativo](/javascript/api/excel/excel.workbook#application)|Representa a instância do aplicativo Excel que contém esta pasta de trabalho.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Representa uma coleção de ligações que fazem parte da pasta de trabalho.|
||[das](/javascript/api/excel/excel.workbook#names)|Representa uma coleção de itens denominados de escopo da pasta de trabalho (chamados intervalos e constantes).|
||[tabelas](/javascript/api/excel/excel.workbook#tables)|Representa uma coleção de tabelas associadas à pasta de trabalho.|
||[planilhas](/javascript/api/excel/excel.workbook#worksheets)|Representa uma coleção de planilhas associadas à pasta de trabalho.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Ative a planilha na interface do usuário do Excel.|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|Exclui a planilha da pasta de trabalho.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|Obtém o objeto de intervalo que contém a célula única com base nos números de linha e de coluna.|
||[GetRange (endereço?: cadeia de caracteres)](/javascript/api/excel/excel.worksheet#getrange-address-)|Obtém o objeto Range, representando um único bloco retangular de células, especificado pelo endereço ou nome.|
||[name](/javascript/api/excel/excel.worksheet#name)|O nome de exibição da planilha.|
||[position](/javascript/api/excel/excel.worksheet#position)|A posição baseada em zero da planilha na pasta de trabalho.|
||[gráficos](/javascript/api/excel/excel.worksheet#charts)|Retorna uma coleção de gráficos que fazem parte da planilha.|
||[id](/javascript/api/excel/excel.worksheet#id)|Retorna um valor que identifica de forma exclusiva a planilha em uma determinada pasta de trabalho.|
||[tabelas](/javascript/api/excel/excel.worksheet#tables)|Coleção de tabelas que fazem parte da planilha.|
||[visibility](/javascript/api/excel/excel.worksheet#visibility)|A visibilidade da planilha.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[Add (Name?: String)](/javascript/api/excel/excel.worksheetcollection#add-name-)|Adiciona uma nova planilha à pasta de trabalho.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|Obtém a planilha ativa no momento na pasta de trabalho.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|Obtém um objeto worksheet usando o Nome ou ID dele.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.1&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
