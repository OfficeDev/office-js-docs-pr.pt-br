---
title: Conjunto de requisitos de API JavaScript do Excel 1,1
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,1
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 90d7ee7cef2e8c48e458b2e14893ba9c13c68a30
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940784"
---
# <a name="excel-javascript-api-requirement-set-11"></a>Conjunto de requisitos de API JavaScript do Excel 1,1

A API JavaScript do Excel 1.1 é a primeira versão da API. É o único conjunto de requisitos específico do Excel suportado pelo Excel 2016.

## <a name="api-list"></a>Lista de APIs

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[Calculate (calculatype: Excel. Calculatype)](/javascript/api/excel/excel.application#calculate-calculationtype-)|Recalcula todas as pastas de trabalho abertas no Excel no momento.|
||[cálculomode](/javascript/api/excel/excel.application#calculationmode)|Retorna o modo de cálculo usado na pasta de trabalho, conforme definido pelas constantes no Excel. Calculation. Os valores possíveis são `Automatic`:, onde o Excel controla o recálculo; `AutomaticExceptTables`, onde o Excel controla o recálculo, mas ignora as alterações nas tabelas; `Manual`, onde o cálculo é feito quando o usuário solicita.|
|[Associação](/javascript/api/excel/excel.binding)|[getRange()](/javascript/api/excel/excel.binding#getrange--)|Retorna o intervalo representado pela associação. Gera um erro quando a associação não é do tipo correto.|
||[getTable()](/javascript/api/excel/excel.binding#gettable--)|Retorna a tabela representada pela associação. Gera um erro quando a associação não é do tipo correto.|
||[getText()](/javascript/api/excel/excel.binding#gettext--)|Retorna o texto representado pela associação. Gera um erro quando a associação não é do tipo correto.|
||[id](/javascript/api/excel/excel.binding#id)|Representa um identificador de associação. Somente leitura.|
||[tipo](/javascript/api/excel/excel.binding#type)|Retorna o tipo da associação. Consulte Excel. BindingType para obter detalhes. Somente leitura.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getItem(id: string)](/javascript/api/excel/excel.bindingcollection#getitem-id-)|Obtém um objeto de associação pela ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.bindingcollection#getitemat-index-)|Obtém um objeto de associação com base em sua posição na matriz dos itens.|
||[Count](/javascript/api/excel/excel.bindingcollection#count)|Retorna o número de associações da coleção. Somente leitura.|
||[items](/javascript/api/excel/excel.bindingcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Chart](/javascript/api/excel/excel.chart)|[delete()](/javascript/api/excel/excel.chart#delete--)|Exclui o objeto de gráfico.|
||[height](/javascript/api/excel/excel.chart#height)|Representa a altura, em pontos, do objeto Chart.|
||[left](/javascript/api/excel/excel.chart#left)|A distância, em pontos, da esquerda do gráfico à origem da planilha.|
||[name](/javascript/api/excel/excel.chart#name)|Representa o nome de um objeto Chart.|
||[Axes](/javascript/api/excel/excel.chart#axes)|Representa os eixos de um gráfico. Somente leitura.|
||[dataLabels](/javascript/api/excel/excel.chart#datalabels)|Representa os rótulos de dados no gráfico. Somente leitura.|
||[format](/javascript/api/excel/excel.chart#format)|Encapsula as propriedades de formato da área do gráfico. Somente leitura.|
||[Legenda](/javascript/api/excel/excel.chart#legend)|Representa a legenda do gráfico. Somente leitura.|
||[series](/javascript/api/excel/excel.chart#series)|Representa uma única série ou uma coleção de séries no gráfico. Somente leitura.|
||[title](/javascript/api/excel/excel.chart#title)|Representa o título do gráfico especificado, incluindo o respectivo texto, a visibilidade, a posição e a formatação. Somente leitura.|
||[setData (sourceData: Range, Seriesby como?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chart#setdata-sourcedata--seriesby-)|Redefine os dados de origem do gráfico.|
||[SETPOSITION (startCell: String \| de intervalo, endcell?: \| cadeia de caracteres de intervalo)](/javascript/api/excel/excel.chart#setposition-startcell--endcell-)|Posiciona o gráfico em relação às células na planilha.|
||[top](/javascript/api/excel/excel.chart#top)|Representa a distância, em pontos, da borda superior do objeto à parte superior da primeira linha de uma planilha ou da área de um gráfico.|
||[width](/javascript/api/excel/excel.chart#width)|Representa a largura, em pontos, do objeto de gráfico.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[fill](/javascript/api/excel/excel.chartareaformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo. Somente leitura.|
||[font](/javascript/api/excel/excel.chartareaformat#font)|Representa os atributos de fonte do objeto atual, como nome, tamanho, cor, dentre outros. Somente leitura.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[categoryAxis](/javascript/api/excel/excel.chartaxes#categoryaxis)|Representa o eixo de categoria em um gráfico. Somente leitura.|
||[seriesAxis](/javascript/api/excel/excel.chartaxes#seriesaxis)|Representa o eixo das séries de um gráfico 3D. Somente leitura.|
||[valueAxis](/javascript/api/excel/excel.chartaxes#valueaxis)|Representa o eixo dos valores em um eixo. Somente leitura.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[majorUnit](/javascript/api/excel/excel.chartaxis#majorunit)|Representa o intervalo entre as duas principais marcas de escala. Pode ser definido como um valor numérico ou uma cadeia de caracteres vazia.  O valor retornado sempre é um número.|
||[maximum](/javascript/api/excel/excel.chartaxis#maximum)|Representa o valor máximo no eixo dos valores.  Pode ser definido como um valor numérico ou uma cadeia de caracteres vazia (para valores automáticos de eixo).  O valor retornado sempre é um número.|
||[minimum](/javascript/api/excel/excel.chartaxis#minimum)|Representa o valor mínimo no eixo dos valores. Pode ser definido como um valor numérico ou uma cadeia de caracteres vazia (para valores automáticos de eixo).  O valor retornado sempre é um número.|
||[minorUnit](/javascript/api/excel/excel.chartaxis#minorunit)|Representa o intervalo entre as duas marcas de escala secundárias. Pode ser definido como um valor numérico ou uma cadeia de caracteres vazia (para valores automáticos de eixo). O valor retornado sempre é um número.|
||[format](/javascript/api/excel/excel.chartaxis#format)|Representa a formatação de um objeto Chart, que inclui formatação de linha e de fonte. Somente leitura.|
||[majorGridlines](/javascript/api/excel/excel.chartaxis#majorgridlines)|Retorna um objeto Gridlines que representa as principais linhas de grade do eixo especificado. Somente leitura.|
||[minorGridlines](/javascript/api/excel/excel.chartaxis#minorgridlines)|Retorna um objeto Gridlines que representa as linhas de grade secundárias do eixo especificado. Somente leitura.|
||[title](/javascript/api/excel/excel.chartaxis#title)|Representa o título do eixo. Somente leitura.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[font](/javascript/api/excel/excel.chartaxisformat#font)|Representa os atributos de fonte de um elemento do eixo do gráfico, como nome, tamanho, cor, etc. Somente leitura.|
||[line](/javascript/api/excel/excel.chartaxisformat#line)|Representa a formatação de linha do gráfico. Somente leitura.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[format](/javascript/api/excel/excel.chartaxistitle#format)|Representa a formatação do título do eixo do gráfico. Somente leitura.|
||[text](/javascript/api/excel/excel.chartaxistitle#text)|Representa o título do eixo.|
||[visible](/javascript/api/excel/excel.chartaxistitle#visible)|Um booliano que especifica a visibilidade de um título do eixo.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[font](/javascript/api/excel/excel.chartaxistitleformat#font)|Representa os atributos de fonte, como nome, tamanho, cor, etc., do objeto do eixo do gráfico. Somente leitura.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[Add (tipo: Excel. ChartType, sourceData: Range, Seriesby como?: Excel. ChartSeriesBy)](/javascript/api/excel/excel.chartcollection#add-type--sourcedata--seriesby-)|Cria um novo gráfico.|
||[getItem(name: string)](/javascript/api/excel/excel.chartcollection#getitem-name-)|Obtém um gráfico usando o respectivo nome. Quando houver vários gráficos com o mesmo nome, o sistema retornará o primeiro deles.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartcollection#getitemat-index-)|Obtém um gráfico com base em sua posição no conjunto.|
||[Count](/javascript/api/excel/excel.chartcollection#count)|Retorna o número de gráficos da planilha. Somente leitura.|
||[items](/javascript/api/excel/excel.chartcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[fill](/javascript/api/excel/excel.chartdatalabelformat#fill)|Representa o formato de preenchimento do rótulo de dados atual do gráfico. Somente leitura.|
||[font](/javascript/api/excel/excel.chartdatalabelformat#font)|Representa os atributos de fonte do rótulo de dados do gráfico, como nome, tamanho, cor, dentre outros. Somente leitura.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[position](/javascript/api/excel/excel.chartdatalabels#position)|Valor de DataLabelPosition que representa a posição do rótulo de dados. Consulte Excel. ChartDataLabelPosition para obter detalhes.|
||[format](/javascript/api/excel/excel.chartdatalabels#format)|Representa o formato dos rótulos de dados do gráfico, que inclui a formatação de fonte e de preenchimento. Somente leitura.|
||[divisória](/javascript/api/excel/excel.chartdatalabels#separator)|Cadeia de caracteres que representa o separador usado para os rótulos de dados em um gráfico.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabels#showbubblesize)|Valor booliano que determina se o tamanho da bolha do rótulo de dados fica visível ou não.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabels#showcategoryname)|Valor booliano que determina se o nome da categoria do rótulo de dados fica visível ou não.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabels#showlegendkey)|Valor booliano que determina se o código de legenda do rótulo de dados fica visível ou não.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabels#showpercentage)|Valor booliano que determina se o percentual do rótulo de dados fica visível ou não.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabels#showseriesname)|Valor booliano que determina se o nome da série do rótulo de dados fica visível ou não.|
||[showValue](/javascript/api/excel/excel.chartdatalabels#showvalue)|Valor booliano que determina se o valor do rótulo de dados fica visível ou não.|
|[ChartFill](/javascript/api/excel/excel.chartfill)|[clear()](/javascript/api/excel/excel.chartfill#clear--)|Limpa a cor de preenchimento de um elemento do gráfico.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.chartfill#setsolidcolor-color-)|Define a formatação de preenchimento de um elemento do gráfico com uma cor uniforme.|
|[ChartFont](/javascript/api/excel/excel.chartfont)|[bold](/javascript/api/excel/excel.chartfont#bold)|Representa o status da fonte em negrito.|
||[color](/javascript/api/excel/excel.chartfont#color)|Representação de código de cor HTML para a cor do texto. Por exemplo #FF0000 representa vermelho.|
||[italic](/javascript/api/excel/excel.chartfont#italic)|Representa o status da fonte em itálico.|
||[name](/javascript/api/excel/excel.chartfont#name)|Nome da fonte (por exemplo, "Calibri")|
||[size](/javascript/api/excel/excel.chartfont#size)|Tamanho da fonte, por exemplo, 11.|
||[underline](/javascript/api/excel/excel.chartfont#underline)|Tipo de sublinhado aplicado à fonte. Consulte Excel. ChartUnderlineStyle para obter detalhes.|
|[ChartGridlines](/javascript/api/excel/excel.chartgridlines)|[format](/javascript/api/excel/excel.chartgridlines#format)|Representa a formatação de linhas de grade do gráfico. Somente leitura.|
||[visible](/javascript/api/excel/excel.chartgridlines#visible)|Valor booleano que determina quando as linhas de grade do eixo ficam visível ou não.|
|[ChartGridlinesFormat](/javascript/api/excel/excel.chartgridlinesformat)|[line](/javascript/api/excel/excel.chartgridlinesformat#line)|Representa a formatação de linha do gráfico. Somente leitura.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[overlay](/javascript/api/excel/excel.chartlegend#overlay)|Valor booleano para determinar quando a legenda do gráfico deve se sobrepor ao corpo principal do gráfico.|
||[position](/javascript/api/excel/excel.chartlegend#position)|Representa a posição da legenda no gráfico. Consulte Excel. ChartLegendPosition para obter detalhes.|
||[format](/javascript/api/excel/excel.chartlegend#format)|Representa a formatação de uma legenda de gráfico, que inclui a formatação de fonte e de preenchimento. Somente leitura.|
||[visible](/javascript/api/excel/excel.chartlegend#visible)|Um valor booliano que representa a visibilidade de um objeto ChartLegend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[fill](/javascript/api/excel/excel.chartlegendformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo. Somente leitura.|
||[font](/javascript/api/excel/excel.chartlegendformat#font)|Representa os atributos de fonte, como nome da fonte, tamanho da fonte, cor, etc. de uma legenda do gráfico. Somente leitura.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[clear()](/javascript/api/excel/excel.chartlineformat#clear--)|Limpar o formato da linha de um elemento do gráfico.|
||[color](/javascript/api/excel/excel.chartlineformat#color)|Código de cores HTML que representa a cor das linhas no gráfico.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[format](/javascript/api/excel/excel.chartpoint#format)|Encapsula as propriedades de formato de um ponto do gráfico. Somente leitura.|
||[value](/javascript/api/excel/excel.chartpoint#value)|Retorna o valor de um ponto do gráfico. Somente leitura.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[fill](/javascript/api/excel/excel.chartpointformat#fill)|Representa o formato de preenchimento de um gráfico, que inclui informações de formatação de plano de fundo. Somente leitura.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartpointscollection#getitemat-index-)|Recupera um ponto com base na respectiva posição dentro da série.|
||[Count](/javascript/api/excel/excel.chartpointscollection#count)|Retorna o número de pontos do gráfico da série. Somente leitura.|
||[items](/javascript/api/excel/excel.chartpointscollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[name](/javascript/api/excel/excel.chartseries#name)|Representa o nome de uma série do gráfico.|
||[format](/javascript/api/excel/excel.chartseries#format)|Representa a formatação de uma série do gráfico, que inclui a formatação de linha e de preenchimento. Somente leitura.|
||[pontos](/javascript/api/excel/excel.chartseries#points)|Representa uma coleção de todos os pontos da série. Somente leitura.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getItemAt(index: number)](/javascript/api/excel/excel.chartseriescollection#getitemat-index-)|Recupera uma série com base na respectiva posição na coleção.|
||[Count](/javascript/api/excel/excel.chartseriescollection#count)|Retorna o número de série da coleção. Somente leitura.|
||[items](/javascript/api/excel/excel.chartseriescollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartSeriesFormat](/javascript/api/excel/excel.chartseriesformat)|[fill](/javascript/api/excel/excel.chartseriesformat#fill)|Representa o formato de preenchimento de uma série do gráfico, que inclui informações sobre a formatação da tela de fundo. Somente leitura.|
||[line](/javascript/api/excel/excel.chartseriesformat#line)|Representa a formatação de linha. Somente leitura.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[overlay](/javascript/api/excel/excel.charttitle#overlay)|Valor booleano que determina quando o título do gráfico deve se sobrepor ao gráfico ou não.|
||[format](/javascript/api/excel/excel.charttitle#format)|Representa a formatação de um título do gráfico, que inclui a formatação de fonte e de preenchimento. Somente leitura.|
||[text](/javascript/api/excel/excel.charttitle#text)|Representa o texto do título de um gráfico.|
||[visible](/javascript/api/excel/excel.charttitle#visible)|Um valor booliano que representa a visibilidade de um objeto de título de gráfico.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[fill](/javascript/api/excel/excel.charttitleformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo. Somente leitura.|
||[font](/javascript/api/excel/excel.charttitleformat#font)|Representa os atributos de fonte de um objeto, como nome, tamanho, cor, dentre outros. Somente leitura.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[getRange()](/javascript/api/excel/excel.nameditem#getrange--)|Retorna o objeto Range associado ao nome. Gerará um erro se o tipo do item nomeado não for um intervalo.|
||[name](/javascript/api/excel/excel.nameditem#name)|O nome do objeto. Somente leitura.|
||[tipo](/javascript/api/excel/excel.nameditem#type)|Indica o tipo do valor retornado pela fórmula do nome. Consulte Excel. NamedItemType para obter detalhes. Somente leitura.|
||[value](/javascript/api/excel/excel.nameditem#value)|Representa o valor calculado pela fórmula do nome. Para um intervalo nomeado, retornará o endereço do intervalo. Somente leitura.|
||[visible](/javascript/api/excel/excel.nameditem#visible)|Determina se o objeto estará visível ou não.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[getItem(name: string)](/javascript/api/excel/excel.nameditemcollection#getitem-name-)|Obtém um objeto NamedItem usando seu nome.|
||[items](/javascript/api/excel/excel.nameditemcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Range](/javascript/api/excel/excel.range)|[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.range#clear-applyto-)|Limpe valores de intervalo, formatação, preenchimento, bordas, etc.|
||[excluir (Shift: Excel. DeleteShiftDirection)](/javascript/api/excel/excel.range#delete-shift-)|Exclui as células associadas ao intervalo.|
||[fórmulas](/javascript/api/excel/excel.range#formulas)|Representa a fórmula em notação A1.|
||[formulasLocal](/javascript/api/excel/excel.range#formulaslocal)|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.  Por exemplo, a fórmula "=SUM(A1, 1.5)" em inglês seria "=SOMA(A1; 1,5)" em português.|
||[getBoundingRect (anotherRange: cadeia \| de caracteres de intervalo)](/javascript/api/excel/excel.range#getboundingrect-anotherrange-)|Obtém o menor objeto Range que engloba os intervalos determinados. Por exemplo, o GetBoundingRect de "B2:C5" e "D10:E15" é "B2:E15".|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.range#getcell-row--column-)|Obtém o objeto de intervalo que contém a célula única com base nos números de linha e de coluna. A célula pode estar fora dos limites de seu intervalo pai, desde que ela permaneça dentro da grade da planilha. A localização da célula retornada está relacionada à célula superior esquerda do intervalo.|
||[getColumn(column: number)](/javascript/api/excel/excel.range#getcolumn-column-)|Obtém uma coluna incluída no intervalo.|
||[getEntireColumn()](/javascript/api/excel/excel.range#getentirecolumn--)|Obtém um objeto que representa a coluna inteira do intervalo (por exemplo, se o intervalo atual representa as células "B4: E11", seu `getEntireColumn` é um intervalo que representa as colunas "B:E").|
||[getEntireRow()](/javascript/api/excel/excel.range#getentirerow--)|Obtém um objeto que representa a linha inteira do intervalo (por exemplo, se o intervalo atual representa as células "B4: E11", seu `GetEntireRow` é um intervalo que representa as linhas "4:11").|
||[getintersection (anotherRange: cadeia \| de caracteres de intervalo)](/javascript/api/excel/excel.range#getintersection-anotherrange-)|Obtém o objeto Range que representa a interseção retangular dos intervalos determinados.|
||[getLastCell()](/javascript/api/excel/excel.range#getlastcell--)|Obtém a última célula do intervalo. Por exemplo, a última célula de "B2:D5" é "D5".|
||[getLastColumn()](/javascript/api/excel/excel.range#getlastcolumn--)|Obtém a última coluna do intervalo. Por exemplo, a última coluna de "B2:D5" é "D2:D5".|
||[getLastRow()](/javascript/api/excel/excel.range#getlastrow--)|Obtém a última linha do intervalo. Por exemplo, a última linha de "B2:D5" é "B5:D5".|
||[getOffsetRange(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-)|Obtém um objeto que representa um intervalo deslocado do intervalo especificado. A dimensão do intervalo retornado corresponde a esse intervalo. Se o intervalo resultante for imposto para fora dos limites da grade da planilha, o sistema gerará um erro.|
||[getRow(row: number)](/javascript/api/excel/excel.range#getrow-row-)|Obtém uma linha contida no intervalo.|
||[Inserir (Shift: Excel. InsertShiftDirection)](/javascript/api/excel/excel.range#insert-shift-)|Insere uma célula ou um intervalo de células na planilha, no lugar desse intervalo, e desloca as outras células para liberar espaço. Retorna um novo objeto Range no espaço em branco atual.|
||[numberFormat](/javascript/api/excel/excel.range#numberformat)|Representa o código de formato de número do Excel para o intervalo especificado.|
||[address](/javascript/api/excel/excel.range#address)|Representa a referência do intervalo no estilo A1. O valor de endereço conterá a referência de planilha (por exemplo, "Planilha1! A1: B4 "). Somente leitura.|
||[addressLocal](/javascript/api/excel/excel.range#addresslocal)|Representa a referência de intervalo para o intervalo especificado no idioma do usuário. Somente leitura.|
||[cellCount](/javascript/api/excel/excel.range#cellcount)|Número de células no intervalo. Essa API retornará -1 se a contagem de células exceder 2^31-1 (2.147.483.647). Somente leitura.|
||[columnCount](/javascript/api/excel/excel.range#columncount)|Representa o número total de colunas no intervalo. Somente leitura.|
||[columnIndex](/javascript/api/excel/excel.range#columnindex)|Representa o número de colunas da primeira célula no intervalo. Indexados com zero. Somente leitura.|
||[format](/javascript/api/excel/excel.range#format)|Retorna um objeto de formato que encapsula a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades do intervalo. Somente leitura.|
||[Validação](/javascript/api/excel/excel.range#rowcount)|Retorna o número total de linhas no intervalo. Somente leitura.|
||[rowIndex](/javascript/api/excel/excel.range#rowindex)|Representa o número de linhas da primeira célula no intervalo. Indexados com zero. Somente leitura.|
||[text](/javascript/api/excel/excel.range#text)|Valores de texto do intervalo especificado. O valor de texto não depende da largura da célula. A substituição pelo sinal #, que ocorre na interface de usuário do Excel, não afeta o valor de texto retornado pela API. Somente leitura.|
||[valueTypes](/javascript/api/excel/excel.range#valuetypes)|Representa o tipo de dados de cada célula. Somente leitura.|
||[worksheet](/javascript/api/excel/excel.range#worksheet)|A planilha que contém o intervalo atual. Somente leitura.|
||[Seleciona.](/javascript/api/excel/excel.range#select--)|Seleciona o intervalo especificado na interface do usuário do Excel.|
||[values](/javascript/api/excel/excel.range#values)|Representa os valores brutos do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[color](/javascript/api/excel/excel.rangeborder#color)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
||[sideIndex](/javascript/api/excel/excel.rangeborder#sideindex)|Valor constante que indica o lado específico da borda. Consulte Excel. BorderIndex para obter detalhes. Somente leitura.|
||[style](/javascript/api/excel/excel.rangeborder#style)|Uma das constantes de estilo de linha especificando o estilo de linha da borda. Consulte Excel. BorderLineStyle para obter detalhes.|
||[weight](/javascript/api/excel/excel.rangeborder#weight)|Especifica o peso da borda em torno de um intervalo. Consulte Excel. BorderWeight para obter detalhes.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[getItem (index: Excel. BorderIndex)](/javascript/api/excel/excel.rangebordercollection#getitem-index-)|Obtém um objeto Border usando o respectivo nome.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangebordercollection#getitemat-index-)|Obtém um objeto Border usando o respectivo índice.|
||[Count](/javascript/api/excel/excel.rangebordercollection#count)|Número de objetos de borda da coleção. Somente leitura.|
||[items](/javascript/api/excel/excel.rangebordercollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[clear()](/javascript/api/excel/excel.rangefill#clear--)|Redefine a tela de fundo do intervalo.|
||[color](/javascript/api/excel/excel.rangefill#color)|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[bold](/javascript/api/excel/excel.rangefont#bold)|Representa o status da fonte em negrito.|
||[color](/javascript/api/excel/excel.rangefont#color)|Representação de código de cor HTML para a cor do texto. Por exemplo #FF0000 representa vermelho.|
||[italic](/javascript/api/excel/excel.rangefont#italic)|Representa o status da fonte em itálico.|
||[name](/javascript/api/excel/excel.rangefont#name)|Nome da fonte (por exemplo, "Calibri")|
||[size](/javascript/api/excel/excel.rangefont#size)|Font Size|
||[underline](/javascript/api/excel/excel.rangefont#underline)|Tipo de sublinhado aplicado à fonte. Consulte Excel. RangeUnderlineStyle para obter detalhes.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[horizontalAlignment](/javascript/api/excel/excel.rangeformat#horizontalalignment)|Representa o alinhamento horizontal do objeto especificado. Consulte Excel. HorizontalAlignment para obter detalhes.|
||[Borders](/javascript/api/excel/excel.rangeformat#borders)|Coleção de objetos border que se aplicam a todo o intervalo. Somente leitura.|
||[fill](/javascript/api/excel/excel.rangeformat#fill)|Retorna o objeto de preenchimento definido em todo o intervalo. Somente leitura.|
||[font](/javascript/api/excel/excel.rangeformat#font)|Retorna o objeto font definido em todo o intervalo. Somente leitura.|
||[verticalAlignment](/javascript/api/excel/excel.rangeformat#verticalalignment)|Representa o alinhamento vertical do objeto especificado. Consulte Excel. VerticalAlignment para obter detalhes.|
||[wrapText](/javascript/api/excel/excel.rangeformat#wraptext)|Indica se o Excel quebra automaticamente a linha de texto no objeto. Um valor nulo indica que o intervalo inteiro não tem configuração de quebra de linha automática uniforme.|
|[Table](/javascript/api/excel/excel.table)|[delete()](/javascript/api/excel/excel.table#delete--)|Exclui a tabela.|
||[getDataBodyRange()](/javascript/api/excel/excel.table#getdatabodyrange--)|Obtém o objeto de intervalo associado ao corpo de dados da tabela.|
||[getHeaderRowRange()](/javascript/api/excel/excel.table#getheaderrowrange--)|Obtém o objeto de intervalo associado à linha de cabeçalho da tabela.|
||[getRange()](/javascript/api/excel/excel.table#getrange--)|Obtém o objeto de intervalo associado a toda a tabela.|
||[getTotalRowRange()](/javascript/api/excel/excel.table#gettotalrowrange--)|Obtém o objeto de intervalo associado à linha de totais da tabela.|
||[name](/javascript/api/excel/excel.table#name)|Nome da tabela.|
||[colunas](/javascript/api/excel/excel.table#columns)|Representa uma coleção de todas as colunas na tabela. Somente leitura.|
||[id](/javascript/api/excel/excel.table#id)|Retorna um valor que identifica de forma exclusiva a tabela em uma determinada pasta de trabalho. O valor do identificador permanece o mesmo, ainda que a tabela seja renomeada. Somente leitura.|
||[rows](/javascript/api/excel/excel.table#rows)|Representa uma coleção de todas as linhas na tabela. Somente leitura.|
||[showHeaders](/javascript/api/excel/excel.table#showheaders)|Indica se a linha do cabeçalho está visível ou não. Esse valor pode ser definido para mostrar ou remover a linha do cabeçalho.|
||[showTotals](/javascript/api/excel/excel.table#showtotals)|Indica se a linha do total está visível ou não. Esse valor pode ser definido para mostrar ou remover a linha do total.|
||[style](/javascript/api/excel/excel.table#style)|Valor da constante que representa o estilo de Tabela. Os valores possíveis são: TableStyleLight1 a TableStyleLight21, TableStyleMedium1 a TableStyleMedium28, TableStyleStyleDark1 a TableStyleStyleDark11. Também é possível usar um estilo definido pelo usuário que esteja presente na planilha.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[Add (endereço: cadeia \| de caracteres de intervalo, hasHeaders: Boolean)](/javascript/api/excel/excel.tablecollection#add-address--hasheaders-)|Cria uma nova tabela. O objeto de intervalo ou endereço de origem determina a planilha à qual a tabela será adicionada. Se a tabela não puder ser adicionada (por exemplo, porque o endereço é inválido ou a tabela se sobreporia a outra), será gerado um erro.|
||[getItem(key: string)](/javascript/api/excel/excel.tablecollection#getitem-key-)|Obtém uma tabela pelo nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecollection#getitemat-index-)|Obtém uma tabela com base em sua posição na coleção.|
||[Count](/javascript/api/excel/excel.tablecollection#count)|Retorna o número de tabelas na pasta de trabalho. Somente leitura.|
||[items](/javascript/api/excel/excel.tablecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableColumn](/javascript/api/excel/excel.tablecolumn)|[delete()](/javascript/api/excel/excel.tablecolumn#delete--)|Exclui a coluna da tabela.|
||[getDataBodyRange()](/javascript/api/excel/excel.tablecolumn#getdatabodyrange--)|Obtém o objeto de intervalo associado ao corpo de dados da coluna.|
||[getHeaderRowRange()](/javascript/api/excel/excel.tablecolumn#getheaderrowrange--)|Obtém o objeto de intervalo associado à linha de cabeçalho da coluna.|
||[getRange()](/javascript/api/excel/excel.tablecolumn#getrange--)|Obtém o objeto de intervalo associado a toda a coluna.|
||[getTotalRowRange()](/javascript/api/excel/excel.tablecolumn#gettotalrowrange--)|Obtém o objeto de intervalo associado à linha de totais da coluna.|
||[name](/javascript/api/excel/excel.tablecolumn#name)|Representa o nome da coluna da tabela.|
||[id](/javascript/api/excel/excel.tablecolumn#id)|Retorna uma chave exclusiva que identifica a coluna na tabela. Somente leitura.|
||[index](/javascript/api/excel/excel.tablecolumn#index)|Retorna o número de índice da coluna na coleção de colunas da tabela. Indexado com zero. Somente leitura.|
||[values](/javascript/api/excel/excel.tablecolumn#values)|Representa os valores brutos do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[Add (index?: Number, Values?: matriz<matriz<\| número \| da cadeia \| de \| caracteres \| Boolean>> número da cadeia de caracteres booleana, Name?: String)](/javascript/api/excel/excel.tablecolumncollection#add-index--values--name-)|Adiciona uma nova coluna à tabela.|
||[getItem (Key: String \| de número)](/javascript/api/excel/excel.tablecolumncollection#getitem-key-)|Obtém um objeto de coluna por nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablecolumncollection#getitemat-index-)|Obtém uma coluna com base em sua posição na coleção.|
||[Count](/javascript/api/excel/excel.tablecolumncollection#count)|Retorna o número de colunas na tabela. Somente leitura.|
||[items](/javascript/api/excel/excel.tablecolumncollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableRow](/javascript/api/excel/excel.tablerow)|[delete()](/javascript/api/excel/excel.tablerow#delete--)|Exclui a linha da tabela.|
||[getRange()](/javascript/api/excel/excel.tablerow#getrange--)|Retorna o objeto de intervalo associado a toda a linha.|
||[index](/javascript/api/excel/excel.tablerow#index)|Retorna o número de índice da linha na coleção de linhas da tabela. Indexados com zero. Somente leitura.|
||[values](/javascript/api/excel/excel.tablerow#values)|Representa os valores brutos do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[Add (index?: Number, Values?: matriz<matriz<\| número \| da cadeia \| de \| caracteres \| Boolean>> número da cadeia de caracteres booleana)](/javascript/api/excel/excel.tablerowcollection#add-index--values-)|Adiciona uma ou mais linhas à tabela. O objeto de retorno será a parte superior das linhas adicionadas recentemente.|
||[getItemAt(index: number)](/javascript/api/excel/excel.tablerowcollection#getitemat-index-)|Obtém uma linha com base em sua posição na coleção.|
||[Count](/javascript/api/excel/excel.tablerowcollection#count)|Retorna o número de linhas na tabela. Somente leitura.|
||[items](/javascript/api/excel/excel.tablerowcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getSelectedRange()](/javascript/api/excel/excel.workbook#getselectedrange--)|Obtém o intervalo único selecionado atualmente da pasta de trabalho. Se houver vários intervalos selecionados, este método gerará um erro.|
||[application](/javascript/api/excel/excel.workbook#application)|Representa a instância do aplicativo Excel que contém esta pasta de trabalho. Somente leitura.|
||[bindings](/javascript/api/excel/excel.workbook#bindings)|Representa uma coleção de ligações que fazem parte da pasta de trabalho. Somente leitura.|
||[names](/javascript/api/excel/excel.workbook#names)|Representa uma coleção de itens denominados de escopo da pasta de trabalho (chamados intervalos e constantes). Somente leitura.|
||[tables](/javascript/api/excel/excel.workbook#tables)|Representa uma coleção de tabelas associadas à pasta de trabalho. Somente leitura.|
||[planilhas](/javascript/api/excel/excel.workbook#worksheets)|Representa uma coleção de planilhas associadas à pasta de trabalho. Somente leitura.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[activate()](/javascript/api/excel/excel.worksheet#activate--)|Ative a planilha na interface do usuário do Excel.|
||[delete()](/javascript/api/excel/excel.worksheet#delete--)|Exclui a planilha da pasta de trabalho. Observe que, se a visibilidade da planilha estiver definida como "VeryHidden", a operação de exclusão falhará com uma Generalexception.|
||[getCell(row: number, column: number)](/javascript/api/excel/excel.worksheet#getcell-row--column-)|Obtém o objeto de intervalo que contém a célula única com base nos números de linha e de coluna. A célula pode estar fora dos limites de seu intervalo pai, desde que ela permaneça dentro da grade da planilha.|
||[GetRange (endereço?: cadeia de caracteres)](/javascript/api/excel/excel.worksheet#getrange-address-)|Obtém o objeto Range, representando um único bloco retangular de células, especificado pelo endereço ou nome.|
||[name](/javascript/api/excel/excel.worksheet#name)|O nome de exibição da planilha.|
||[position](/javascript/api/excel/excel.worksheet#position)|A posição baseada em zero da planilha na pasta de trabalho.|
||[charts](/javascript/api/excel/excel.worksheet#charts)|Retorna uma coleção de gráficos que fazem parte da planilha. Somente leitura.|
||[id](/javascript/api/excel/excel.worksheet#id)|Retorna um valor que identifica de forma exclusiva a planilha em uma determinada pasta de trabalho. O valor do identificador permanece o mesmo, ainda que a planilha seja renomeada ou movida. Somente leitura.|
||[tables](/javascript/api/excel/excel.worksheet#tables)|Coleção de tabelas que fazem parte da planilha. Somente leitura.|
||[visibilidade](/javascript/api/excel/excel.worksheet#visibility)|A visibilidade da planilha.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[Add (Name?: String)](/javascript/api/excel/excel.worksheetcollection#add-name-)|Adiciona uma nova planilha à pasta de trabalho. A planilha será adicionada ao final das planilhas existentes. Se você quiser ativar a planilha recém-adicionada, chame “.activate()” nela.|
||[getActiveWorksheet()](/javascript/api/excel/excel.worksheetcollection#getactiveworksheet--)|Obtém a planilha ativa no momento na pasta de trabalho.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcollection#getitem-key-)|Obtém um objeto worksheet usando o Nome ou ID dele.|
||[items](/javascript/api/excel/excel.worksheetcollection#items)|Obtém os itens filhos carregados nesta coleção.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do Excel](/javascript/api/excel)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
