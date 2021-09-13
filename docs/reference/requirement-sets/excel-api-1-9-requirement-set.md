---
title: Excel Conjunto de requisitos da API JavaScript 1.9
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.9.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: dde36db799a7f0612439e934d50af4f3ab04077e
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151719"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>Novidades na API JavaScript 1.9 Excel JavaScript

Mais de 500 novas APIs do Excel foram introduzidas com o conjunto de requisitos 1.9. A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Shapes](../../excel/excel-add-ins-shapes.md) | Inserir, posicionar e formatar imagens, formas geométricas e caixas de texto. | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [Filtro Automático](../../excel/excel-add-ins-worksheets.md#filter-data) | Adicionar filtros aos intervalos. | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Areas](../../excel/excel-add-ins-multiple-ranges.md) | Suporte para intervalos descontínuos. | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [Células Especiais](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | Obtenha células que contêm datas, comentários ou fórmulas dentro de um intervalo. | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [Find](../../excel/excel-add-ins-ranges-string-match.md) | Encontre valores ou fórmulas em uma planilha ou intervalo. | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [Copiar e colar](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | Copie fórmulas, formatos e valores de um intervalo para outro. | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Calculation](../../excel/performance.md#suspend-calculation-temporarily) | Maior controle sobre o mecanismo de cálculo do Excel. | [Aplicativo](/javascript/api/excel/excel.application) |
| Novos Gráficos | Explore os novos tipos de gráficos compatíveis: mapas, caixa estreita, cascata, explosão solar, pareto. e funil. | [Chart](/javascript/api/excel/excel.charttype) |
| RangeFormat | Novos recursos com formatos de intervalo. | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.9. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.9 ou anterior, consulte Excel APIs no conjunto de requisitos [1.9](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)ou anterior .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationEngineVersion)|Retorna a versão do mecanismo de cálculo do Excel usada para o último recálculo completo.|
||[calculationState](/javascript/api/excel/excel.application#calculationState)|Retorna o estado de cálculo do aplicativo.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativeCalculation)|Retorna as configurações de cálculo iterativo.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendScreenUpdatingUntilNextSync__)|Suspende a atualização de tela até que a próxima `context.sync()` seja chamada.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply_range__columnIndex__criteria_)|Aplica o AutoFiltro a um intervalo.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearCriteria__)|Limpa os critérios de filtro do AutoFiltro.|
||[getRange()](/javascript/api/excel/excel.autofilter#getRange__)|Retorna o `Range` objeto que representa o intervalo ao qual o AutoFilter se aplica.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getRangeOrNullObject__)|Retorna o `Range` objeto que representa o intervalo ao qual o AutoFilter se aplica.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Uma matriz que contém todos os critérios de filtro no intervalo de autofiltro.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Especifica se o AutoFilter está habilitado.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isDataFiltered)|Especifica se o Filtro Automático tem critérios de filtro.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply__)|Aplica o objeto Autofilter especificado que está atualmente no intervalo.|
||[remove()](/javascript/api/excel/excel.autofilter#remove__)|Remove o Filtro automático do intervalo.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|Representa a propriedade `color` de uma única borda.|
||[style](/javascript/api/excel/excel.cellborder#style)|Representa a propriedade `style` de uma única borda.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintAndShade)|Representa a propriedade `tintAndShade` de uma única borda.|
||[weight](/javascript/api/excel/excel.cellborder#weight)|Representa a propriedade `weight` de uma única borda.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|Representa a propriedade `format.borders.bottom`.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonalDown)|Representa a propriedade `format.borders.diagonalDown`.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalUp)|Representa a propriedade `format.borders.diagonalUp`.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|Representa a propriedade `format.borders.horizontal`.|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|Representa a propriedade `format.borders.left`.|
||[direita](/javascript/api/excel/excel.cellbordercollection#right)|Representa a propriedade `format.borders.right`.|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|Representa a propriedade `format.borders.top`.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|Representa a propriedade `format.borders.vertical`.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|Representa a propriedade `address`.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addressLocal)|Representa a propriedade `addressLocal`.|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|Representa a propriedade `hidden`.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Representa a propriedade `format.fill.color`.|
||[padrão](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Representa a propriedade `format.fill.pattern`.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patternColor)|Representa a propriedade `format.fill.patternColor`.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patternTintAndShade)|Representa a propriedade `format.fill.patternTintAndShade`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintAndShade)|Representa a propriedade `format.fill.tintAndShade`.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|Representa a propriedade `format.font.bold`.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|Representa a propriedade `format.font.color`.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|Representa a propriedade `format.font.italic`.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|Representa a propriedade`format.font.name`.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|Representa a propriedade`format.font.size`.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|Representa a propriedade `format.font.strikethrough`.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|Representa a propriedade `format.font.subscript`.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|Representa a propriedade `format.font.superscript`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintAndShade)|Representa a propriedade `format.font.tintAndShade`.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|Representa a propriedade `format.font.underline`.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoIndent)|Representa a propriedade `autoIndent`.|
||[Borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|Representa a propriedade `borders`.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|Representa a propriedade `fill`.|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|Representa a propriedade `font`.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalAlignment)|Representa a propriedade `horizontalAlignment`.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentLevel)|Representa a propriedade `indentLevel`.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|Representa a propriedade `protection`.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingOrder)|Representa a propriedade `readingOrder`.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinkToFit)|Representa a propriedade `shrinkToFit`.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textOrientation)|Representa a propriedade `textOrientation`.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight)|Representa a propriedade `useStandardHeight`.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth)|Representa a propriedade `useStandardWidth`.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalAlignment)|Representa a propriedade `verticalAlignment`.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wrapText)|Representa a propriedade `wrapText`.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulaHidden)|Representa a propriedade `format.protection.formulaHidden`.|
||[bloqueado](/javascript/api/excel/excel.cellpropertiesprotection#locked)|Representa a propriedade `format.protection.locked`.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueAfter)|Representa o valor após a alteração.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valueBefore)|Representa o valor antes da alteração.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valueTypeAfter)|Representa o tipo de valor após a alteração.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valueTypeBefore)|Representa o tipo de valor antes da alteração.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate__)|Ativa o gráfico na interface do usuário do Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotOptions)|Encapsula as opções para um gráfico dinâmico.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorScheme)|Especifica o esquema de cores do gráfico.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedCorners)|Especifica se a área do gráfico do gráfico tem cantos arredondados.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linkNumberFormat)|Especifica se o formato de número está vinculado às células.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowOverflow)|Especifica se o estouro de bin está habilitado em um gráfico de histograma ou gráfico de pareto.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowUnderflow)|Especifica se o subfluxo da lixeira está habilitado em um gráfico de histograma ou gráfico de pareto.|
||[Count](/javascript/api/excel/excel.chartbinoptions#count)|Especifica a contagem bin de um gráfico de histograma ou gráfico de pareto.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowValue)|Especifica o valor de estouro da lixeira de um gráfico de histograma ou gráfico de pareto.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Especifica o tipo da lixeira para um gráfico de histograma ou gráfico de pareto.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowValue)|Especifica o valor de subfluxo da lixeira de um gráfico de histograma ou gráfico de pareto.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Especifica o valor de largura da lixeira de um gráfico de histograma ou gráfico de pareto.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartileCalculation)|Especifica se o tipo de cálculo quartil de uma caixa e um gráfico de whisker.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showInnerPoints)|Especifica se os pontos internos são mostrados em uma caixa e um gráfico de whisker.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showMeanLine)|Especifica se a linha média é mostrada em uma caixa e um gráfico de whisker.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showMeanMarker)|Especifica se o marcador médio é mostrado em uma caixa e um gráfico de whisker.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showOutlierPoints)|Especifica se os pontos de outlier são mostrados em uma caixa e um gráfico de whisker.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linkNumberFormat)|Especifica se o formato de número está vinculado às células (para que o formato de número mude nos rótulos quando ele muda nas células).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linkNumberFormat)|Especifica se o formato de número está vinculado às células.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endStyleCap)|Especifica se as barras de erro têm um limite de estilo final.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Especifica quais partes das barras de erro devem ser incluídas.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Especifica o tipo de formatação das barras de erro.|
||[tipo](/javascript/api/excel/excel.charterrorbars#type)|O tipo de intervalo marcado pelas barras de erro.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Especifica se as barras de erro são exibidas.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Representa a formatação de linha do gráfico.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelStrategy)|Especifica a estratégia de rótulos de mapa de série de um gráfico de mapa de região.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Especifica o nível de mapeamento de série de um gráfico de mapa de região.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectionType)|Especifica o tipo de projeção de série de um gráfico de mapa de região.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showAxisFieldButtons)|Especifica se os botões do campo de eixo serão exibidos em um Gráfico Dinâmico.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showLegendFieldButtons)|Especifica se os botões do campo de legenda serão exibidos em um Gráfico Dinâmico.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showReportFilterFieldButtons)|Especifica se os botões do campo de filtro de relatório serão exibidos em um Gráfico Dinâmico.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showValueFieldButtons)|Especifica se os botões do campo mostrar valor serão exibidos em um Gráfico Dinâmico.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubbleScale)|Este pode ser um valor inteiro de 0 (zero) a 300, representando a porcentagem do tamanho padrão.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientMaximumColor)|Especifica a cor para o valor máximo de uma série de gráficos de mapa de região.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientMaximumType)|Especifica o tipo para o valor máximo de uma série de gráficos de mapa de região.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientMaximumValue)|Especifica o valor máximo de uma série de gráficos de mapa de região.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientMidpointColor)|Especifica a cor do valor do ponto médio de uma série de gráficos de mapa de região.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientMidpointType)|Especifica o tipo para o valor do ponto médio de uma série de gráficos de mapa de região.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientMidpointValue)|Especifica o valor do ponto médio de uma série de gráficos de mapa de região.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientMinimumColor)|Especifica a cor do valor mínimo de uma série de gráficos de mapa de região.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientMinimumType)|Especifica o tipo para o valor mínimo de uma série de gráficos de mapa de região.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientMinimumValue)|Especifica o valor mínimo de uma série de gráficos de mapa de região.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientStyle)|Especifica o estilo de gradiente de série de um gráfico de mapa de região.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertColor)|Especifica a cor de preenchimento para pontos de dados negativos em uma série.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentLabelStrategy)|Especifica a área de estratégia de rótulo pai da série para um gráfico de mapa de árvore.|
||[binOptions](/javascript/api/excel/excel.chartseries#binOptions)|Encapsula as opções de bin para gráficos de histograma e gráficos de pareto.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskerOptions)|Encapsula as opções para os gráficos de caixa estreita.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapOptions)|Encapsula as opções para um gráfico de mapa de região.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xErrorBars)|Representa o objeto da barra de erros de uma série de gráficos.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yErrorBars)|Representa o objeto da barra de erros de uma série de gráficos.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showConnectorLines)|Especifica se as linhas do conector são mostradas em gráficos de cascata.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showLeaderLines)|Especifica se as linhas de líder são exibidas para cada rótulo de dados na série.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitValue)|Especifica o valor limite que separa duas seções de um gráfico de pizza de pizza ou um gráfico de barras de pizza.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linkNumberFormat)|Especifica se o formato de número está vinculado às células (para que o formato de número mude nos rótulos quando ele muda nas células).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|Representa a propriedade `address`.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addressLocal)|Representa a propriedade `addressLocal`.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnIndex)|Representa a propriedade `columnIndex`.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getRanges__)|Retorna o , compreendendo um ou mais intervalos retangulares, aos quais o `RangeAreas` formato conditonal é aplicado.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getInvalidCells__)|Retorna um `RangeAreas` objeto, compreendendo um ou mais intervalos retangulares, com valores de célula inválidos.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getInvalidCellsOrNullObject__)|Retorna um `RangeAreas` objeto, compreendendo um ou mais intervalos retangulares, com valores de célula inválidos.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subField)|A propriedade usada pelo filtro para fazer um filtro rico em valores ricos.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Retorna o identificador de forma.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Retorna o `Shape` objeto para a forma geométrica.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getCount__)|Retorna o número de formas no grupo de forma.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getItem_key_)|Obtém uma forma usando seu nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getItemAt_index_)|Obtém uma forma com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerFooter)|O rodapé central da planilha.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerHeader)|O header central da planilha.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftFooter)|O rodapé esquerdo da planilha.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftHeader)|O header esquerdo da planilha.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightFooter)|O rodapé direito da planilha.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightHeader)|O header direito da planilha.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultForAllPages)|O cabeçalho/rodapé geral, usado em todas as páginas, a menos que seja especificada a página par/ímpar ou a primeira página.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenPages)|O cabeçalho/rodapé a ser usado para páginas pares, o cabeçalho/rodapé ímpar deve ser especificado para páginas ímpares.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstPage)|O cabeçalho/rodapé da primeira página. Para todas as outras páginas, geral ou par/ímpar é usado.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddPages)|O cabeçalho/rodapé a ser usado para páginas ímpares, o cabeçalho/rodapé par deve ser especificado para páginas pares.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|O estado pelo qual os headers/rodapés estão definidos.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#useSheetMargins)|Obtém ou define um sinalizador indicando se os cabeçalhos/rodapés estão alinhados com as margens da página que foram definidas nas opções de layout de página da planilha.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#useSheetScale)|Obtém ou define um sinalizador que indica se os cabeçalhos/rodapés devem ser dimensionados pela escala de porcentagem da página definida nas opções de layout de página da planilha.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Retorna o formato da imagem.|
||[id](/javascript/api/excel/excel.image#id)|Especifica o identificador de forma do objeto image.|
||[shape](/javascript/api/excel/excel.image#shape)|Retorna o `Shape` objeto associado à imagem.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|True se o Excel usará a interação para resolver referências circulares.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxChange)|Especifica a quantidade máxima de alteração entre cada iteração à medida que Excel resolve referências circulares.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxIteration)|Especifica o número máximo de iterações que Excel pode usar para resolver uma referência circular.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginArrowheadLength)|Representa o comprimento da ponta da seta no início da linha especificada.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginArrowheadStyle)|Representa o estilo da ponta de seta no início da linha especificada.|
||[BeginArrowheadWidth](/javascript/api/excel/excel.line#beginArrowheadWidth)|Representa a largura da ponta da seta no início da linha especificada.|
||[connectBeginShape (forma: Excel.Shape, connectionSite: número)](/javascript/api/excel/excel.line#connectBeginShape_shape__connectionSite_)|Conecta o início do conector especificado a uma forma específica.|
||[connectEndShape (forma: Excel.Shape, connectionSite: número)](/javascript/api/excel/excel.line#connectEndShape_shape__connectionSite_)|Anexa o final do conector especificado a uma forma específica.|
||[connectorType](/javascript/api/excel/excel.line#connectorType)|Representa o tipo de conector de linha.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectBeginShape__)|Desconecta o início do conector especificado de uma forma.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectEndShape__)|Desconecta o final do conector especificado de uma forma.|
||[endArrowheadLength](/javascript/api/excel/excel.line#endArrowheadLength)|Representa o comprimento da ponta de seta no final da linha especificada.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endArrowheadStyle)|Representa o estilo da ponta de seta no final da linha especificada.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endArrowheadWidth)|Representa a largura da ponta de seta no final da linha especificada.|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginConnectedShape)|Representa a forma na qual o início da linha especificada está conectado.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginConnectedSite)|Representa o site de conexão ao qual o início de um conector está conectado.|
||[endConnectedShape](/javascript/api/excel/excel.line#endConnectedShape)|Representa a forma na qual o final da linha especificada está conectado.|
||[endConnectedSite](/javascript/api/excel/excel.line#endConnectedSite)|Representa o site de conexão ao qual o final de um conector está conectado.|
||[id](/javascript/api/excel/excel.line#id)|Especifica o identificador de forma.|
||[isBeginConnected](/javascript/api/excel/excel.line#isBeginConnected)|Especifica se o início da linha especificada está conectado a uma forma.|
||[isEndConnected](/javascript/api/excel/excel.line#isEndConnected)|Especifica se o final da linha especificada está conectado a uma forma.|
||[shape](/javascript/api/excel/excel.line#shape)|Retorna o `Shape` objeto associado à linha.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete__)|Exclui um objeto de quebra de página.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getCellAfterBreak__)|Obtém a primeira célula após a quebra de página.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnIndex)|Especifica o índice de coluna para a quebra de página.|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowIndex)|Especifica o índice de linha para a quebra de página.|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add_pageBreakRange_)|Adiciona uma quebra de página antes da célula superior esquerda do intervalo especificado.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getCount__)|Obtém o número de quebras de página na coleção.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getItem_index_)|Obtém um objeto de quebra de página através do índice.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removePageBreaks__)|Redefine todas as quebras de página manuais na coleção.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackAndWhite)|A opção de impressão em preto e branco da planilha.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottomMargin)|A margem de página inferior da planilha a ser usada para impressão em pontos.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerHorizontally)|O sinalizador horizontal do centro da planilha.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centerVertically)|O sinalizador vertical do centro da planilha.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftMode)|A opção de modo de rascunho da planilha.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstPageNumber)|O número da primeira página da planilha a ser impressa.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footerMargin)|A margem do rodapé da planilha, em pontos, para uso ao imprimir.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getPrintArea__)|Obtém `RangeAreas` o objeto, compreendendo um ou mais intervalos retangulares, que representa a área de impressão da planilha.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintAreaOrNullObject__)|Obtém `RangeAreas` o objeto, compreendendo um ou mais intervalos retangulares, que representa a área de impressão da planilha.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getPrintTitleColumns__)|Obtém o objeto range que representa as colunas de título.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintTitleColumnsOrNullObject__)|Obtém o objeto range que representa as colunas de título.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getPrintTitleRows__)|Obtém o objeto range representando as linhas do título.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getPrintTitleRowsOrNullObject__)|Obtém o objeto range representando as linhas do título.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headerMargin)|A margem do header da planilha, em pontos, para uso ao imprimir.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftMargin)|A margem esquerda da planilha, em pontos, para uso ao imprimir.|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|A orientação da planilha da página.|
||[paperSize](/javascript/api/excel/excel.pagelayout#paperSize)|O tamanho do papel da página da planilha.|
||[printComments](/javascript/api/excel/excel.pagelayout#printComments)|Especifica se os comentários da planilha devem ser exibidos durante a impressão.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printErrors)|A opção erros de impressão da planilha.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printGridlines)|Especifica se as linhas de grade da planilha serão impressas.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printHeadings)|Especifica se os títulos da planilha serão impressos.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printOrder)|A opção de ordem de impressão de página da planilha.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersFooters)|Configuração de cabeçalho e rodapé da planilha.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightMargin)|A margem direita da planilha, em pontos, para uso ao imprimir.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setPrintArea_printArea_)|Define a área de impressão da planilha.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setPrintMargins_unit__marginOptions_)|Define as margens das páginas da planilha com unidades.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setPrintTitleColumns_printTitleColumns_)|Define as colunas que contêm as células que serão repetidas à esquerda de cada página da planilha para impressão.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setPrintTitleRows_printTitleRows_)|Define as linhas que contêm as células que serão repetidas na parte de cada página da planilha para impressão.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topMargin)|A margem superior da planilha, em pontos, para uso ao imprimir.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|As opções de zoom de impressão da planilha.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Especifica a margem inferior do layout da página na unidade especificada para ser usada para impressão.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Especifica a margem do rodapé de layout de página na unidade especificada para ser usada para impressão.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Especifica a margem do header de layout da página na unidade especificada para ser usada para impressão.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Especifica a margem esquerda do layout da página na unidade especificada para ser usada para impressão.|
||[direita](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Especifica a margem direita do layout da página na unidade especificada para ser usada para impressão.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Especifica a margem superior do layout da página na unidade especificada para ser usada para impressão.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalFitToPages)|Número de páginas a ser horizontalmente ajustado.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|O valor do dimensionamento da página de impressão pode estar entre 10 e 400.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalFitToPages)|Número de páginas a ser verticalmente ajustado.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortByValues_sortBy__valuesHierarchy__pivotItemScope_)|Classifica o Campo dinâmico por valores especificados em um determinado escopo.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoFormat)|Especifica se a formatação será formatada automaticamente quando for atualizada ou quando os campos são movidos.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getDataHierarchy_cell_)|Obtém o DataHierarchy que é usado para calcular o valor em um intervalo especificado dentro da Tabela dinâmica.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getPivotItems_axis__cell_)|Obtém os Itens dinâmicos de um eixo que compõem o valor em um intervalo especificado dentro da Tabela dinâmica.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveFormatting)|Especifica se a formatação é preservada quando o relatório é atualizado ou recalculado por operações como pivoting, classificação ou alteração de itens de campo de página.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setAutoSortOnCell_cell__sortBy_)|Define a Tabela Dinâmica para classificar automaticamente usando a célula especificada para selecionar automaticamente todos os critérios e contextos necessários.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enableDataValueEditing)|Especifica se a Tabela Dinâmica permite que os valores no corpo dos dados sejam editados pelo usuário.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#useCustomSortLists)|Especifica se a Tabela Dinâmica usa listas personalizadas ao classificar.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange?: Cadeia \| de caracteres de intervalo, autoFillType?: Excel. AutoFillType)](/javascript/api/excel/excel.range#autoFill_destinationRange__autoFillType_)|Preenche o intervalo do intervalo atual até o intervalo de destino usando a lógica AutoFill especificada.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertDataTypeToText__)|Converte as células de intervalo com tipos de dados em texto.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#convertToLinkedDataType_serviceID__languageCulture_)|Converte as células de intervalo em tipos de dados vinculados na planilha.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyFrom_sourceRange__copyType__skipBlanks__transpose_)|Copia dados da célula ou formatação do intervalo de origem ou `RangeAreas` do intervalo atual.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find_text__criteria_)|Localiza certa cadeia de caracteres com base em critérios especificados.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findOrNullObject_text__criteria_)|Localiza certa cadeia de caracteres com base em critérios especificados.|
||[flashFill()](/javascript/api/excel/excel.range#flashFill__)|Faz um Preenchimento Flash no intervalo atual.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getCellProperties_cellPropertiesLoadOptions_)|Retorna uma matriz 2D encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada célula.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getColumnProperties_columnPropertiesLoadOptions_)|Retorna uma única matriz dimensional encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada coluna.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getRowProperties_rowPropertiesLoadOptions_)|Retorna uma única matriz dimensional encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada célula.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getSpecialCells_cellType__cellValueType_)|Obtém o objeto, compreendendo um ou mais intervalos retangulares, que representa todas as células que corresponderem ao `RangeAreas` tipo e ao valor especificados.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getSpecialCellsOrNullObject_cellType__cellValueType_)|Obtém `RangeAreas` o objeto, compreendendo um ou mais intervalos, que representa todas as células que corresponderem ao tipo e ao valor especificados.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#getTables_fullyContained_)|Obtém uma coleção de tabelas com escopo que se sobrepõe ao intervalo.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkedDataTypeState)|Representa o estado do tipo de dados de cada célula.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeDuplicates_columns__includesHeader_)|Remove valores duplicados do intervalo especificado pelas colunas.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceAll_text__replacement__criteria_)|Localiza e substitui a cadeia de caracteres fornecida com base nos critérios especificados no intervalo atual.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setCellProperties_cellPropertiesData_)|Atualiza o intervalo com base em uma matriz 2D de propriedades de células, encapsulando coisas como fonte, preenchimento, bordas e alinhamento.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setColumnProperties_columnPropertiesData_)|Atualiza o intervalo com base em uma matriz unidimensional de propriedades de coluna, encapsulando coisas como fonte, preenchimento, bordas e alinhamento.|
||[setDirty()](/javascript/api/excel/excel.range#setDirty__)|Define um intervalo a ser recalculado quando o próximo recálculo ocorrer.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setRowProperties_rowPropertiesData_)|Atualiza o intervalo com base em uma matriz unidimensional de propriedades de linha, encapsulando coisas como fonte, preenchimento, bordas e alinhamento.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate__)|Calcula todas as células no `RangeAreas` .|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear_applyTo_)|Limpa valores, formato, preenchimento, borda e outras propriedades em cada uma das áreas que compõem esse `RangeAreas` objeto.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertDataTypeToText__)|Converte todas as células no `RangeAreas` com tipos de dados em texto.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#convertToLinkedDataType_serviceID__languageCulture_)|Converte todas as células nos tipos `RangeAreas` de dados vinculados.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyFrom_sourceRange__copyType__skipBlanks__transpose_)|Copia dados da célula ou formatação do intervalo de origem `RangeAreas` ou para o atual `RangeAreas` .|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getEntireColumn__)|Retorna um objeto que representa as colunas inteiras do (por exemplo, se a atual representa as células `RangeAreas` `RangeAreas` `RangeAreas` "B4:E11, H2", ele retorna uma que representa `RangeAreas` colunas "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getEntireRow__)|Retorna um objeto que representa as linhas inteiras do (por exemplo, se a atual representa as células "B4:E11", ele retorna um que representa linhas `RangeAreas` `RangeAreas` `RangeAreas` `RangeAreas` "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getIntersection_anotherRange_)|Retorna o `RangeAreas` objeto que representa a interseção dos intervalos determinados ou `RangeAreas` .|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getIntersectionOrNullObject_anotherRange_)|Retorna o `RangeAreas` objeto que representa a interseção dos intervalos determinados ou `RangeAreas` .|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getOffsetRangeAreas_rowOffset__columnOffset_)|Retorna um objeto que é deslocado pelo deslocamento de linha e `RangeAreas` coluna específico.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getSpecialCells_cellType__cellValueType_)|Retorna um `RangeAreas` objeto que representa todas as células que corresponderem ao tipo e ao valor especificados.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getSpecialCellsOrNullObject_cellType__cellValueType_)|Retorna um `RangeAreas` objeto que representa todas as células que corresponderem ao tipo e ao valor especificados.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#getTables_fullyContained_)|Retorna uma coleção de tabelas com escopo que se sobrepõem a qualquer intervalo neste `RangeAreas` objeto.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getUsedRangeAreas_valuesOnly_)|Retorna o usado `RangeAreas` que compreende todas as áreas usadas de intervalos retangulares individuais no `RangeAreas` objeto.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getUsedRangeAreasOrNullObject_valuesOnly_)|Retorna o usado `RangeAreas` que compreende todas as áreas usadas de intervalos retangulares individuais no `RangeAreas` objeto.|
||[address](/javascript/api/excel/excel.rangeareas#address)|Retorna a `RangeAreas` referência no estilo A1.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addressLocal)|Retorna a `RangeAreas` referência na localidade do usuário.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areaCount)|Retorna o número de intervalos retangulares que compõem esse `RangeAreas` objeto.|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|Retorna uma coleção de intervalos retangulares que compõem esse `RangeAreas` objeto.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellCount)|Retorna o número de células no objeto, somando as contagens de células de todos `RangeAreas` os intervalos retangulares individuais.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalFormats)|Retorna uma coleção de formatos condicionais que se cruzam com qualquer célula neste `RangeAreas` objeto.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#dataValidation)|Retorna um objeto de validação de dados para todos os intervalos no `RangeAreas` .|
||[format](/javascript/api/excel/excel.rangeareas#format)|Retorna um objeto, encapsulando a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades para todos os `RangeFormat` intervalos do `RangeAreas` objeto.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isEntireColumn)|Especifica se todos os intervalos neste objeto representam `RangeAreas` colunas inteiras (por exemplo, "A:C, P:Z").|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isEntireRow)|Especifica se todos os intervalos neste objeto representam linhas `RangeAreas` inteiras (por exemplo, "1:3, 5:7").|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Retorna a planilha para o `RangeAreas` atual .|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setDirty__)|Define o `RangeAreas` a ser recalculado quando o próximo recálculo ocorrer.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Representa o estilo de todos os intervalos neste `RangeAreas` objeto.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintAndShade)|Especifica um duplo que clareia ou escurece uma cor para a borda do intervalo, o valor está entre -1 (mais escuro) e 1 (mais brilhante), com 0 para a cor original.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintAndShade)|Especifica um duplo que clareia ou escurece uma cor para bordas de intervalo.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getCount__)|Retorna o número de intervalos no `RangeCollection` .|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getItemAt_index_)|Retorna o objeto range com base em sua posição no `RangeCollection` .|
||[items](/javascript/api/excel/excel.rangecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[padrão](/javascript/api/excel/excel.rangefill#pattern)|O padrão de um intervalo.|
||[patternColor](/javascript/api/excel/excel.rangefill#patternColor)|O código de cor HTML que representa a cor do padrão de intervalo, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patternTintAndShade)|Especifica um duplo que clareia ou escurece uma cor de padrão para o preenchimento do intervalo.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintAndShade)|Especifica um duplo que clareia ou escurece uma cor para o preenchimento do intervalo.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Especifica o status tachado da fonte.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Especifica o status de subscrito da fonte.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Especifica o status sobrescrito da fonte.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintAndShade)|Especifica um duplo que clareia ou escurece uma cor para a fonte de intervalo.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoIndent)|Especifica se o texto é recuado automaticamente quando o alinhamento de texto é definido como distribuição igual.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentLevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingOrder)|A ordem de leitura para o intervalo.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinkToFit)|Especifica se o texto reduz automaticamente para caber na largura da coluna disponível.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Número de linhas duplicadas removidas pela operação.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueRemaining)|Número de linhas restantes exclusivas presentes no intervalo resultante.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completeMatch)|Especifica se a combinação precisa ser completa ou parcial.|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchCase)|Especifica se a combinação é sensível a minúsculas.|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|Representa a propriedade `address`.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addressLocal)|Representa a propriedade `addressLocal`.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowIndex)|Representa a propriedade `rowIndex`.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completeMatch)|Especifica se a combinação precisa ser completa ou parcial.|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchCase)|Especifica se a combinação é sensível a minúsculas.|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchDirection)|Especifica a direção da pesquisa.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Representa a propriedade `format`.|
||[hiperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Representa a propriedade `hyperlink`.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Representa a propriedade `style`.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnHidden)|Representa a propriedade `columnHidden`.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnWidth)||
||[format: Excel. CellPropertiesFormat & {
            columnWidth?] (/javascript/api/excel/excel.settablecolumnproperties#format)|Representa a propriedade `format`.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel. CellPropertiesFormat & {
            rowHeight?] (/javascript/api/excel/excel.settablerowproperties#format)|Representa a propriedade `format`.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowHeight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowHidden)|Representa a propriedade `rowHidden`.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#altTextDescription)|Especifica o texto de descrição alternativo para um `Shape` objeto.|
||[altTextTitle](/javascript/api/excel/excel.shape#altTextTitle)|Especifica o texto de título alternativo para um `Shape` objeto.|
||[delete()](/javascript/api/excel/excel.shape#delete__)|Remove a forma da planilha.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricShapeType)|Especifica o tipo de forma geométrica dessa forma geométrica.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getAsImage_format_)|Converte a forma em uma imagem e retorna a imagem como uma cadeia de caracteres de base 64.|
||[height](/javascript/api/excel/excel.shape#height)|Especifica a altura, em pontos, da forma.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementLeft_increment_)|Move a forma horizontalmente pelo número especificado de pontos.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementRotation_increment_)|O formato é girado em sentido horário ao redor do eixo z pelo número especificado de graus.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementTop_increment_)|Move a forma verticalmente pelo número especificado de pontos.|
||[left](/javascript/api/excel/excel.shape#left)|A distância, em pontos, da lateral esquerda da forma do lado  esquerdo da planilha.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockAspectRatio)|Especifica se a proporção dessa forma está bloqueada.|
||[name](/javascript/api/excel/excel.shape#name)|Especifica o nome da forma.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionSiteCount)|Retorna o número de locais de conexão nessa forma.|
||[fill](/javascript/api/excel/excel.shape#fill)|Retorna a formatação de preenchimento dessa forma.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricShape)|Retorna a forma geométrica associada à forma.|
||[group](/javascript/api/excel/excel.shape#group)|Retorna o grupo de forma associado à forma.|
||[id](/javascript/api/excel/excel.shape#id)|Especifica o identificador de forma.|
||[image](/javascript/api/excel/excel.shape#image)|Retorna a imagem associada à forma.|
||[level](/javascript/api/excel/excel.shape#level)|Especifica o nível da forma especificada.|
||[line](/javascript/api/excel/excel.shape#line)|Retorna a linha associada à forma.|
||[lineFormat](/javascript/api/excel/excel.shape#lineFormat)|Retorna a formatação de linha do objeto de forma.|
||[onActivated](/javascript/api/excel/excel.shape#onActivated)|Ocorre quando a forma é ativada.|
||[onDeactivated](/javascript/api/excel/excel.shape#onDeactivated)|Ocorre quando a forma é desativada.|
||[parentGroup](/javascript/api/excel/excel.shape#parentGroup)|Especifica o grupo pai dessa forma.|
||[textFrame](/javascript/api/excel/excel.shape#textFrame)|Retorna o objeto text frame de uma forma.|
||[tipo](/javascript/api/excel/excel.shape#type)|Retorna o tipo dessa forma.|
||[zOrderPosition](/javascript/api/excel/excel.shape#zOrderPosition)|Retorna a posição da forma especificada na ordem z, com 0 representando a parte inferior da pilha do pedido.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Especifica a rotação, em graus, da forma.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleHeight_scaleFactor__scaleType__scaleFrom_)|Dimensiona a altura da forma por um fator especificado.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleWidth_scaleFactor__scaleType__scaleFrom_)|Dimensiona a largura da forma por um fator especificado.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setZOrder_position_)|Move a forma especificada para cima ou para baixo na ordem z da coleção, que a desloca para frente ou para trás de outras formas.|
||[top](/javascript/api/excel/excel.shape#top)|A distância, em pontos, da borda superior da forma até a borda superior da planilha.|
||[visible](/javascript/api/excel/excel.shape#visible)|Especifica se a forma está visível.|
||[width](/javascript/api/excel/excel.shape#width)|Especifica a largura, em pontos, da forma.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeId)|Obtém a ID da forma ativada.|
||[tipo](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetId)|Obtém a ID da planilha na qual a forma é ativada.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addGeometricShape_geometricShapeType_)|Adiciona uma forma geométrica à planilha.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addGroup_values_)|Um subconjunto de formas na planilha do conjunto de grupos.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addImage_base64ImageString_)|Cria uma imagem de uma cadeia de caracteres na base 64 e a adiciona à planilha.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addLine_startLeft__startTop__endLeft__endTop__connectorType_)|Adiciona uma linha à planilha.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addTextBox_text_)|Adiciona uma caixa de texto na planilha com o texto fornecido como conteúdo.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getCount__)|Retorna o número de formas da planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getItem_key_)|Obtém uma forma usando seu nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getItemAt_index_)|Obtém uma forma usando sua posição na coleção.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeId)|Obtém a ID da forma desativada.|
||[tipo](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetId)|Obtém a ID da planilha na qual a forma é desativada.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear__)|Limpa a formatação do preenchimento de um objeto de forma.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundColor)|Representa a cor de primeiro plano de preenchimento da forma no formato de cor HTML, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[tipo](/javascript/api/excel/excel.shapefill#type)|Retorna o tipo de preenchimento da forma.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setSolidColor_color_)|Define a formatação de preenchimento de um formato com uma cor uniforme.|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Especifica a porcentagem de transparência do preenchimento como um valor de 0,0 (opaco) a 1,0 (claro).|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Representa o status da fonte em negrito.|
||[color](/javascript/api/excel/excel.shapefont#color)|Representação de código de cor HTML da cor do texto (por exemplo, "#FF0000" representa vermelho).|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Representa o status da fonte em itálico.|
||[name](/javascript/api/excel/excel.shapefont#name)|Representa o nome da fonte (por exemplo, "Calibri").|
||[size](/javascript/api/excel/excel.shapefont#size)|Representa o tamanho da fonte em pontos (por exemplo, 11).|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Tipo de sublinhado aplicado à fonte.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Especifica o identificador de forma.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Retorna o `Shape` objeto associado ao grupo.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Retorna a coleção de `Shape` objetos.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup__)|Desagrupa todas as formas agrupadas no grupo de forma especificado.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Representa a cor da linha no formato de cor HTML, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashStyle)|Representa o estilo de linha da forma.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Representa o estilo de linha da forma.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Representa o grau de transparência da linha especificada como um valor de 0,0 (opaco) a 1,0 (claro).|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Especifica se a formatação de linha de um elemento de forma está visível.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Representa a espessura da linha, em pontos.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subField)|Especifica o subcampo que é o nome da propriedade de destino de um valor rico a ser classificação.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getCount__)|Obtém o número de estilos na coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getItemAt_index_)|Obtém um estilo com base em sua posição na coleção.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autoFilter)|Representa o `AutoFilter` objeto da tabela.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Obtém a origem do evento.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableId)|Obtém a ID da tabela adicionada.|
||[tipo](/javascript/api/excel/excel.tableaddedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetId)|Obtém a ID da planilha na qual a tabela é adicionada.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[detalhes](/javascript/api/excel/excel.tablechangedeventargs#details)|Obtém as informações sobre os detalhes da alteração.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onAdded)|Ocorre quando uma nova tabela é adicionada a uma workbook.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#onDeleted)|Ocorre quando a tabela especificada é excluída em uma pasta de trabalho.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Obtém a origem do evento.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableId)|Obtém a ID da tabela excluída.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tableName)|Obtém o nome da tabela excluída.|
||[tipo](/javascript/api/excel/excel.tabledeletedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetId)|Obtém a ID da planilha na qual a tabela é excluída.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getCount__)|Obtém o número de tabelas na coleção.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getFirst__)|Obtém a primeira tabela na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItem_key_)|Obtém uma tabela pelo nome ou ID.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autoSizeSetting)|As configurações de redação automáticas do quadro de texto.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottomMargin)|Representa margem inferior, em pontos, do quadro de texto.|
||[deleteText()](/javascript/api/excel/excel.textframe#deleteText__)|Exclui todo o texto no quadro de texto.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalAlignment)|Representa o alinhamento horizontal do quadro de texto.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontalOverflow)|Representa o comportamento de excedente horizontal do quadro de texto.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftMargin)|Representa margem esquerda, em pontos, do quadro de texto.|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|Representa o ângulo para o qual o texto é orientado para o quadro de texto.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingOrder)|Representa a ordem de leitura do quadro de texto, da direita para a esquerda ou da direita para a esquerda.|
||[hasText](/javascript/api/excel/excel.textframe#hasText)|Especifica se o quadro de texto contém texto.|
||[textRange](/javascript/api/excel/excel.textframe#textRange)|Representa o texto que está anexado a uma forma, bem como propriedades e métodos para manipular o texto.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightMargin)|Representa margem direita, em pontos, do quadro de texto.|
||[topMargin](/javascript/api/excel/excel.textframe#topMargin)|Representa margem superior, em pontos, do quadro de texto.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalAlignment)|Representa o alinhamento vertical do quadro de texto.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticalOverflow)|Representa o comportamento de excedente vertical do quadro de texto.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getSubstring_start__length_)|Retorna um objeto TextRange para a subcadeia de caracteres no intervalo especificado.|
||[font](/javascript/api/excel/excel.textrange#font)|Retorna um `ShapeFont` objeto que representa os atributos de fonte para o intervalo de texto.|
||[text](/javascript/api/excel/excel.textrange#text)|Representa o conteúdo de texto sem formatação do intervalo de texto.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartDataPointTrack)|True se todos os gráficos na pasta de trabalho estiverem rastreando os pontos de dados reais aos quais eles estão anexados.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getActiveChart__)|Obtém o gráfico ativo no momento na pasta de trabalho.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getActiveChartOrNullObject__)|Obtém o gráfico ativo no momento na pasta de trabalho.|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getIsActiveCollabSession__)|Retorna `true` se a workbook estiver sendo editada por vários usuários (por meio de coautor).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getSelectedRanges__)|Obtém um ou mais intervalos atualmente selecionados da pasta de trabalho.|
||[isDirty](/javascript/api/excel/excel.workbook#isDirty)|Especifica se as alterações foram feitas desde a última vez que a workbook foi salva.|
||[autoSave](/javascript/api/excel/excel.workbook#autoSave)|Especifica se a workbook está no modo AutoSave.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationEngineVersion)|Retorna um número sobre a versão do Mecanismo de Cálculo do Excel.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onAutoSaveSettingChanged)|Ocorre quando a configuração AutoSave é alterada na manual de trabalho.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslySaved)|Especifica se a workbook já foi salva localmente ou online.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#usePrecisionAsDisplayed)|True se os cálculos dessa pasta de trabalho forem efetuados usando apenas a precisão dos números conforme forem exibidos.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[tipo](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Obtém o tipo do evento.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enableCalculation)|Determina se Excel deve recalcular a planilha quando necessário.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findAll_text__criteria_)|Localiza todas as ocorrências da cadeia de caracteres determinada com base nos critérios especificados e retorna-as como um objeto, compreendendo um ou mais `RangeAreas` intervalos retangulares.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findAllOrNullObject_text__criteria_)|Localiza todas as ocorrências da cadeia de caracteres determinada com base nos critérios especificados e retorna-as como um objeto, compreendendo um ou mais `RangeAreas` intervalos retangulares.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getRanges_address_)|Obtém `RangeAreas` o objeto, representando um ou mais blocos de intervalos retangulares, especificados pelo endereço ou nome.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autoFilter)|Representa o `AutoFilter` objeto da planilha.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalPageBreaks)|Obtém a coleção de quebra de página horizontal da planilha.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onFormatChanged)|Ocorre quando o formato é alterado em uma planilha específica.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pageLayout)|Obtém `PageLayout` o objeto da planilha.|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|Retorna a coleção de todos os objetos Shape na planilha.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalPageBreaks)|Obtém a coleção de quebra de página vertical da planilha.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceAll_text__replacement__criteria_)|Localiza e substitui a cadeia de caracteres fornecida com base nos critérios especificados na planilha atual.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[detalhes](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Representa as informações sobre os detalhes da alteração.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onChanged)|Ocorre quando uma planilha da pasta de trabalho é alterada.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onFormatChanged)|Ocorre quando qualquer planilha na pasta de trabalho tem um formato alterado.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onSelectionChanged)|Ocorre quando a seleção é alterada em uma planilha.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getRange_ctx_)|Obtém o intervalo que representa a área alterada de uma planilha específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getRangeOrNullObject_ctx_)|Obtém o intervalo que representa a área alterada de uma planilha específica.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetId)|Obtém a ID da planilha na qual os dados foram alterados.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completeMatch)|Especifica se a combinação precisa ser completa ou parcial.|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchCase)|Especifica se a combinação é sensível a minúsculas.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
