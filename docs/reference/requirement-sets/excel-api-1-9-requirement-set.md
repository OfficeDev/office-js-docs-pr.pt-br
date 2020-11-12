---
title: Conjunto de requisitos de API JavaScript do Excel 1,9
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1,9.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e3878954bca943e1895a44ea9482f1c67cba9211
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996505"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>O que há de novo na API JavaScript do Excel 1,9

Mais de 500 novas APIs do Excel foram introduzidas com o conjunto de requisitos 1.9. A primeira tabela fornece um resumo conciso das APIs e, a tabela subsequente, fornece uma lista detalhada.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| [Shapes](../../excel/excel-add-ins-shapes.md) | Inserir, posicionar e formatar imagens, formas geométricas e caixas de texto. | [ShapeCollection](/javascript/api/excel/excel.shapecollection) [Shape](/javascript/api/excel/excel.shape) [GeometricShape](/javascript/api/excel/excel.geometricshape)  [Image](/javascript/api/excel/excel.image) |
| [Filtro Automático](../../excel/excel-add-ins-worksheets.md#filter-data) | Adicionar filtros aos intervalos. | [AutoFilter](/javascript/api/excel/excel.autofilter) |
| [Areas](../../excel/excel-add-ins-multiple-ranges.md) | Suporte para intervalos descontínuos. | [RangeAreas](/javascript/api/excel/excel.rangeareas) |
| [Células Especiais](../../excel/excel-add-ins-multiple-ranges.md#get-special-cells-from-multiple-ranges) | Obtenha células que contêm datas, comentários ou fórmulas dentro de um intervalo. | [Range](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|
| [Find](../../excel/excel-add-ins-ranges.md#find-a-cell-using-string-matching) | Encontre valores ou fórmulas em uma planilha ou intervalo. | [Range](/javascript/api/excel/excel.range#find-text--criteria-)[Worksheet](/javascript/api/excel/excel.worksheet#findall-text--criteria-) |
| [Copiar e colar](../../excel/excel-add-ins-ranges-advanced.md#copy-and-paste) | Copie fórmulas, formatos e valores de um intervalo para outro. | [Range](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) |
| [Calculation](../../excel/performance.md#suspend-calculation-temporarily) | Maior controle sobre o mecanismo de cálculo do Excel. | [Aplicativo](/javascript/api/excel/excel.application) |
| Novos Gráficos | Explore os novos tipos de gráficos compatíveis: mapas, caixa estreita, cascata, explosão solar, pareto. e funil. | [Chart](/javascript/api/excel/excel.charttype) |
| RangeFormat | Novos recursos com formatos de intervalo. | [Range](/javascript/api/excel/excel.rangeformat) |

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Excel 1,9. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Excel 1,9 ou anterior, confira [APIs do Excel no conjunto de requisitos 1,9 ou anterior](/javascript/api/excel?view=excel-js-1.9&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Retorna a versão do mecanismo de cálculo do Excel usada para o último recálculo completo.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Retorna o estado de cálculo do aplicativo.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Retorna as configurações do Cálculo iterativo.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Suspende a atualização da tela até que o próximo `context.sync()` seja chamado.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Aplica o AutoFiltro a um intervalo.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Limpa os critérios de filtro do AutoFiltro.|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Retorna um objeto Range que representa o intervalo no qual o Filtro automático se aplica.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|Retorna um objeto Range que representa o intervalo no qual o Filtro automático se aplica.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Uma matriz que contém todos os critérios de filtro no intervalo de autofiltro.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Especifica se o AutoFiltro está habilitado.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Especifica se o AutoFiltro tem critérios de filtro.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|Aplica o objeto Autofilter especificado que está atualmente no intervalo.|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|Remove o Filtro automático do intervalo.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#color)|Representa a propriedade `color` de uma única borda.|
||[style](/javascript/api/excel/excel.cellborder#style)|Representa a propriedade `style` de uma única borda.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#tintandshade)|Representa a propriedade `tintAndShade` de uma única borda.|
||[weight](/javascript/api/excel/excel.cellborder#weight)|Representa a propriedade `weight` de uma única borda.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#bottom)|Representa a propriedade `format.borders.bottom`.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#diagonaldown)|Representa a propriedade `format.borders.diagonalDown`.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#diagonalup)|Representa a propriedade `format.borders.diagonalUp`.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#horizontal)|Representa a propriedade `format.borders.horizontal`.|
||[left](/javascript/api/excel/excel.cellbordercollection#left)|Representa a propriedade `format.borders.left`.|
||[direita](/javascript/api/excel/excel.cellbordercollection#right)|Representa a propriedade `format.borders.right`.|
||[top](/javascript/api/excel/excel.cellbordercollection#top)|Representa a propriedade `format.borders.top`.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#vertical)|Representa a propriedade `format.borders.vertical`.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#address)|Representa a propriedade `address`.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#addresslocal)|Representa a propriedade `addressLocal`.|
||[hidden](/javascript/api/excel/excel.cellproperties#hidden)|Representa a propriedade `hidden`.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Representa a propriedade `format.fill.color`.|
||[padrão](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Representa a propriedade `format.fill.pattern`.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|Representa a propriedade `format.fill.patternColor`.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|Representa a propriedade `format.fill.patternTintAndShade`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|Representa a propriedade `format.fill.tintAndShade`.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|Representa a propriedade `format.font.bold`.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|Representa a propriedade `format.font.color`.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|Representa a propriedade `format.font.italic`.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|Representa a propriedade`format.font.name`.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|Representa a propriedade`format.font.size`.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|Representa a propriedade `format.font.strikethrough`.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|Representa a propriedade `format.font.subscript`.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|Representa a propriedade `format.font.superscript`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)|Representa a propriedade `format.font.tintAndShade`.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|Representa a propriedade `format.font.underline`.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#autoindent)|Representa a propriedade `autoIndent`.|
||[Borders](/javascript/api/excel/excel.cellpropertiesformat#borders)|Representa a propriedade `borders`.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#fill)|Representa a propriedade `fill`.|
||[font](/javascript/api/excel/excel.cellpropertiesformat#font)|Representa a propriedade `font`.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#horizontalalignment)|Representa a propriedade `horizontalAlignment`.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#indentlevel)|Representa a propriedade `indentLevel`.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#protection)|Representa a propriedade `protection`.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#readingorder)|Representa a propriedade `readingOrder`.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#shrinktofit)|Representa a propriedade `shrinkToFit`.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#textorientation)|Representa a propriedade `textOrientation`.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#usestandardheight)|Representa a propriedade `useStandardHeight`.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#usestandardwidth)|Representa a propriedade `useStandardWidth`.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#verticalalignment)|Representa a propriedade `verticalAlignment`.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#wraptext)|Representa a propriedade `wrapText`.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)|Representa a propriedade `format.protection.formulaHidden`.|
||[bloqueado](/javascript/api/excel/excel.cellpropertiesprotection#locked)|Representa a propriedade `format.protection.locked`.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|Representa o valor após a alteração.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|Representa o valor antes da alteração.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|Representa o tipo de valor após a alteração.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|Representa o tipo de valor antes da alteração.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Ativa o gráfico na interface do usuário do Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|Encapsula as opções para um gráfico dinâmico.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|Especifica o esquema de cores do gráfico.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|Especifica se a área do gráfico do gráfico tem cantos arredondados.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|Especifica se o formato de número está vinculado às células.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Especifica se o estouro de compartimento está habilitado em um gráfico de histograma ou gráfico Pareto.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Especifica se o estouro negativo de compartimento está habilitado em um gráfico de histograma ou gráfico Pareto.|
||[Count](/javascript/api/excel/excel.chartbinoptions#count)|Especifica a contagem bin de um gráfico de histograma ou gráfico Pareto.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Especifica o valor de estouro de bin de um gráfico de histograma ou gráfico Pareto.|
||[type](/javascript/api/excel/excel.chartbinoptions#type)|Especifica o tipo de compartimento de um gráfico de histograma ou gráfico Pareto.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Especifica o valor de Subfluxo bin de um gráfico de histograma ou gráfico Pareto.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Especifica o valor da largura bin de um gráfico de histograma ou gráfico Pareto.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Especifica se o tipo de cálculo de quartil de uma caixa e um gráfico de caixa estreita.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Especifica se os pontos internos são mostrados em um gráfico de caixa e estreita.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Especifica se a linha de média é mostrada em um gráfico de caixa e estreitamento.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Especifica se o marcador médio é mostrado em um gráfico de caixa e mais à estreita.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Especifica se os pontos de exceção são mostrados em um gráfico de caixa e mais à estreita.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Especifica se o formato de número está vinculado às células (de modo que o formato de número seja alterado nos rótulos quando for alterado nas células).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Especifica se o formato de número está vinculado às células.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Especifica se as barras de erro têm um limite de estilo final.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Especifica quais partes das barras de erro devem ser incluídas.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Especifica o tipo de formatação das barras de erro.|
||[tipo](/javascript/api/excel/excel.charterrorbars#type)|O tipo de intervalo marcado pelas barras de erro.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Especifica se as barras de erro são exibidas.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Representa a formatação de linha do gráfico.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Especifica a estratégia de rótulos de mapa de séries de um gráfico de mapa de região.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Especifica o nível de mapeamento de séries de um gráfico de mapa de região.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Especifica o tipo de projeção de série de um gráfico de mapa de região.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Especifica se os botões de campo de eixo devem ser exibidos em um gráfico dinâmico.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Especifica se os botões de campo de legenda devem ser exibidos em um gráfico dinâmico.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Especifica se os botões de campo de filtro de relatório devem ser exibidos em um gráfico dinâmico.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Especifica se os botões de campo de valor de exibição devem ser exibidos em um gráfico dinâmico.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|Este pode ser um valor inteiro de 0 (zero) a 300, representando a porcentagem do tamanho padrão.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|Especifica a cor para o valor máximo de uma série de gráficos do mapa de região.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|Especifica o tipo de valor máximo de uma série de gráfico do mapa de região.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|Especifica o valor máximo de uma série de gráficos de mapas de região.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|Especifica a cor do valor intermediário de uma série de gráficos do mapa de região.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|Especifica o tipo de valor de ponto médio de uma série de gráficos de mapas de região.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|Especifica o valor de ponto médio de uma série de gráficos de mapas de região.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|Especifica a cor do valor mínimo de uma série de gráfico do mapa de região.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|Especifica o tipo de valor mínimo de uma série de gráfico do mapa de região.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|Especifica o valor mínimo de uma série de gráficos de mapas de região.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|Especifica o estilo de gradiente de série de um gráfico de mapa de região.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|Especifica a cor de preenchimento de pontos de dados negativos em uma série.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|Especifica a área de estratégia de rótulo pai da série para um gráfico de mapa de região.|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|Encapsula as opções de bin para gráficos de histograma e gráficos de pareto.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|Encapsula as opções para os gráficos de caixa estreita.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|Encapsula as opções para um gráfico de mapa de região.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|Representa o objeto da barra de erros de uma série de gráficos.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|Representa o objeto da barra de erros de uma série de gráficos.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|Especifica se as linhas de conexão são mostradas em gráficos de cascata.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|Especifica se as linhas de preenchimento são exibidas para cada rótulo de dados na série.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|Especifica o valor de limite que separa duas seções de um gráfico de pizza de pizza ou de barra de pizza.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Especifica se o formato de número está vinculado às células (de modo que o formato de número seja alterado nos rótulos quando for alterado nas células).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|Representa a propriedade `address`.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|Representa a propriedade `addressLocal`.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|Representa a propriedade `columnIndex`.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Retorna o RangeAreas, compreendendo um ou mais intervalos retangulares, ao qual o formato condicional é aplicado.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Retorna um RangeAreas, que consiste em um ou mais intervalos retangulares, com valores inválidos de célula.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Retorna um RangeAreas, que consiste em um ou mais intervalos retangulares, com valores inválidos de célula.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|A propriedade usada pelo filtro para realizar a filtragem avançada em richvalues.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Retorna o identificador de forma.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Retorna o objeto de Forma para a forma geométrica.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Retorna o número de formas no grupo de forma.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|Obtém uma forma usando seu nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Obtém uma forma com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|O rodapé central da planilha.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|O cabeçalho central da planilha.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|O rodapé esquerdo da planilha.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|O cabeçalho esquerdo da planilha.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|O rodapé direito da planilha.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|O cabeçalho direito da planilha.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|O cabeçalho/rodapé geral, usado em todas as páginas, a menos que seja especificada a página par/ímpar ou a primeira página.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|O cabeçalho/rodapé a ser usado para páginas pares, o cabeçalho/rodapé ímpar deve ser especificado para páginas ímpares.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|O cabeçalho/rodapé da primeira página. Para todas as outras páginas, geral ou par/ímpar é usado.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|O cabeçalho/rodapé a ser usado para páginas ímpares, o cabeçalho/rodapé par deve ser especificado para páginas pares.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|O estado pelo qual os cabeçalhos/rodapés são definidos.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|Obtém ou define um sinalizador indicando se os cabeçalhos/rodapés estão alinhados com as margens da página que foram definidas nas opções de layout de página da planilha.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|Obtém ou define um sinalizador que indica se os cabeçalhos/rodapés devem ser dimensionados pela escala de porcentagem da página definida nas opções de layout de página da planilha.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Retorna o formato da imagem.|
||[id](/javascript/api/excel/excel.image#id)|Especifica o identificador de forma para o objeto Image.|
||[shape](/javascript/api/excel/excel.image#shape)|Retorna o objeto de forma associado à imagem.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|True se o Excel usará a interação para resolver referências circulares.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Especifica a quantidade máxima de alteração entre cada iteração à medida que o Excel resolve referências circulares.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Especifica o número máximo de iterações que o Excel pode usar para resolver uma referência circular.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#beginarrowheadlength)|Representa o comprimento da ponta da seta no início da linha especificada.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#beginarrowheadstyle)|Representa o estilo da ponta de seta no início da linha especificada.|
||[BeginArrowheadWidth](/javascript/api/excel/excel.line#beginarrowheadwidth)|Representa a largura da ponta da seta no início da linha especificada.|
||[connectBeginShape (forma: Excel.Shape, connectionSite: número)](/javascript/api/excel/excel.line#connectbeginshape-shape--connectionsite-)|Conecta o início do conector especificado a uma forma específica.|
||[connectEndShape (forma: Excel.Shape, connectionSite: número)](/javascript/api/excel/excel.line#connectendshape-shape--connectionsite-)|Anexa o final do conector especificado a uma forma específica.|
||[connectorType](/javascript/api/excel/excel.line#connectortype)|Representa o tipo de conector de linha.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#disconnectbeginshape--)|Desconecta o início do conector especificado de uma forma.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#disconnectendshape--)|Desconecta o final do conector especificado de uma forma.|
||[endArrowheadLength](/javascript/api/excel/excel.line#endarrowheadlength)|Representa o comprimento da ponta de seta no final da linha especificada.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#endarrowheadstyle)|Representa o estilo da ponta de seta no final da linha especificada.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#endarrowheadwidth)|Representa a largura da ponta de seta no final da linha especificada.|
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|Representa a forma na qual o início da linha especificada está conectado.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|Representa o site de conexão ao qual o início de um conector está conectado.|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|Representa a forma na qual o final da linha especificada está conectado.|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|Representa o site de conexão ao qual o final de um conector está conectado.|
||[id](/javascript/api/excel/excel.line#id)|Especifica o identificador da forma.|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|Especifica se o início da linha especificada está conectado a uma forma.|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|Especifica se o final da linha especificada está conectado a uma forma.|
||[shape](/javascript/api/excel/excel.line#shape)|Retorna o objeto de forma associado à linha.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Exclui um objeto de quebra de página.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|Obtém a primeira célula após a quebra de página.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Especifica o índice de coluna para a quebra de página|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Especifica o índice de linha para a quebra de página|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Adiciona uma quebra de página antes da célula superior esquerda do intervalo especificado.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Obtém o número de quebras de página na coleção.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Obtém um objeto de quebra de página através do índice.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|Redefine todas as quebras de página manuais na coleção.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|A opção de impressão em preto e branco da planilha.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|A margem inferior da página da planilha a ser usada para impressão em pontos.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|O sinalizador da planilha centralizado horizontalmente.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|O sinalizador da planilha centralizado verticalmente.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|A opção de modo de rascunho da planilha.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|O número da primeira página da planilha a ser impressa.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|A margem de rodapé da planilha, em pontos, para uso na impressão.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos retangulares, que representa a área de impressão da planilha.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos retangulares, que representa a área de impressão da planilha.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|Obtém o objeto range que representa as colunas de título.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|Obtém o objeto range que representa as colunas de título.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|Obtém o objeto range representando as linhas do título.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|Obtém o objeto range representando as linhas do título.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|A margem do cabeçalho da planilha, em pontos, para uso na impressão.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|A margem esquerda da planilha, em pontos, para uso na impressão.|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|A orientação da planilha da página.|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|O tamanho de papel da página da planilha.|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|Especifica se os comentários da planilha devem ser exibidos ao imprimir.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|A opção de erros de impressão da planilha.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|Especifica se as linhas de grade da planilha serão impressas.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|Especifica se os títulos da planilha serão impressos.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|A opção de ordem de impressão de página da planilha.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|Configuração de cabeçalho e rodapé da planilha.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|A margem direita da planilha, em pontos, para uso na impressão.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|Define a área de impressão da planilha.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Define as margens das páginas da planilha com unidades.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|Define as colunas que contêm as células que serão repetidas à esquerda de cada página da planilha para impressão.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|Define as linhas que contêm as células que serão repetidas na parte de cada página da planilha para impressão.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|A margem superior da planilha, em pontos, para uso na impressão.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|As opções de zoom de impressão da planilha.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Especifica a margem inferior do layout da página na unidade especificada para ser usada para impressão.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Especifica a margem de rodapé do layout da página na unidade especificada para ser usada para impressão.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Especifica a margem do cabeçalho do layout da página na unidade especificada para ser usada para impressão.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Especifica a margem esquerda do layout da página na unidade especificada para ser usada para impressão.|
||[direita](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Especifica a margem direita do layout da página na unidade especificada para ser usada para impressão.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Especifica a margem superior do layout da página na unidade especificada para ser usada para impressão.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|Número de páginas a ser horizontalmente ajustado.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|O valor do dimensionamento da página de impressão pode estar entre 10 e 400.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|Número de páginas a ser verticalmente ajustado.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Classifica o Campo dinâmico por valores especificados em um determinado escopo.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|Especifica se a formatação será formatada automaticamente quando for atualizada ou quando os campos forem movidos.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Obtém o DataHierarchy que é usado para calcular o valor em um intervalo especificado dentro da Tabela dinâmica.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Obtém os Itens dinâmicos de um eixo que compõem o valor em um intervalo especificado dentro da Tabela dinâmica.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|Especifica se a formatação é preservada quando o relatório é atualizado ou recalculado por operações como dinamização, classificação ou alteração de itens de campo de página.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Define a Tabela Dinâmica para classificar automaticamente usando a célula especificada para selecionar automaticamente todos os critérios e contextos necessários.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|Especifica se a tabela dinâmica permite que valores no corpo de dados sejam editados pelo usuário.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|Especifica se a tabela dinâmica usa listas personalizadas ao classificar.|
|[Range](/javascript/api/excel/excel.range)|[Preenchimento automático (destinationRange?: \| cadeia de caracteres de intervalo, Autofilltype?: Excel. Autofilltype)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Preenche o intervalo do intervalo atual com o intervalo de destino especificado usando a lógica de preenchimento automático especificada.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Converte o intervalo de células com tipos de dados em texto.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Converte as células de intervalo em um tipo de dados vinculado na planilha.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copia a formatação ou dados da célula do intervalo de origem ou de RangeAreas para o intervalo atual.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Localiza certa cadeia de caracteres com base em critérios especificados.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Localiza certa cadeia de caracteres com base em critérios especificados.|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|Faz o preenchimento relâmpago no intervalo atual. O preenchimento relâmpago preenche automaticamente dados quando detecta um padrão. Portanto, o intervalo deve ser de coluna única e ter dados em torno para encontrar o padrão.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Retorna uma matriz 2D encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada célula.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Retorna uma única matriz dimensional encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada coluna.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Retorna uma única matriz dimensional encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada célula.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos retangulares, que representa todas as células que correspondem ao tipo e valor especificado.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos, que representa todas as células que correspondem ao tipo e valor especificado.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|Obtém uma coleção de tabelas com escopo que se sobrepõe ao intervalo.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|Representa o estado do tipo de dados de cada célula.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|Remove valores duplicados do intervalo especificado pelas colunas.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|Localiza e substitui a cadeia de caracteres fornecida com base nos critérios especificados no intervalo atual.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|Atualiza o intervalo com base em uma matriz 2D de propriedades da célula, encapsulando itens como fonte, preenchimento, bordas, alinhamento e assim por diante.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|Atualiza o intervalo com base em uma única matriz dimensional de propriedades da coluna, encapsulando itens como fonte, preenchimento, bordas, alinhamento e assim por diante.|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|Define um intervalo a ser recalculado quando o próximo recálculo ocorrer.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|Atualiza o intervalo com base em uma única matriz dimensional de propriedades da linha, encapsulando itens como fonte, preenchimento, bordas, alinhamento e assim por diante.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|Calcula todas as células no RangeAreas.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Limpa valores, formato, preenchimento, borda, etc. em cada uma das áreas que compõe este objeto RangeAreas.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|Converte todas as células de RangeAreas com tipos de dados em texto.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|Converte todas as células de RangeAreas em tipos de dados vinculados.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copia a formatação ou dados da célula do intervalo de origem ou de RangeAreas para o RangeAreas atual.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|Retorna um objeto RangeAreas que representa as colunas inteiras dos objetos RangeAreas (por exemplo, se o RangeAreas atual representa as células "B4:E11, H2", ele retorna um RangeAreas que representa as colunas "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|Retorna um objeto RangeAreas que representa as linhas inteiras dos objetos RangeAreas (por exemplo, se o RangeAreas atual representa as células "B4:E11", ele retorna um RangeAreas que representa as linhas "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|Retorna o objeto RangeAreas que representa a interseção dos intervalos fornecidos ou RangeAreas.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|Retorna o objeto RangeAreas que representa a interseção dos intervalos fornecidos ou RangeAreas.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|Retorna um objeto RangeAreas que é deslocado pelo deslocamento de linha e coluna específico.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Retorna um objeto RangeAreas que representa todas as células que correspondem ao tipo e valor especificados.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Retorna um objeto RangeAreas que representa todas as células que correspondem ao tipo e valor especificados.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|Retorna uma coleção de tabelas com escopo que se sobrepõe a qualquer intervalo neste objeto RangeAreas.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|Retorna o RangeAreas usado que compreende todas as áreas utilizadas de intervalos retangulares individuais no objeto RangeAreas.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|Retorna o RangeAreas usado que compreende todas as áreas utilizadas de intervalos retangulares individuais no objeto RangeAreas.|
||[address](/javascript/api/excel/excel.rangeareas#address)|Retorna a referência RangeAreas em estilo a1.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|Retorna a referência RangeAreas na localidade do usuário.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|Retorna o número de intervalos retangulares que compõem este objeto RangeAreas.|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|Retorna uma coleção de intervalos retangulares que compõem este objeto RangeAreas.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|Retorna o número de células no objeto RangeAreas somando as contagens de células de todos os intervalos retangulares individuais.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|Retorna uma coleção de ConditionalFormats que se cruza com qualquer célula nesse objeto RangeAreas.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|Retorna um objeto dataValidation para todos os intervalos no RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareas#format)|Retorna um objeto RangeFormat, encapsulando a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades para todos os intervalos no objeto RangeAreas.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|Especifica se todos os intervalos deste objeto RangeAreas representam colunas inteiras (por exemplo, "A:C, Q:Z").|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|Especifica se todos os intervalos deste objeto RangeAreas representam linhas inteiras (por exemplo, "1:3, 5:7").|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Retorna a planilha para o RangeAreas atual.|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|Define o RangeAreas que será recalculado quando o próximo recálculo ocorrer.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Representa o estilo de todos os intervalos nesse objeto RangeAreas.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Especifica um duplo que clareia ou escurece uma cor para a borda do intervalo, o valor é entre-1 (mais escuro) e 1 (mais brilhante), com 0 para a cor original.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Especifica um duplo que clareia ou escurece uma cor para bordas de intervalo, o valor é entre-1 (mais escuro) e 1 (mais brilhante), com 0 para a cor original.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Retorna o número de intervalos no RangeCollection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Retorna o objeto range com base em sua posição no RangeCollection.|
||[items](/javascript/api/excel/excel.rangecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[padrão](/javascript/api/excel/excel.rangefill#pattern)|O padrão de um intervalo.|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|O código de cor HTML que representa a cor do padrão de intervalo, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|Especifica um duplo que clareia ou escurece uma cor de padrão para o preenchimento de intervalo, o valor é entre-1 (mais escuro) e 1 (mais brilhante), com 0 para a cor original.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|Especifica um duplo que clareia ou escurece uma cor para o preenchimento de intervalo, o valor é entre-1 (mais escuro) e 1 (mais brilhante), com 0 para a cor original.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Especifica o status tachado da fonte.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Especifica o status subscrito da fonte.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Especifica o status sobrescrito da fonte.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|Especifica um duplo que clareia ou escurece uma cor para a fonte do intervalo, o valor é entre-1 (mais escuro) e 1 (mais brilhante), com 0 para a cor original.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|Especifica se o texto será recuado automaticamente quando o alinhamento do texto for definido como distribuição igual.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|A ordem de leitura para o intervalo.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|Especifica se o texto é automaticamente reduzido para se ajustar à largura de coluna disponível.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Número de linhas duplicadas removidas pela operação.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|Número de linhas restantes exclusivas presentes no intervalo resultante.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|Especifica se a correspondência precisa ser completa ou parcial.|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|Especifica se a correspondência diferencia maiúsculas de minúsculas.|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|Representa a propriedade `address`.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|Representa a propriedade `addressLocal`.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|Representa a propriedade `rowIndex`.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|Especifica se a correspondência precisa ser completa ou parcial.|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|Especifica se a correspondência diferencia maiúsculas de minúsculas.|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|Especifica a direção da pesquisa.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Representa a propriedade `format`.|
||[hiperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Representa a propriedade `hyperlink`.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Representa a propriedade `style`.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|Representa a propriedade `columnHidden`.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[formato: Excel. CellPropertiesFormat & {columnWidth?](/javascript/api/excel/excel.settablecolumnproperties#format)|Representa a propriedade `format`.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[formato: Excel. CellPropertiesFormat & {AlturaDaLinha?](/javascript/api/excel/excel.settablerowproperties#format)|Representa a propriedade `format`.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|Representa a propriedade `rowHidden`.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Especifica o texto de descrição alternativa para um objeto Shape.|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Especifica o texto de título alternativo para um objeto Shape.|
||[delete()](/javascript/api/excel/excel.shape#delete--)|Remove a forma da planilha.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|Especifica o tipo de forma geométrica dessa forma geométrica.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|Converte a forma em uma imagem e retorna a imagem como uma cadeia de caracteres de base 64.|
||[height](/javascript/api/excel/excel.shape#height)|Especifica a altura, em pontos, da forma.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|Move a forma horizontalmente pelo número especificado de pontos.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|O formato é girado em sentido horário ao redor do eixo z pelo número especificado de graus.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|Move a forma verticalmente pelo número especificado de pontos.|
||[left](/javascript/api/excel/excel.shape#left)|A distância, em pontos, da lateral esquerda da forma do lado  esquerdo da planilha.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|Especifica se a taxa de proporção desta forma está bloqueada.|
||[name](/javascript/api/excel/excel.shape#name)|Especifica o nome da forma.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|Retorna o número de locais de conexão nessa forma.|
||[fill](/javascript/api/excel/excel.shape#fill)|Retorna a formatação de preenchimento dessa forma.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|Retorna a forma geométrica associada à forma.|
||[group](/javascript/api/excel/excel.shape#group)|Retorna o grupo de forma associado à forma.|
||[id](/javascript/api/excel/excel.shape#id)|Especifica o identificador da forma.|
||[image](/javascript/api/excel/excel.shape#image)|Retorna a imagem associada à forma.|
||[level](/javascript/api/excel/excel.shape#level)|Especifica o nível da forma especificada.|
||[line](/javascript/api/excel/excel.shape#line)|Retorna a linha associada à forma.|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|Retorna a formatação de linha do objeto de forma.|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|Ocorre quando a forma é ativada.|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|Ocorre quando a forma é desativada.|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|Especifica o grupo pai desta forma.|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|Retorna o objeto text frame de uma forma.|
||[tipo](/javascript/api/excel/excel.shape#type)|Retorna o tipo dessa forma.|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|Retorna a posição da forma especificada na ordem z, com 0 representando a parte inferior da pilha do pedido.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Especifica a rotação, em graus, da forma.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Dimensiona a altura da forma por um fator especificado.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Dimensiona a largura da forma por um fator especificado.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|Move a forma especificada para cima ou para baixo na ordem z da coleção, que a desloca para frente ou para trás de outras formas.|
||[top](/javascript/api/excel/excel.shape#top)|A distância, em pontos, da borda superior da forma até a borda superior da planilha.|
||[visible](/javascript/api/excel/excel.shape#visible)|Especifica se a forma está visível.|
||[width](/javascript/api/excel/excel.shape#width)|Especifica a largura, em pontos, da forma.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Obtém o id da forma ativada.|
||[tipo](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Obtém a id da planilha na qual a forma está ativada.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Adiciona uma forma geométrica à planilha.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Um subconjunto de formas na planilha do conjunto de grupos.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Cria uma imagem de uma cadeia de caracteres na base 64 e a adiciona à planilha.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Adiciona uma linha à planilha.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Adiciona uma caixa de texto na planilha com o texto fornecido como conteúdo.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Retorna o número de formas da planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|Obtém uma forma usando seu nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Obtém uma forma usando sua posição na coleção.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Obtém o id da forma que está desativada.|
||[tipo](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Obtém a id da planilha na qual a forma está desativada.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Limpa a formatação do preenchimento de um objeto de forma.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|Representa a cor de primeiro plano do preenchimento da forma no formato de cor HTML, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[tipo](/javascript/api/excel/excel.shapefill#type)|Retorna o tipo de preenchimento da forma.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Define a formatação de preenchimento de um formato com uma cor uniforme.|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Especifica a porcentagem de transparência do preenchimento como um valor de 0,0 (opaco) a 1,0 (claro).|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Representa o status da fonte em negrito.|
||[color](/javascript/api/excel/excel.shapefont#color)|Representação do código de cor HTML da cor do texto (por exemplo, "#FF0000" representa vermelho).|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Representa o status da fonte em itálico.|
||[name](/javascript/api/excel/excel.shapefont#name)|Representa o nome da fonte (por exemplo, "Calibri").|
||[size](/javascript/api/excel/excel.shapefont#size)|Representa o tamanho da fonte em pontos (por exemplo, 11).|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Tipo de sublinhado aplicado à fonte.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Especifica o identificador da forma.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Retorna o objeto de forma associado ao grupo.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Retorna uma coleção de objetos de forma.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Desagrupa todas as formas agrupadas no grupo de forma especificado.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Representa a cor da linha no formato de cor HTML, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Representa o estilo de linha da forma.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Representa o estilo de linha da forma.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Representa o grau de transparência da linha especificada como um valor de 0,0 (opaco) a 1,0 (claro).|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Especifica se a formatação de linha de um elemento Shape é visível.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Representa a espessura da linha, em pontos.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|Especifica o subcampo que é o nome da propriedade de destino de um valor avançado a ser classificado.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|Obtém o número de estilos na coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|Obtém um estilo com base em sua posição na coleção.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|Representa o objeto AutoFilter da tabela.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Obtém a origem do evento.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|Obtém a id da tabela que é adicionada.|
||[tipo](/javascript/api/excel/excel.tableaddedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|Obtém a id da planilha na qual o gráfico é adicionado.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[detalhes](/javascript/api/excel/excel.tablechangedeventargs#details)|Obtém as informações sobre os detalhes de alteração.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|Ocorre quando uma nova tabela é adicionada na pasta de trabalho.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|Ocorre quando a tabela especificada é excluída em uma pasta de trabalho.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Obtém a origem do evento.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Obtém a ID da tabela que é excluída.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Obtém o nome da tabela que é excluída.|
||[tipo](/javascript/api/excel/excel.tabledeletedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Obtém a ID da planilha na qual a tabela é excluída.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Obtém o número de tabelas na coleção.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Obtém a primeira tabela na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Obtém uma tabela pelo nome ou ID.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|As configurações de dimensionamento automático do quadro de texto.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|Representa margem inferior, em pontos, do quadro de texto.|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|Exclui todo o texto no quadro de texto.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|Representa o alinhamento horizontal do quadro de texto.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|Representa o comportamento de excedente horizontal do quadro de texto.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|Representa margem esquerda, em pontos, do quadro de texto.|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|Representa o ângulo no qual o texto é orientado para o quadro de texto.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|Representa a ordem de leitura do quadro de texto, da direita para a esquerda ou da direita para a esquerda.|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|Especifica se o quadro de texto contém texto.|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|Representa o texto que está anexado a uma forma, bem como propriedades e métodos para manipular o texto.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|Representa margem direita, em pontos, do quadro de texto.|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Representa margem superior, em pontos, do quadro de texto.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Representa o alinhamento vertical do quadro de texto.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Representa o comportamento de excedente vertical do quadro de texto.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|Retorna um objeto TextRange para a subcadeia de caracteres no intervalo especificado.|
||[font](/javascript/api/excel/excel.textrange#font)|Retorna um objeto ShapeFont que representa os atributos de fonte do intervalo de texto.|
||[text](/javascript/api/excel/excel.textrange#text)|Representa o conteúdo de texto sem formatação do intervalo de texto.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|True se todos os gráficos na pasta de trabalho estiverem rastreando os pontos de dados reais aos quais eles estão anexados.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|Obtém o gráfico ativo no momento na pasta de trabalho.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|Obtém o gráfico ativo no momento na pasta de trabalho.|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|True se a pasta de trabalho estiver sendo editada por vários usuários (coautoria).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|Obtém um ou mais intervalos atualmente selecionados da pasta de trabalho.|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|Especifica se foram feitas alterações desde a última vez em que a pasta de trabalho foi salva.|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|Especifica se a pasta de trabalho está no modo de salvamento automático.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Retorna um número sobre a versão do Mecanismo de Cálculo do Excel.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|Ocorre quando a configuração Salvamento automático é alterada na pasta de trabalho.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|Especifica se a pasta de trabalho já foi salva localmente ou online.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|True se os cálculos dessa pasta de trabalho forem efetuados usando apenas a precisão dos números conforme forem exibidos.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[tipo](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Obtém o tipo do evento.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|Determina se o Excel deve recalcular a planilha quando necessário.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|Encontra todas as ocorrências de determinada cadeia de caracteres com base nos critérios especificados e as retorna como um objeto RangeAreas, compreendendo um ou mais intervalos retangulares.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|Encontra todas as ocorrências de determinada cadeia de caracteres com base nos critérios especificados e as retorna como um objeto RangeAreas, compreendendo um ou mais intervalos retangulares.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|Obtém o objeto RangeAreas que representa um ou mais blocos de intervalos retangulares especificados pelo endereço ou nome.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|Representa o objeto AutoFilter da planilha.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|Obtém a coleção de quebra de página horizontal da planilha.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|Ocorre quando o formato é alterado em uma planilha específica.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|Obtém o objeto PageLayout da planilha.|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|Retorna a coleção de todos os objetos Shape na planilha.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|Obtém a coleção de quebra de página vertical da planilha.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|Localiza e substitui a cadeia de caracteres fornecida com base nos critérios especificados na planilha atual.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[detalhes](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Representa as informações sobre os detalhes da alteração.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|Ocorre quando uma planilha da pasta de trabalho é alterada.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|Ocorre quando uma planilha na pasta de trabalho tem o formato alterado.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|Ocorre quando a seleção é alterada em uma planilha.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Obtém o intervalo que representa a área alterada de uma planilha específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Obtém o intervalo que representa a área alterada de uma planilha específica.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Obtém o id da planilha na qual os dados são alterados.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Especifica se a correspondência precisa ser completa ou parcial.|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Especifica se a correspondência diferencia maiúsculas de minúsculas.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
