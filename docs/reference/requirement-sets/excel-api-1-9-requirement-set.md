---
title: Excel de requisitos da API JavaScript 1.9
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.9.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
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

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.9. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.9 ou anterior, consulte Excel APIs no conjunto de requisitos [1.9](/javascript/api/excel?view=excel-js-1.9&preserve-view=true) ou anterior.

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#excel-excel-application-calculationengineversion-member)|Retorna a versão do mecanismo de cálculo do Excel usada para o último recálculo completo.|
||[calculationState](/javascript/api/excel/excel.application#excel-excel-application-calculationstate-member)|Retorna o estado de cálculo do aplicativo.|
||[iterativeCalculation](/javascript/api/excel/excel.application#excel-excel-application-iterativecalculation-member)|Retorna as configurações de cálculo iterativo.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#excel-excel-application-suspendscreenupdatinguntilnextsync-member(1))|Suspende a atualização de tela até que a próxima `context.sync()` seja chamada.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-apply-member(1))|Aplica o AutoFiltro a um intervalo.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-clearcriteria-member(1))|Limpa os critérios de filtro e o estado de classificação do AutoFilter.|
||[criteria](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-criteria-member)|Uma matriz que contém todos os critérios de filtro no intervalo de autofiltro.|
||[enabled](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-enabled-member)|Especifica se o AutoFilter está habilitado.|
||[getRange()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrange-member(1))|Retorna o `Range` objeto que representa o intervalo ao qual o AutoFilter se aplica.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-getrangeornullobject-member(1))|Retorna o `Range` objeto que representa o intervalo ao qual o AutoFilter se aplica.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-isdatafiltered-member)|Especifica se o Filtro Automático tem critérios de filtro.|
||[reapply()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-reapply-member(1))|Aplica o objeto Autofilter especificado que está atualmente no intervalo.|
||[remove()](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-remove-member(1))|Remove o Filtro automático do intervalo.|
|[CellBorder](/javascript/api/excel/excel.cellborder)|[color](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-color-member)|Representa a propriedade `color` de uma única borda.|
||[style](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-style-member)|Representa a propriedade `style` de uma única borda.|
||[tintAndShade](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-tintandshade-member)|Representa a propriedade `tintAndShade` de uma única borda.|
||[weight](/javascript/api/excel/excel.cellborder#excel-excel-cellborder-weight-member)|Representa a propriedade `weight` de uma única borda.|
|[CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)|[bottom](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-bottom-member)|Representa a propriedade `format.borders.bottom`.|
||[diagonalDown](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-diagonaldown-member)|Representa a propriedade `format.borders.diagonalDown`.|
||[diagonalUp](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-diagonalup-member)|Representa a propriedade `format.borders.diagonalUp`.|
||[horizontal](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-horizontal-member)|Representa a propriedade `format.borders.horizontal`.|
||[left](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-left-member)|Representa a propriedade `format.borders.left`.|
||[direita](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-right-member)|Representa a propriedade `format.borders.right`.|
||[top](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-top-member)|Representa a propriedade `format.borders.top`.|
||[vertical](/javascript/api/excel/excel.cellbordercollection#excel-excel-cellbordercollection-vertical-member)|Representa a propriedade `format.borders.vertical`.|
|[CellProperties](/javascript/api/excel/excel.cellproperties)|[address](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-address-member)|Representa a propriedade `address`.|
||[addressLocal](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-addresslocal-member)|Representa a propriedade `addressLocal`.|
||[hidden](/javascript/api/excel/excel.cellproperties#excel-excel-cellproperties-hidden-member)|Representa a propriedade `hidden`.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-color-member)|Representa a propriedade `format.fill.color`.|
||[padrão](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-pattern-member)|Representa a propriedade `format.fill.pattern`.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-patterncolor-member)|Representa a propriedade `format.fill.patternColor`.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-patterntintandshade-member)|Representa a propriedade `format.fill.patternTintAndShade`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#excel-excel-cellpropertiesfill-tintandshade-member)|Representa a propriedade `format.fill.tintAndShade`.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-bold-member)|Representa a propriedade `format.font.bold`.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-color-member)|Representa a propriedade `format.font.color`.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-italic-member)|Representa a propriedade `format.font.italic`.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-name-member)|Representa a propriedade`format.font.name`.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-size-member)|Representa a propriedade`format.font.size`.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-strikethrough-member)|Representa a propriedade `format.font.strikethrough`.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-subscript-member)|Representa a propriedade `format.font.subscript`.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-superscript-member)|Representa a propriedade `format.font.superscript`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-tintandshade-member)|Representa a propriedade `format.font.tintAndShade`.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#excel-excel-cellpropertiesfont-underline-member)|Representa a propriedade `format.font.underline`.|
|[CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-autoindent-member)|Representa a propriedade `autoIndent`.|
||[Borders](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-borders-member)|Representa a propriedade `borders`.|
||[fill](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-fill-member)|Representa a propriedade `fill`.|
||[font](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-font-member)|Representa a propriedade `font`.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-horizontalalignment-member)|Representa a propriedade `horizontalAlignment`.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-indentlevel-member)|Representa a propriedade `indentLevel`.|
||[protection](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-protection-member)|Representa a propriedade `protection`.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-readingorder-member)|Representa a propriedade `readingOrder`.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-shrinktofit-member)|Representa a propriedade `shrinkToFit`.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-textorientation-member)|Representa a propriedade `textOrientation`.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardheight-member)|Representa a propriedade `useStandardHeight`.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-usestandardwidth-member)|Representa a propriedade `useStandardWidth`.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-verticalalignment-member)|Representa a propriedade `verticalAlignment`.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformat#excel-excel-cellpropertiesformat-wraptext-member)|Representa a propriedade `wrapText`.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#excel-excel-cellpropertiesprotection-formulahidden-member)|Representa a propriedade `format.protection.formulaHidden`.|
||[bloqueado](/javascript/api/excel/excel.cellpropertiesprotection#excel-excel-cellpropertiesprotection-locked-member)|Representa a propriedade `format.protection.locked`.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valueafter-member)|Representa o valor após a alteração.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuebefore-member)|Representa o valor antes da alteração.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuetypeafter-member)|Representa o tipo de valor após a alteração.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#excel-excel-changedeventdetail-valuetypebefore-member)|Representa o tipo de valor antes da alteração.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#excel-excel-chart-activate-member(1))|Ativa o gráfico na interface do usuário do Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#excel-excel-chart-pivotoptions-member)|Encapsula as opções para um gráfico dinâmico.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-colorscheme-member)|Especifica o esquema de cores do gráfico.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#excel-excel-chartareaformat-roundedcorners-member)|Especifica se a área do gráfico do gráfico tem cantos arredondados.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-linknumberformat-member)|Especifica se o formato de número está vinculado às células.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-allowoverflow-member)|Especifica se o estouro de bin está habilitado em um gráfico de histograma ou gráfico de pareto.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-allowunderflow-member)|Especifica se o subfluxo da lixeira está habilitado em um gráfico de histograma ou gráfico de pareto.|
||[Count](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-count-member)|Especifica a contagem bin de um gráfico de histograma ou gráfico de pareto.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-overflowvalue-member)|Especifica o valor de estouro da lixeira de um gráfico de histograma ou gráfico de pareto.|
||[tipo](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-type-member)|Especifica o tipo da lixeira para um gráfico de histograma ou gráfico de pareto.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-underflowvalue-member)|Especifica o valor de subfluxo da lixeira de um gráfico de histograma ou gráfico de pareto.|
||[width](/javascript/api/excel/excel.chartbinoptions#excel-excel-chartbinoptions-width-member)|Especifica o valor de largura da lixeira de um gráfico de histograma ou gráfico de pareto.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-quartilecalculation-member)|Especifica se o tipo de cálculo quartil de uma caixa e um gráfico de whisker.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showinnerpoints-member)|Especifica se os pontos internos são mostrados em uma caixa e um gráfico de whisker.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showmeanline-member)|Especifica se a linha média é mostrada em uma caixa e um gráfico de whisker.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showmeanmarker-member)|Especifica se o marcador médio é mostrado em uma caixa e um gráfico de whisker.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#excel-excel-chartboxwhiskeroptions-showoutlierpoints-member)|Especifica se os pontos de outlier são mostrados em uma caixa e um gráfico de whisker.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-linknumberformat-member)|Especifica se o formato de número está vinculado às células (para que o formato de número mude nos rótulos quando ele muda nas células).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-linknumberformat-member)|Especifica se o formato de número está vinculado às células.|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-endstylecap-member)|Especifica se as barras de erro têm um limite de estilo final.|
||[format](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-format-member)|Especifica o tipo de formatação das barras de erro.|
||[include](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-include-member)|Especifica quais partes das barras de erro devem ser incluídas.|
||[tipo](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-type-member)|O tipo de intervalo marcado pelas barras de erro.|
||[visible](/javascript/api/excel/excel.charterrorbars#excel-excel-charterrorbars-visible-member)|Especifica se as barras de erro são exibidas.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#excel-excel-charterrorbarsformat-line-member)|Representa a formatação de linha do gráfico.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-labelstrategy-member)|Especifica a estratégia de rótulos de mapa de série de um gráfico de mapa de região.|
||[level](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-level-member)|Especifica o nível de mapeamento de série de um gráfico de mapa de região.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#excel-excel-chartmapoptions-projectiontype-member)|Especifica o tipo de projeção de série de um gráfico de mapa de região.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showaxisfieldbuttons-member)|Especifica se os botões do campo de eixo serão exibidos em um Gráfico Dinâmico.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showlegendfieldbuttons-member)|Especifica se os botões do campo de legenda serão exibidos em um Gráfico Dinâmico.|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showreportfilterfieldbuttons-member)|Especifica se os botões do campo de filtro de relatório serão exibidos em um Gráfico Dinâmico.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#excel-excel-chartpivotoptions-showvaluefieldbuttons-member)|Especifica se os botões do campo mostrar valor serão exibidos em um Gráfico Dinâmico.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[binOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-binoptions-member)|Encapsula as opções de bin para gráficos de histograma e gráficos de pareto.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-boxwhiskeroptions-member)|Encapsula as opções para os gráficos de caixa estreita.|
||[bubbleScale](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-bubblescale-member)|Este pode ser um valor inteiro de 0 (zero) a 300, representando a porcentagem do tamanho padrão.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumcolor-member)|Especifica a cor para o valor máximo de uma série de gráficos de mapa de região.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumtype-member)|Especifica o tipo para o valor máximo de uma série de gráficos de mapa de região.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmaximumvalue-member)|Especifica o valor máximo de uma série de gráficos de mapa de região.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointcolor-member)|Especifica a cor do valor do ponto médio de uma série de gráficos de mapa de região.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointtype-member)|Especifica o tipo para o valor do ponto médio de uma série de gráficos de mapa de região.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientmidpointvalue-member)|Especifica o valor do ponto médio de uma série de gráficos de mapa de região.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumcolor-member)|Especifica a cor do valor mínimo de uma série de gráficos de mapa de região.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumtype-member)|Especifica o tipo para o valor mínimo de uma série de gráficos de mapa de região.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientminimumvalue-member)|Especifica o valor mínimo de uma série de gráficos de mapa de região.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-gradientstyle-member)|Especifica o estilo de gradiente de série de um gráfico de mapa de região.|
||[invertColor](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-invertcolor-member)|Especifica a cor de preenchimento para pontos de dados negativos em uma série.|
||[mapOptions](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-mapoptions-member)|Encapsula as opções para um gráfico de mapa de região.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-parentlabelstrategy-member)|Especifica a área de estratégia de rótulo pai da série para um gráfico de mapa de árvore.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showconnectorlines-member)|Especifica se as linhas do conector são mostradas em gráficos de cascata.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-showleaderlines-member)|Especifica se as linhas de líder são exibidas para cada rótulo de dados na série.|
||[splitValue](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-splitvalue-member)|Especifica o valor limite que separa duas seções de um gráfico de pizza de pizza ou um gráfico de barras de pizza.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-xerrorbars-member)|Representa o objeto da barra de erros de uma série de gráficos.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-yerrorbars-member)|Representa o objeto da barra de erros de uma série de gráficos.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-linknumberformat-member)|Especifica se o formato de número está vinculado às células (para que o formato de número mude nos rótulos quando ele muda nas células).|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-address-member)|Representa a propriedade `address`.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-addresslocal-member)|Representa a propriedade `addressLocal`.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#excel-excel-columnproperties-columnindex-member)|Representa a propriedade `columnIndex`.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#excel-excel-conditionalformat-getranges-member(1))|Retorna o `RangeAreas`, compreendendo um ou mais intervalos retangulares, aos quais o formato conditonal é aplicado.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-getinvalidcells-member(1))|Retorna um `RangeAreas` objeto, compreendendo um ou mais intervalos retangulares, com valores de célula inválidos.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-getinvalidcellsornullobject-member(1))|Retorna um `RangeAreas` objeto, compreendendo um ou mais intervalos retangulares, com valores de célula inválidos.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#excel-excel-filtercriteria-subfield-member)|A propriedade usada pelo filtro para fazer um filtro rico em valores ricos.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#excel-excel-geometricshape-id-member)|Retorna o identificador de forma.|
||[shape](/javascript/api/excel/excel.geometricshape#excel-excel-geometricshape-shape-member)|Retorna o `Shape` objeto para a forma geométrica.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getcount-member(1))|Retorna o número de formas no grupo de forma.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitem-member(1))|Obtém uma forma usando seu nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitemat-member(1))|Obtém uma forma com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerfooter-member)|O rodapé central da planilha.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-centerheader-member)|O header central da planilha.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftfooter-member)|O rodapé esquerdo da planilha.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-leftheader-member)|O header esquerdo da planilha.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightfooter-member)|O rodapé direito da planilha.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#excel-excel-headerfooter-rightheader-member)|O header direito da planilha.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-defaultforallpages-member)|O cabeçalho/rodapé geral, usado em todas as páginas, a menos que seja especificada a página par/ímpar ou a primeira página.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-evenpages-member)|O cabeçalho/rodapé a ser usado para páginas pares, o cabeçalho/rodapé ímpar deve ser especificado para páginas ímpares.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-firstpage-member)|O cabeçalho/rodapé da primeira página. Para todas as outras páginas, geral ou par/ímpar é usado.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-oddpages-member)|O cabeçalho/rodapé a ser usado para páginas ímpares, o cabeçalho/rodapé par deve ser especificado para páginas pares.|
||[state](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-state-member)|O estado pelo qual os headers/rodapés estão definidos.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-usesheetmargins-member)|Obtém ou define um sinalizador indicando se os cabeçalhos/rodapés estão alinhados com as margens da página que foram definidas nas opções de layout de página da planilha.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#excel-excel-headerfootergroup-usesheetscale-member)|Obtém ou define um sinalizador que indica se os cabeçalhos/rodapés devem ser dimensionados pela escala de porcentagem da página definida nas opções de layout de página da planilha.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#excel-excel-image-format-member)|Retorna o formato da imagem.|
||[id](/javascript/api/excel/excel.image#excel-excel-image-id-member)|Especifica o identificador de forma do objeto image.|
||[shape](/javascript/api/excel/excel.image#excel-excel-image-shape-member)|Retorna o `Shape` objeto associado à imagem.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-enabled-member)|True se o Excel usará a interação para resolver referências circulares.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-maxchange-member)|Especifica a quantidade máxima de alteração entre cada iteração à medida que Excel resolve referências circulares.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#excel-excel-iterativecalculation-maxiteration-member)|Especifica o número máximo de iterações que Excel pode usar para resolver uma referência circular.|
|[Line](/javascript/api/excel/excel.line)|[beginArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadlength-member)|Representa o comprimento da ponta da seta no início da linha especificada.|
||[beginArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadstyle-member)|Representa o estilo da ponta de seta no início da linha especificada.|
||[BeginArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-beginarrowheadwidth-member)|Representa a largura da ponta da seta no início da linha especificada.|
||[beginConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedshape-member)|Representa a forma na qual o início da linha especificada está conectado.|
||[beginConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-beginconnectedsite-member)|Representa o site de conexão ao qual o início de um conector está conectado.|
||[connectBeginShape (forma: Excel.Shape, connectionSite: número)](/javascript/api/excel/excel.line#excel-excel-line-connectbeginshape-member(1))|Conecta o início do conector especificado a uma forma específica.|
||[connectEndShape (forma: Excel.Shape, connectionSite: número)](/javascript/api/excel/excel.line#excel-excel-line-connectendshape-member(1))|Anexa o final do conector especificado a uma forma específica.|
||[connectorType](/javascript/api/excel/excel.line#excel-excel-line-connectortype-member)|Representa o tipo de conector de linha.|
||[disconnectBeginShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectbeginshape-member(1))|Desconecta o início do conector especificado de uma forma.|
||[disconnectEndShape()](/javascript/api/excel/excel.line#excel-excel-line-disconnectendshape-member(1))|Desconecta o final do conector especificado de uma forma.|
||[endArrowheadLength](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadlength-member)|Representa o comprimento da ponta de seta no final da linha especificada.|
||[endArrowheadStyle](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadstyle-member)|Representa o estilo da ponta de seta no final da linha especificada.|
||[endArrowheadWidth](/javascript/api/excel/excel.line#excel-excel-line-endarrowheadwidth-member)|Representa a largura da ponta de seta no final da linha especificada.|
||[endConnectedShape](/javascript/api/excel/excel.line#excel-excel-line-endconnectedshape-member)|Representa a forma na qual o final da linha especificada está conectado.|
||[endConnectedSite](/javascript/api/excel/excel.line#excel-excel-line-endconnectedsite-member)|Representa o site de conexão ao qual o final de um conector está conectado.|
||[id](/javascript/api/excel/excel.line#excel-excel-line-id-member)|Especifica o identificador de forma.|
||[isBeginConnected](/javascript/api/excel/excel.line#excel-excel-line-isbeginconnected-member)|Especifica se o início da linha especificada está conectado a uma forma.|
||[isEndConnected](/javascript/api/excel/excel.line#excel-excel-line-isendconnected-member)|Especifica se o final da linha especificada está conectado a uma forma.|
||[shape](/javascript/api/excel/excel.line#excel-excel-line-shape-member)|Retorna o `Shape` objeto associado à linha.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[columnIndex](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-columnindex-member)|Especifica o índice de coluna para a quebra de página.|
||[delete()](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-delete-member(1))|Exclui um objeto de quebra de página.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-getcellafterbreak-member(1))|Obtém a primeira célula após a quebra de página.|
||[rowIndex](/javascript/api/excel/excel.pagebreak#excel-excel-pagebreak-rowindex-member)|Especifica o índice de linha para a quebra de página.|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-add-member(1))|Adiciona uma quebra de página antes da célula superior esquerda do intervalo especificado.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-getcount-member(1))|Obtém o número de quebras de página na coleção.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-getitem-member(1))|Obtém um objeto de quebra de página através do índice.|
||[items](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#excel-excel-pagebreakcollection-removepagebreaks-member(1))|Redefine todas as quebras de página manuais na coleção.|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-blackandwhite-member)|A opção de impressão em preto e branco da planilha.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-bottommargin-member)|A margem de página inferior da planilha a ser usada para impressão em pontos.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-centerhorizontally-member)|O sinalizador horizontal do centro da planilha.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-centervertically-member)|O sinalizador vertical do centro da planilha.|
||[draftMode](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-draftmode-member)|A opção de modo de rascunho da planilha.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-firstpagenumber-member)|O número da primeira página da planilha a ser impressa.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-footermargin-member)|A margem do rodapé da planilha, em pontos, para uso ao imprimir.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprintarea-member(1))|Obtém `RangeAreas` o objeto, compreendendo um ou mais intervalos retangulares, que representa a área de impressão da planilha.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprintareaornullobject-member(1))|Obtém `RangeAreas` o objeto, compreendendo um ou mais intervalos retangulares, que representa a área de impressão da planilha.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlecolumns-member(1))|Obtém o objeto range que representa as colunas de título.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlecolumnsornullobject-member(1))|Obtém o objeto range que representa as colunas de título.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlerows-member(1))|Obtém o objeto range representando as linhas do título.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-getprinttitlerowsornullobject-member(1))|Obtém o objeto range representando as linhas do título.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-headermargin-member)|A margem do header da planilha, em pontos, para uso ao imprimir.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-headersfooters-member)|Configuração de cabeçalho e rodapé da planilha.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-leftmargin-member)|A margem esquerda da planilha, em pontos, para uso ao imprimir.|
||[orientation](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-orientation-member)|A orientação da planilha da página.|
||[paperSize](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-papersize-member)|O tamanho do papel da página da planilha.|
||[printComments](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printcomments-member)|Especifica se os comentários da planilha devem ser exibidos durante a impressão.|
||[printErrors](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printerrors-member)|A opção erros de impressão da planilha.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printgridlines-member)|Especifica se as linhas de grade da planilha serão impressas.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printheadings-member)|Especifica se os títulos da planilha serão impressos.|
||[printOrder](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-printorder-member)|A opção de ordem de impressão de página da planilha.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-rightmargin-member)|A margem direita da planilha, em pontos, para uso ao imprimir.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprintarea-member(1))|Define a área de impressão da planilha.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprintmargins-member(1))|Define as margens das páginas da planilha com unidades.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprinttitlecolumns-member(1))|Define as colunas que contêm as células que serão repetidas à esquerda de cada página da planilha para impressão.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-setprinttitlerows-member(1))|Define as linhas que contêm as células que serão repetidas na parte de cada página da planilha para impressão.|
||[topMargin](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-topmargin-member)|A margem superior da planilha, em pontos, para uso ao imprimir.|
||[zoom](/javascript/api/excel/excel.pagelayout#excel-excel-pagelayout-zoom-member)|As opções de zoom de impressão da planilha.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-bottom-member)|Especifica a margem inferior do layout da página na unidade especificada para ser usada para impressão.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-footer-member)|Especifica a margem do rodapé de layout de página na unidade especificada para ser usada para impressão.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-header-member)|Especifica a margem do header de layout da página na unidade especificada para ser usada para impressão.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-left-member)|Especifica a margem esquerda do layout da página na unidade especificada para ser usada para impressão.|
||[direita](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-right-member)|Especifica a margem direita do layout da página na unidade especificada para ser usada para impressão.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#excel-excel-pagelayoutmarginoptions-top-member)|Especifica a margem superior do layout da página na unidade especificada para ser usada para impressão.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-horizontalfittopages-member)|Número de páginas a ser horizontalmente ajustado.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-scale-member)|O valor do dimensionamento da página de impressão pode estar entre 10 e 400.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#excel-excel-pagelayoutzoomoptions-verticalfittopages-member)|Número de páginas a ser verticalmente ajustado.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbyvalues-member(1))|Classifica o Campo dinâmico por valores especificados em um determinado escopo.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-autoformat-member)|Especifica se a formatação será formatada automaticamente quando for atualizada ou quando os campos são movidos.|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatahierarchy-member(1))|Obtém o DataHierarchy que é usado para calcular o valor em um intervalo especificado dentro da Tabela dinâmica.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getpivotitems-member(1))|Obtém os Itens dinâmicos de um eixo que compõem o valor em um intervalo especificado dentro da Tabela dinâmica.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-preserveformatting-member)|Especifica se a formatação é preservada quando o relatório é atualizado ou recalculado por operações como pivoting, classificação ou alteração de itens de campo de página.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-setautosortoncell-member(1))|Define a Tabela Dinâmica para classificar automaticamente usando a célula especificada para selecionar automaticamente todos os critérios e contextos necessários.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-enabledatavalueediting-member)|Especifica se a Tabela Dinâmica permite que os valores no corpo dos dados sejam editados pelo usuário.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-usecustomsortlists-member)|Especifica se a Tabela Dinâmica usa listas personalizadas ao classificar.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange?: Cadeia de caracteres \| de intervalo, autoFillType?: Excel. AutoFillType)](/javascript/api/excel/excel.range#excel-excel-range-autofill-member(1))|Preenche o intervalo do intervalo atual até o intervalo de destino usando a lógica AutoFill especificada.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#excel-excel-range-convertdatatypetotext-member(1))|Converte as células de intervalo com tipos de dados em texto.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#excel-excel-range-converttolinkeddatatype-member(1))|Converte as células de intervalo em tipos de dados vinculados na planilha.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-copyfrom-member(1))|Copia dados da célula ou formatação do intervalo de origem ou `RangeAreas` do intervalo atual.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#excel-excel-range-find-member(1))|Localiza certa cadeia de caracteres com base em critérios especificados.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#excel-excel-range-findornullobject-member(1))|Localiza certa cadeia de caracteres com base em critérios especificados.|
||[flashFill()](/javascript/api/excel/excel.range#excel-excel-range-flashfill-member(1))|Faz um Preenchimento Flash no intervalo atual.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getcellproperties-member(1))|Retorna uma matriz 2D encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada célula.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getcolumnproperties-member(1))|Retorna uma única matriz dimensional encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada coluna.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#excel-excel-range-getrowproperties-member(1))|Retorna uma única matriz dimensional encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada célula.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcells-member(1))|Obtém `RangeAreas` o objeto, compreendendo um ou mais intervalos retangulares, que representa todas as células que corresponderem ao tipo e ao valor especificados.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#excel-excel-range-getspecialcellsornullobject-member(1))|Obtém `RangeAreas` o objeto, compreendendo um ou mais intervalos, que representa todas as células que corresponderem ao tipo e ao valor especificados.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-gettables-member(1))|Obtém uma coleção de tabelas com escopo que se sobrepõe ao intervalo.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#excel-excel-range-linkeddatatypestate-member)|Representa o estado do tipo de dados de cada célula.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1))|Remove valores duplicados do intervalo especificado pelas colunas.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#excel-excel-range-replaceall-member(1))|Localiza e substitui a cadeia de caracteres fornecida com base nos critérios especificados no intervalo atual.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#excel-excel-range-setcellproperties-member(1))|Atualiza o intervalo com base em uma matriz 2D de propriedades de células, encapsulando coisas como fonte, preenchimento, bordas e alinhamento.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setcolumnproperties-member(1))|Atualiza o intervalo com base em uma matriz unidimensional de propriedades de coluna, encapsulando coisas como fonte, preenchimento, bordas e alinhamento.|
||[setDirty()](/javascript/api/excel/excel.range#excel-excel-range-setdirty-member(1))|Define um intervalo a ser recalculado quando o próximo recálculo ocorrer.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#excel-excel-range-setrowproperties-member(1))|Atualiza o intervalo com base em uma matriz unidimensional de propriedades de linha, encapsulando coisas como fonte, preenchimento, bordas e alinhamento.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[address](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-address-member)|Retorna a `RangeAreas` referência no estilo A1.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-addresslocal-member)|Retorna a `RangeAreas` referência na localidade do usuário.|
||[areaCount](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-areacount-member)|Retorna o número de intervalos retangulares que compõem esse `RangeAreas` objeto.|
||[areas](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-areas-member)|Retorna uma coleção de intervalos retangulares que compõem esse `RangeAreas` objeto.|
||[calculate()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-calculate-member(1))|Calcula todas as células no `RangeAreas`.|
||[cellCount](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-cellcount-member)|Retorna o número de células no `RangeAreas` objeto, somando as contagens de células de todos os intervalos retangulares individuais.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-clear-member(1))|Limpa valores, formato, preenchimento, borda e outras propriedades em cada uma das áreas que compõem esse `RangeAreas` objeto.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-conditionalformats-member)|Retorna uma coleção de formatos condicionais que se cruzam com qualquer célula neste `RangeAreas` objeto.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-convertdatatypetotext-member(1))|Converte todas as células no com tipos `RangeAreas` de dados em texto.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-converttolinkeddatatype-member(1))|Converte todas as células nos tipos `RangeAreas` de dados vinculados.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-copyfrom-member(1))|Copia dados da célula ou formatação do intervalo de origem ou `RangeAreas` para o `RangeAreas`atual .|
||[dataValidation](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-datavalidation-member)|Retorna um objeto de validação de dados para todos os intervalos no `RangeAreas`.|
||[format](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-format-member)|Retorna um `RangeFormat` objeto, encapsulando a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades para todos os intervalos do `RangeAreas` objeto.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getentirecolumn-member(1))|Retorna um objeto que representa as colunas inteiras `RangeAreas` do (por exemplo, `RangeAreas` se a `RangeAreas` atual representa as células "B4:E11, H2", `RangeAreas` ele retorna uma que representa colunas "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getentirerow-member(1))|Retorna um objeto que representa as linhas inteiras `RangeAreas` do (por exemplo, `RangeAreas` se a `RangeAreas` atual representa as células "B4:E11", `RangeAreas` ele retorna um que representa linhas "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getintersection-member(1))|Retorna o `RangeAreas` objeto que representa a interseção dos intervalos determinados ou `RangeAreas`.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getintersectionornullobject-member(1))|Retorna o `RangeAreas` objeto que representa a interseção dos intervalos determinados ou `RangeAreas`.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getoffsetrangeareas-member(1))|Retorna um `RangeAreas` objeto que é deslocado pelo deslocamento de linha e coluna específico.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getspecialcells-member(1))|Retorna um `RangeAreas` objeto que representa todas as células que corresponderem ao tipo e ao valor especificados.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getspecialcellsornullobject-member(1))|Retorna um `RangeAreas` objeto que representa todas as células que corresponderem ao tipo e ao valor especificados.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-gettables-member(1))|Retorna uma coleção de tabelas com escopo que se sobrepõem a qualquer intervalo neste `RangeAreas` objeto.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getusedrangeareas-member(1))|Retorna o usado `RangeAreas` que compreende todas as áreas usadas de intervalos retangulares individuais no `RangeAreas` objeto.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-getusedrangeareasornullobject-member(1))|Retorna o usado `RangeAreas` que compreende todas as áreas usadas de intervalos retangulares individuais no `RangeAreas` objeto.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-isentirecolumn-member)|Especifica se todos os intervalos `RangeAreas` neste objeto representam colunas inteiras (por exemplo, "A:C, P:Z").|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-isentirerow-member)|Especifica se todos os intervalos `RangeAreas` neste objeto representam linhas inteiras (por exemplo, "1:3, 5:7").|
||[setDirty()](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-setdirty-member(1))|Define o `RangeAreas` a ser recalculado quando o próximo recálculo ocorrer.|
||[style](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-style-member)|Representa o estilo de todos os intervalos neste `RangeAreas` objeto.|
||[worksheet](/javascript/api/excel/excel.rangeareas#excel-excel-rangeareas-worksheet-member)|Retorna a planilha para o `RangeAreas`atual .|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#excel-excel-rangeborder-tintandshade-member)|Especifica um duplo que clareia ou escurece uma cor para a borda do intervalo, o valor está entre -1 (mais escuro) e 1 (mais brilhante), com 0 para a cor original.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#excel-excel-rangebordercollection-tintandshade-member)|Especifica um duplo que clareia ou escurece uma cor para bordas de intervalo.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-getcount-member(1))|Retorna o número de intervalos no `RangeCollection`.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-getitemat-member(1))|Retorna o objeto range com base em sua posição no `RangeCollection`.|
||[items](/javascript/api/excel/excel.rangecollection#excel-excel-rangecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[padrão](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-pattern-member)|O padrão de um intervalo.|
||[patternColor](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-patterncolor-member)|O código de cor HTML que representa a cor do padrão de intervalo, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-patterntintandshade-member)|Especifica um duplo que clareia ou escurece uma cor de padrão para o preenchimento do intervalo.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#excel-excel-rangefill-tintandshade-member)|Especifica um duplo que clareia ou escurece uma cor para o preenchimento do intervalo.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-strikethrough-member)|Especifica o status tachado da fonte.|
||[subscript](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-subscript-member)|Especifica o status de subscrito da fonte.|
||[superscript](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-superscript-member)|Especifica o status sobrescrito da fonte.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#excel-excel-rangefont-tintandshade-member)|Especifica um duplo que clareia ou escurece uma cor para a fonte de intervalo.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-autoindent-member)|Especifica se o texto é recuado automaticamente quando o alinhamento de texto é definido como distribuição igual.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-indentlevel-member)|Um número inteiro entre 0 e 250 que indica o nível de recuo.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-readingorder-member)|A ordem de leitura para o intervalo.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#excel-excel-rangeformat-shrinktofit-member)|Especifica se o texto reduz automaticamente para caber na largura da coluna disponível.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#excel-excel-removeduplicatesresult-removed-member)|Número de linhas duplicadas removidas pela operação.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#excel-excel-removeduplicatesresult-uniqueremaining-member)|Número de linhas restantes exclusivas presentes no intervalo resultante.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#excel-excel-replacecriteria-completematch-member)|Especifica se a combinação precisa ser completa ou parcial.|
||[matchCase](/javascript/api/excel/excel.replacecriteria#excel-excel-replacecriteria-matchcase-member)|Especifica se a combinação é sensível a minúsculas.|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-address-member)|Representa a propriedade `address`.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-addresslocal-member)|Representa a propriedade `addressLocal`.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#excel-excel-rowproperties-rowindex-member)|Representa a propriedade `rowIndex`.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-completematch-member)|Especifica se a combinação precisa ser completa ou parcial.|
||[matchCase](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-matchcase-member)|Especifica se a combinação é sensível a minúsculas.|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#excel-excel-searchcriteria-searchdirection-member)|Especifica a direção da pesquisa.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-format-member)|Representa a propriedade `format`.|
||[hiperlink](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-hyperlink-member)|Representa a propriedade `hyperlink`.|
||[style](/javascript/api/excel/excel.settablecellproperties#excel-excel-settablecellproperties-style-member)|Representa a propriedade `style`.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-columnhidden-member)|Representa a propriedade `columnHidden`.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-columnwidth-member)||
||[format: Excel. CellPropertiesFormat & { columnWidth](/javascript/api/excel/excel.settablecolumnproperties#excel-excel-settablecolumnproperties-format-member)|Representa a propriedade `format`.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[format: Excel. CellPropertiesFormat & { rowHeight](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-format-member)|Representa a propriedade `format`.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-rowheight-member)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#excel-excel-settablerowproperties-rowhidden-member)|Representa a propriedade `rowHidden`.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#excel-excel-shape-alttextdescription-member)|Especifica o texto de descrição alternativo para um `Shape` objeto.|
||[altTextTitle](/javascript/api/excel/excel.shape#excel-excel-shape-alttexttitle-member)|Especifica o texto de título alternativo para um `Shape` objeto.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#excel-excel-shape-connectionsitecount-member)|Retorna o número de locais de conexão nessa forma.|
||[delete()](/javascript/api/excel/excel.shape#excel-excel-shape-delete-member(1))|Remove a forma da planilha.|
||[fill](/javascript/api/excel/excel.shape#excel-excel-shape-fill-member)|Retorna a formatação de preenchimento dessa forma.|
||[geometricShape](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshape-member)|Retorna a forma geométrica associada à forma.|
||[geometricShapeType](/javascript/api/excel/excel.shape#excel-excel-shape-geometricshapetype-member)|Especifica o tipo de forma geométrica dessa forma geométrica.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#excel-excel-shape-getasimage-member(1))|Converte a forma em uma imagem e retorna a imagem como uma cadeia de caracteres de base 64.|
||[group](/javascript/api/excel/excel.shape#excel-excel-shape-group-member)|Retorna o grupo de forma associado à forma.|
||[height](/javascript/api/excel/excel.shape#excel-excel-shape-height-member)|Especifica a altura, em pontos, da forma.|
||[id](/javascript/api/excel/excel.shape#excel-excel-shape-id-member)|Especifica o identificador de forma.|
||[image](/javascript/api/excel/excel.shape#excel-excel-shape-image-member)|Retorna a imagem associada à forma.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementleft-member(1))|Move a forma horizontalmente pelo número especificado de pontos.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementrotation-member(1))|O formato é girado em sentido horário ao redor do eixo z pelo número especificado de graus.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#excel-excel-shape-incrementtop-member(1))|Move a forma verticalmente pelo número especificado de pontos.|
||[left](/javascript/api/excel/excel.shape#excel-excel-shape-left-member)|A distância, em pontos, da lateral esquerda da forma do lado  esquerdo da planilha.|
||[level](/javascript/api/excel/excel.shape#excel-excel-shape-level-member)|Especifica o nível da forma especificada.|
||[line](/javascript/api/excel/excel.shape#excel-excel-shape-line-member)|Retorna a linha associada à forma.|
||[lineFormat](/javascript/api/excel/excel.shape#excel-excel-shape-lineformat-member)|Retorna a formatação de linha do objeto de forma.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#excel-excel-shape-lockaspectratio-member)|Especifica se a proporção dessa forma está bloqueada.|
||[name](/javascript/api/excel/excel.shape#excel-excel-shape-name-member)|Especifica o nome da forma.|
||[onActivated](/javascript/api/excel/excel.shape#excel-excel-shape-onactivated-member)|Ocorre quando a forma é ativada.|
||[onDeactivated](/javascript/api/excel/excel.shape#excel-excel-shape-ondeactivated-member)|Ocorre quando a forma é desativada.|
||[parentGroup](/javascript/api/excel/excel.shape#excel-excel-shape-parentgroup-member)|Especifica o grupo pai dessa forma.|
||[rotation](/javascript/api/excel/excel.shape#excel-excel-shape-rotation-member)|Especifica a rotação, em graus, da forma.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scaleheight-member(1))|Dimensiona a altura da forma por um fator especificado.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#excel-excel-shape-scalewidth-member(1))|Dimensiona a largura da forma por um fator especificado.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#excel-excel-shape-setzorder-member(1))|Move a forma especificada para cima ou para baixo na ordem z da coleção, que a desloca para frente ou para trás de outras formas.|
||[textFrame](/javascript/api/excel/excel.shape#excel-excel-shape-textframe-member)|Retorna o objeto text frame de uma forma.|
||[top](/javascript/api/excel/excel.shape#excel-excel-shape-top-member)|A distância, em pontos, da borda superior da forma até a borda superior da planilha.|
||[tipo](/javascript/api/excel/excel.shape#excel-excel-shape-type-member)|Retorna o tipo dessa forma.|
||[visible](/javascript/api/excel/excel.shape#excel-excel-shape-visible-member)|Especifica se a forma está visível.|
||[width](/javascript/api/excel/excel.shape#excel-excel-shape-width-member)|Especifica a largura, em pontos, da forma.|
||[zOrderPosition](/javascript/api/excel/excel.shape#excel-excel-shape-zorderposition-member)|Retorna a posição da forma especificada na ordem z, com 0 representando a parte inferior da pilha do pedido.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-shapeid-member)|Obtém a ID da forma ativada.|
||[tipo](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#excel-excel-shapeactivatedeventargs-worksheetid-member)|Obtém a ID da planilha na qual a forma é ativada.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgeometricshape-member(1))|Adiciona uma forma geométrica à planilha.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addgroup-member(1))|Um subconjunto de formas na planilha do conjunto de grupos.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addimage-member(1))|Cria uma imagem de uma cadeia de caracteres na base 64 e a adiciona à planilha.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addline-member(1))|Adiciona uma linha à planilha.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-addtextbox-member(1))|Adiciona uma caixa de texto na planilha com o texto fornecido como conteúdo.|
||[getCount()](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getcount-member(1))|Retorna o número de formas da planilha.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitem-member(1))|Obtém uma forma usando seu nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitemat-member(1))|Obtém uma forma usando sua posição na coleção.|
||[items](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-shapeid-member)|Obtém a ID da forma desativada.|
||[tipo](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#excel-excel-shapedeactivatedeventargs-worksheetid-member)|Obtém a ID da planilha na qual a forma é desativada.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-clear-member(1))|Limpa a formatação do preenchimento de um objeto de forma.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-foregroundcolor-member)|Representa a cor de primeiro plano de preenchimento da forma no formato de cor HTML, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-setsolidcolor-member(1))|Define a formatação de preenchimento de um formato com uma cor uniforme.|
||[transparency](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-transparency-member)|Especifica a porcentagem de transparência do preenchimento como um valor de 0,0 (opaco) a 1,0 (claro).|
||[tipo](/javascript/api/excel/excel.shapefill#excel-excel-shapefill-type-member)|Retorna o tipo de preenchimento da forma.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-bold-member)|Representa o status da fonte em negrito.|
||[color](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-color-member)|Representação de código de cor HTML da cor do texto (por exemplo, "#FF0000" representa vermelho).|
||[italic](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-italic-member)|Representa o status da fonte em itálico.|
||[name](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-name-member)|Representa o nome da fonte (por exemplo, "Calibri").|
||[size](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-size-member)|Representa o tamanho da fonte em pontos (por exemplo, 11).|
||[underline](/javascript/api/excel/excel.shapefont#excel-excel-shapefont-underline-member)|Tipo de sublinhado aplicado à fonte.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-id-member)|Especifica o identificador de forma.|
||[shape](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-shape-member)|Retorna o `Shape` objeto associado ao grupo.|
||[shapes](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-shapes-member)|Retorna a coleção de `Shape` objetos.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#excel-excel-shapegroup-ungroup-member(1))|Desagrupa todas as formas agrupadas no grupo de forma especificado.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-color-member)|Representa a cor da linha no formato de cor HTML, no formato #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-dashstyle-member)|Representa o estilo de linha da forma.|
||[style](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-style-member)|Representa o estilo de linha da forma.|
||[transparency](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-transparency-member)|Representa o grau de transparência da linha especificada como um valor de 0,0 (opaco) a 1,0 (claro).|
||[visible](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-visible-member)|Especifica se a formatação de linha de um elemento de forma está visível.|
||[weight](/javascript/api/excel/excel.shapelineformat#excel-excel-shapelineformat-weight-member)|Representa a espessura da linha, em pontos.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#excel-excel-sortfield-subfield-member)|Especifica o subcampo que é o nome da propriedade de destino de um valor rico a ser classificação.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getcount-member(1))|Obtém o número de estilos na coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitemat-member(1))|Obtém um estilo com base em sua posição na coleção.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#excel-excel-table-autofilter-member)|Representa o `AutoFilter` objeto da tabela.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-source-member)|Obtém a origem do evento.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-tableid-member)|Obtém a ID da tabela adicionada.|
||[tipo](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#excel-excel-tableaddedeventargs-worksheetid-member)|Obtém a ID da planilha na qual a tabela é adicionada.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[detalhes](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-details-member)|Obtém as informações sobre os detalhes da alteração.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-onadded-member)|Ocorre quando uma nova tabela é adicionada a uma workbook.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-ondeleted-member)|Ocorre quando a tabela especificada é excluída em uma pasta de trabalho.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-source-member)|Obtém a origem do evento.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-tableid-member)|Obtém a ID da tabela excluída.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-tablename-member)|Obtém o nome da tabela excluída.|
||[tipo](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#excel-excel-tabledeletedeventargs-worksheetid-member)|Obtém a ID da planilha na qual a tabela é excluída.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getcount-member(1))|Obtém o número de tabelas na coleção.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getfirst-member(1))|Obtém a primeira tabela na coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getitem-member(1))|Obtém uma tabela pelo nome ou ID.|
||[items](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#excel-excel-textframe-autosizesetting-member)|As configurações de redação automáticas do quadro de texto.|
||[bottomMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-bottommargin-member)|Representa margem inferior, em pontos, do quadro de texto.|
||[deleteText()](/javascript/api/excel/excel.textframe#excel-excel-textframe-deletetext-member(1))|Exclui todo o texto no quadro de texto.|
||[hasText](/javascript/api/excel/excel.textframe#excel-excel-textframe-hastext-member)|Especifica se o quadro de texto contém texto.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#excel-excel-textframe-horizontalalignment-member)|Representa o alinhamento horizontal do quadro de texto.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#excel-excel-textframe-horizontaloverflow-member)|Representa o comportamento de excedente horizontal do quadro de texto.|
||[leftMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-leftmargin-member)|Representa margem esquerda, em pontos, do quadro de texto.|
||[orientation](/javascript/api/excel/excel.textframe#excel-excel-textframe-orientation-member)|Representa o ângulo para o qual o texto é orientado para o quadro de texto.|
||[readingOrder](/javascript/api/excel/excel.textframe#excel-excel-textframe-readingorder-member)|Representa a ordem de leitura do quadro de texto, da direita para a esquerda ou da direita para a esquerda.|
||[rightMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-rightmargin-member)|Representa margem direita, em pontos, do quadro de texto.|
||[textRange](/javascript/api/excel/excel.textframe#excel-excel-textframe-textrange-member)|Representa o texto que está anexado a uma forma, bem como propriedades e métodos para manipular o texto.|
||[topMargin](/javascript/api/excel/excel.textframe#excel-excel-textframe-topmargin-member)|Representa margem superior, em pontos, do quadro de texto.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#excel-excel-textframe-verticalalignment-member)|Representa o alinhamento vertical do quadro de texto.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#excel-excel-textframe-verticaloverflow-member)|Representa o comportamento de excedente vertical do quadro de texto.|
|[TextRange](/javascript/api/excel/excel.textrange)|[font](/javascript/api/excel/excel.textrange#excel-excel-textrange-font-member)|Retorna um `ShapeFont` objeto que representa os atributos de fonte para o intervalo de texto.|
||[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#excel-excel-textrange-getsubstring-member(1))|Retorna um objeto TextRange para a subcadeia de caracteres no intervalo especificado.|
||[text](/javascript/api/excel/excel.textrange#excel-excel-textrange-text-member)|Representa o conteúdo de texto sem formatação do intervalo de texto.|
|[Workbook](/javascript/api/excel/excel.workbook)|[autoSave](/javascript/api/excel/excel.workbook#excel-excel-workbook-autosave-member)|Especifica se a workbook está no modo AutoSave.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#excel-excel-workbook-calculationengineversion-member)|Retorna um número sobre a versão do Mecanismo de Cálculo do Excel.|
||[chartDataPointTrack](/javascript/api/excel/excel.workbook#excel-excel-workbook-chartdatapointtrack-member)|True se todos os gráficos na pasta de trabalho estiverem rastreando os pontos de dados reais aos quais eles estão anexados.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechart-member(1))|Obtém o gráfico ativo no momento na pasta de trabalho.|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactivechartornullobject-member(1))|Obtém o gráfico ativo no momento na pasta de trabalho.|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getisactivecollabsession-member(1))|Retorna `true` se a workbook estiver sendo editada por vários usuários (por meio de coautor).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedranges-member(1))|Obtém um ou mais intervalos atualmente selecionados da pasta de trabalho.|
||[isDirty](/javascript/api/excel/excel.workbook#excel-excel-workbook-isdirty-member)|Especifica se as alterações foram feitas desde a última vez que a workbook foi salva.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#excel-excel-workbook-onautosavesettingchanged-member)|Ocorre quando a configuração AutoSave é alterada na manual de trabalho.|
||[previouslySaved](/javascript/api/excel/excel.workbook#excel-excel-workbook-previouslysaved-member)|Especifica se a workbook já foi salva localmente ou online.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#excel-excel-workbook-useprecisionasdisplayed-member)|True se os cálculos dessa pasta de trabalho forem efetuados usando apenas a precisão dos números conforme forem exibidos.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[tipo](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#excel-excel-workbookautosavesettingchangedeventargs-type-member)|Obtém o tipo do evento.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[autoFilter](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-autofilter-member)|Representa o `AutoFilter` objeto da planilha.|
||[enableCalculation](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-enablecalculation-member)|Determina se Excel deve recalcular a planilha quando necessário.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findall-member(1))|Localiza todas as ocorrências da `RangeAreas` cadeia de caracteres determinada com base nos critérios especificados e retorna-as como um objeto, compreendendo um ou mais intervalos retangulares.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-findallornullobject-member(1))|Localiza todas as ocorrências da `RangeAreas` cadeia de caracteres determinada com base nos critérios especificados e retorna-as como um objeto, compreendendo um ou mais intervalos retangulares.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getranges-member(1))|Obtém `RangeAreas` o objeto, representando um ou mais blocos de intervalos retangulares, especificados pelo endereço ou nome.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-horizontalpagebreaks-member)|Obtém a coleção de quebra de página horizontal da planilha.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformatchanged-member)|Ocorre quando o formato é alterado em uma planilha específica.|
||[pageLayout](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-pagelayout-member)|Obtém `PageLayout` o objeto da planilha.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-replaceall-member(1))|Localiza e substitui a cadeia de caracteres fornecida com base nos critérios especificados na planilha atual.|
||[shapes](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-shapes-member)|Retorna a coleção de todos os objetos Shape na planilha.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-verticalpagebreaks-member)|Obtém a coleção de quebra de página vertical da planilha.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[detalhes](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-details-member)|Representa as informações sobre os detalhes da alteração.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onchanged-member)|Ocorre quando uma planilha da pasta de trabalho é alterada.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformatchanged-member)|Ocorre quando qualquer planilha na pasta de trabalho tem um formato alterado.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onselectionchanged-member)|Ocorre quando a seleção é alterada em uma planilha.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-address-member)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrange-member(1))|Obtém o intervalo que representa a área alterada de uma planilha específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-getrangeornullobject-member(1))|Obtém o intervalo que representa a área alterada de uma planilha específica.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-source-member)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#excel-excel-worksheetformatchangedeventargs-worksheetid-member)|Obtém a ID da planilha na qual os dados foram alterados.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#excel-excel-worksheetsearchcriteria-completematch-member)|Especifica se a combinação precisa ser completa ou parcial.|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#excel-excel-worksheetsearchcriteria-matchcase-member)|Especifica se a combinação é sensível a minúsculas.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.9&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
