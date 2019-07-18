---
title: Conjunto de requisitos de API JavaScript do Excel 1,9
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,9
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1c7361debe7ba09c3477d39d9337c35bf5df3066
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771999"
---
# <a name="whats-new-in-excel-javascript-api-19"></a>O que há de novo na API JavaScript do Excel 1.9

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

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[calculationEngineVersion](/javascript/api/excel/excel.application#calculationengineversion)|Retorna a versão do mecanismo de cálculo do Excel usada para o último recálculo completo. Somente leitura.|
||[calculationState](/javascript/api/excel/excel.application#calculationstate)|Retorna o estado de cálculo do aplicativo. Para saber detalhes, confira Excel.CalculationState. Somente leitura.|
||[iterativeCalculation](/javascript/api/excel/excel.application#iterativecalculation)|Retorna as configurações do Cálculo iterativo.|
||[suspendScreenUpdatingUntilNextSync()](/javascript/api/excel/excel.application#suspendscreenupdatinguntilnextsync--)|Suspende a atualização da tela até que o próximo "context.sync()" seja chamado.|
|[ApplicationData](/javascript/api/excel/excel.applicationdata)|[calculationEngineVersion](/javascript/api/excel/excel.applicationdata#calculationengineversion)|Retorna a versão do mecanismo de cálculo do Excel usada para o último recálculo completo. Somente leitura.|
||[calculationState](/javascript/api/excel/excel.applicationdata#calculationstate)|Retorna o estado de cálculo do aplicativo. Para saber detalhes, confira Excel.CalculationState. Somente leitura.|
||[iterativeCalculation](/javascript/api/excel/excel.applicationdata#iterativecalculation)|Retorna as configurações do Cálculo iterativo.|
|[ApplicationLoadOptions](/javascript/api/excel/excel.applicationloadoptions)|[calculationEngineVersion](/javascript/api/excel/excel.applicationloadoptions#calculationengineversion)|Retorna a versão do mecanismo de cálculo do Excel usada para o último recálculo completo. Somente leitura.|
||[calculationState](/javascript/api/excel/excel.applicationloadoptions#calculationstate)|Retorna o estado de cálculo do aplicativo. Para saber detalhes, confira Excel.CalculationState. Somente leitura.|
||[iterativeCalculation](/javascript/api/excel/excel.applicationloadoptions#iterativecalculation)|Retorna as configurações do Cálculo iterativo.|
|[ApplicationUpdateData](/javascript/api/excel/excel.applicationupdatedata)|[iterativeCalculation](/javascript/api/excel/excel.applicationupdatedata#iterativecalculation)|Retorna as configurações do Cálculo iterativo.|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[apply(range: Range \| string, columnIndex?: number, criteria?: Excel.FilterCriteria)](/javascript/api/excel/excel.autofilter#apply-range--columnindex--criteria-)|Aplica o AutoFiltro a um intervalo. Isso filtra a coluna se o índice de coluna e os critérios de filtro forem especificados.|
||[clearCriteria()](/javascript/api/excel/excel.autofilter#clearcriteria--)|Limpa os critérios de filtro do AutoFiltro.|
||[getRange()](/javascript/api/excel/excel.autofilter#getrange--)|Retorna um objeto Range que representa o intervalo no qual o Filtro automático se aplica.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.autofilter#getrangeornullobject--)|Retorna um objeto Range que representa o intervalo no qual o Filtro automático se aplica.|
||[criteria](/javascript/api/excel/excel.autofilter#criteria)|Uma matriz que contém todos os critérios de filtro no intervalo de autofiltro. Somente Leitura.|
||[enabled](/javascript/api/excel/excel.autofilter#enabled)|Indica se o Filtro automático está ativado ou não. Somente Leitura.|
||[isDataFiltered](/javascript/api/excel/excel.autofilter#isdatafiltered)|Indica se o Filtro automático tem critérios de filtro. Somente Leitura.|
||[reapply()](/javascript/api/excel/excel.autofilter#reapply--)|Aplica o objeto Autofilter especificado que está atualmente no intervalo.|
||[remove()](/javascript/api/excel/excel.autofilter#remove--)|Remove o Filtro automático do intervalo.|
|[AutoFilterData](/javascript/api/excel/excel.autofilterdata)|[criteria](/javascript/api/excel/excel.autofilterdata#criteria)|Uma matriz que contém todos os critérios de filtro no intervalo de autofiltro. Somente Leitura.|
||[enabled](/javascript/api/excel/excel.autofilterdata#enabled)|Indica se o Filtro automático está ativado ou não. Somente Leitura.|
||[isDataFiltered](/javascript/api/excel/excel.autofilterdata#isdatafiltered)|Indica se o Filtro automático tem critérios de filtro. Somente Leitura.|
|[AutoFilterLoadOptions](/javascript/api/excel/excel.autofilterloadoptions)|[$all](/javascript/api/excel/excel.autofilterloadoptions#$all)||
||[criteria](/javascript/api/excel/excel.autofilterloadoptions#criteria)|Uma matriz que contém todos os critérios de filtro no intervalo de autofiltro. Somente Leitura.|
||[enabled](/javascript/api/excel/excel.autofilterloadoptions#enabled)|Indica se o Filtro automático está ativado ou não. Somente Leitura.|
||[isDataFiltered](/javascript/api/excel/excel.autofilterloadoptions#isdatafiltered)|Indica se o Filtro automático tem critérios de filtro. Somente Leitura.|
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
|[CellPropertiesBorderLoadOptions](/javascript/api/excel/excel.cellpropertiesborderloadoptions)|[color](/javascript/api/excel/excel.cellpropertiesborderloadoptions#color)|Especifica se a `color` propriedade deve ser carregada.|
||[style](/javascript/api/excel/excel.cellpropertiesborderloadoptions#style)|Especifica se a `style` propriedade deve ser carregada.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesborderloadoptions#tintandshade)|Especifica se a `tintAndShade` propriedade deve ser carregada.|
||[weight](/javascript/api/excel/excel.cellpropertiesborderloadoptions#weight)|Especifica se a `weight` propriedade deve ser carregada.|
|[CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)|[color](/javascript/api/excel/excel.cellpropertiesfill#color)|Representa a propriedade `format.fill.color`.|
||[padrão](/javascript/api/excel/excel.cellpropertiesfill#pattern)|Representa a propriedade `format.fill.pattern`.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfill#patterncolor)|Representa a propriedade `format.fill.patternColor`.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfill#patterntintandshade)|Representa a propriedade `format.fill.patternTintAndShade`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfill#tintandshade)|Representa a propriedade `format.fill.tintAndShade`.|
|[CellPropertiesFillLoadOptions](/javascript/api/excel/excel.cellpropertiesfillloadoptions)|[color](/javascript/api/excel/excel.cellpropertiesfillloadoptions#color)|Especifica se a `color` propriedade deve ser carregada.|
||[padrão](/javascript/api/excel/excel.cellpropertiesfillloadoptions#pattern)|Especifica se a `pattern` propriedade deve ser carregada.|
||[patternColor](/javascript/api/excel/excel.cellpropertiesfillloadoptions#patterncolor)|Especifica se a `patternColor` propriedade deve ser carregada.|
||[patternTintAndShade](/javascript/api/excel/excel.cellpropertiesfillloadoptions#patterntintandshade)|Especifica se a `patternTintAndShade` propriedade deve ser carregada.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfillloadoptions#tintandshade)|Especifica se a `tintAndShade` propriedade deve ser carregada.|
|[CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)|[bold](/javascript/api/excel/excel.cellpropertiesfont#bold)|Representa a propriedade `format.font.bold`.|
||[color](/javascript/api/excel/excel.cellpropertiesfont#color)|Representa a propriedade `format.font.color`.|
||[italic](/javascript/api/excel/excel.cellpropertiesfont#italic)|Representa a propriedade `format.font.italic`.|
||[name](/javascript/api/excel/excel.cellpropertiesfont#name)|Representa a propriedade `format.font.name`.|
||[size](/javascript/api/excel/excel.cellpropertiesfont#size)|Representa a propriedade`format.font.size`.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfont#strikethrough)|Representa a propriedade `format.font.strikethrough`.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfont#subscript)|Representa a propriedade `format.font.subscript`.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfont#superscript)|Representa a propriedade `format.font.superscript`.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfont#tintandshade)|Representa a propriedade `format.font.tintAndShade`.|
||[underline](/javascript/api/excel/excel.cellpropertiesfont#underline)|Representa a propriedade `format.font.underline`.|
|[CellPropertiesFontLoadOptions](/javascript/api/excel/excel.cellpropertiesfontloadoptions)|[bold](/javascript/api/excel/excel.cellpropertiesfontloadoptions#bold)|Especifica se a `bold` propriedade deve ser carregada.|
||[color](/javascript/api/excel/excel.cellpropertiesfontloadoptions#color)|Especifica se a `color` propriedade deve ser carregada.|
||[italic](/javascript/api/excel/excel.cellpropertiesfontloadoptions#italic)|Especifica se a `italic` propriedade deve ser carregada.|
||[name](/javascript/api/excel/excel.cellpropertiesfontloadoptions#name)|Especifica se a `name` propriedade deve ser carregada.|
||[size](/javascript/api/excel/excel.cellpropertiesfontloadoptions#size)|Especifica se a `size` propriedade deve ser carregada.|
||[strikethrough](/javascript/api/excel/excel.cellpropertiesfontloadoptions#strikethrough)|Especifica se a `strikethrough` propriedade deve ser carregada.|
||[subscript](/javascript/api/excel/excel.cellpropertiesfontloadoptions#subscript)|Especifica se a `subscript` propriedade deve ser carregada.|
||[superscript](/javascript/api/excel/excel.cellpropertiesfontloadoptions#superscript)|Especifica se a `superscript` propriedade deve ser carregada.|
||[tintAndShade](/javascript/api/excel/excel.cellpropertiesfontloadoptions#tintandshade)|Especifica se a `tintAndShade` propriedade deve ser carregada.|
||[underline](/javascript/api/excel/excel.cellpropertiesfontloadoptions#underline)|Especifica se a `underline` propriedade deve ser carregada.|
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
|[CellPropertiesFormatLoadOptions](/javascript/api/excel/excel.cellpropertiesformatloadoptions)|[autoIndent](/javascript/api/excel/excel.cellpropertiesformatloadoptions#autoindent)|Especifica se a `autoIndent` propriedade deve ser carregada.|
||[Borders](/javascript/api/excel/excel.cellpropertiesformatloadoptions#borders)|Especifica se a `borders` propriedade deve ser carregada.|
||[fill](/javascript/api/excel/excel.cellpropertiesformatloadoptions#fill)|Especifica se a `fill` propriedade deve ser carregada.|
||[font](/javascript/api/excel/excel.cellpropertiesformatloadoptions#font)|Especifica se a `font` propriedade deve ser carregada.|
||[horizontalAlignment](/javascript/api/excel/excel.cellpropertiesformatloadoptions#horizontalalignment)|Especifica se a `horizontalAlignment` propriedade deve ser carregada.|
||[indentLevel](/javascript/api/excel/excel.cellpropertiesformatloadoptions#indentlevel)|Especifica se a `indentLevel` propriedade deve ser carregada.|
||[protection](/javascript/api/excel/excel.cellpropertiesformatloadoptions#protection)|Especifica se a `protection` propriedade deve ser carregada.|
||[readingOrder](/javascript/api/excel/excel.cellpropertiesformatloadoptions#readingorder)|Especifica se a `readingOrder` propriedade deve ser carregada.|
||[shrinkToFit](/javascript/api/excel/excel.cellpropertiesformatloadoptions#shrinktofit)|Especifica se a `shrinkToFit` propriedade deve ser carregada.|
||[textOrientation](/javascript/api/excel/excel.cellpropertiesformatloadoptions#textorientation)|Especifica se a `textOrientation` propriedade deve ser carregada.|
||[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformatloadoptions#usestandardheight)|Especifica se a `useStandardHeight` propriedade deve ser carregada.|
||[useStandardWidth](/javascript/api/excel/excel.cellpropertiesformatloadoptions#usestandardwidth)|Especifica se a `useStandardWidth` propriedade deve ser carregada.|
||[verticalAlignment](/javascript/api/excel/excel.cellpropertiesformatloadoptions#verticalalignment)|Especifica se a `verticalAlignment` propriedade deve ser carregada.|
||[wrapText](/javascript/api/excel/excel.cellpropertiesformatloadoptions#wraptext)|Especifica se a `wrapText` propriedade deve ser carregada.|
|[CellPropertiesLoadOptions](/javascript/api/excel/excel.cellpropertiesloadoptions)|[address](/javascript/api/excel/excel.cellpropertiesloadoptions#address)|Especifica se a `address` propriedade deve ser carregada.|
||[addressLocal](/javascript/api/excel/excel.cellpropertiesloadoptions#addresslocal)|Especifica se a `addressLocal` propriedade deve ser carregada.|
||[format](/javascript/api/excel/excel.cellpropertiesloadoptions#format)|Especifica se a `format` propriedade deve ser carregada.|
||[hidden](/javascript/api/excel/excel.cellpropertiesloadoptions#hidden)|Especifica se a `hidden` propriedade deve ser carregada.|
||[hiperlink](/javascript/api/excel/excel.cellpropertiesloadoptions#hyperlink)|Especifica se a `hyperlink` propriedade deve ser carregada.|
||[style](/javascript/api/excel/excel.cellpropertiesloadoptions#style)|Especifica se a `style` propriedade deve ser carregada.|
|[CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)|[formulaHidden](/javascript/api/excel/excel.cellpropertiesprotection#formulahidden)|Representa a propriedade `format.protection.formulaHidden`.|
||[bloqueado](/javascript/api/excel/excel.cellpropertiesprotection#locked)|Representa a propriedade `format.protection.locked`.|
|[ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail)|[valueAfter](/javascript/api/excel/excel.changedeventdetail#valueafter)|Representa o valor após a alteração. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
||[valueBefore](/javascript/api/excel/excel.changedeventdetail#valuebefore)|Representa o valor antes da alteração. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
||[valueTypeAfter](/javascript/api/excel/excel.changedeventdetail#valuetypeafter)|Representa o tipo de valor após a alteração.|
||[valueTypeBefore](/javascript/api/excel/excel.changedeventdetail#valuetypebefore)|Representa o tipo de valor antes da alteração.|
|[Chart](/javascript/api/excel/excel.chart)|[activate()](/javascript/api/excel/excel.chart#activate--)|Ativa o gráfico na interface do usuário do Excel.|
||[pivotOptions](/javascript/api/excel/excel.chart#pivotoptions)|Encapsula as opções para um gráfico dinâmico. Somente leitura.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[colorScheme](/javascript/api/excel/excel.chartareaformat#colorscheme)|Retorna ou define o esquema de cores do gráfico. Leitura/gravação.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformat#roundedcorners)|Especifica se a área do gráfico tem ou não cantos arredondados. Leitura/gravação.|
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[colorScheme](/javascript/api/excel/excel.chartareaformatdata#colorscheme)|Retorna ou define o esquema de cores do gráfico. Leitura/gravação.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatdata#roundedcorners)|Especifica se a área do gráfico tem ou não cantos arredondados. Leitura/gravação.|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[colorScheme](/javascript/api/excel/excel.chartareaformatloadoptions#colorscheme)|Retorna ou define o esquema de cores do gráfico. Leitura/gravação.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatloadoptions#roundedcorners)|Especifica se a área do gráfico tem ou não cantos arredondados. Leitura/gravação.|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[colorScheme](/javascript/api/excel/excel.chartareaformatupdatedata#colorscheme)|Retorna ou define o esquema de cores do gráfico. Leitura/gravação.|
||[roundedCorners](/javascript/api/excel/excel.chartareaformatupdatedata#roundedcorners)|Especifica se a área do gráfico tem ou não cantos arredondados. Leitura/gravação.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[linkNumberFormat](/javascript/api/excel/excel.chartaxis#linknumberformat)|Representa se o formato numérico está vinculado ou não às células. Se verdadeiro, o formato numérico será alterado nos rótulos quando ele for alterado nas células.|
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisdata#linknumberformat)|Representa se o formato numérico está vinculado ou não às células. Se verdadeiro, o formato numérico será alterado nos rótulos quando ele for alterado nas células.|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisloadoptions#linknumberformat)|Representa se o formato numérico está vinculado ou não às células. Se verdadeiro, o formato numérico será alterado nos rótulos quando ele for alterado nas células.|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartaxisupdatedata#linknumberformat)|Representa se o formato numérico está vinculado ou não às células. Se verdadeiro, o formato numérico será alterado nos rótulos quando ele for alterado nas células.|
|[ChartBinOptions](/javascript/api/excel/excel.chartbinoptions)|[allowOverflow](/javascript/api/excel/excel.chartbinoptions#allowoverflow)|Especifica se o estouro de bin está ativado ou não em um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptions#allowunderflow)|Especifica se o estouro negativo está ou não ativado em um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[Count](/javascript/api/excel/excel.chartbinoptions#count)|Retorna ou define a contagem de bin de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptions#overflowvalue)|Retorna ou define o valor de estouro de bin de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[Set (Propriedades: Excel. ChartBinOptions)](/javascript/api/excel/excel.chartbinoptions#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartBinOptionsUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.chartbinoptions#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[tipo](/javascript/api/excel/excel.chartbinoptions#type)|Retorna ou define o tipo de bin para um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptions#underflowvalue)|Retorna ou define o valor de caixa insuficiente de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[width](/javascript/api/excel/excel.chartbinoptions#width)|Retorna ou define o valor de largura de bin de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
|[ChartBinOptionsData](/javascript/api/excel/excel.chartbinoptionsdata)|[allowOverflow](/javascript/api/excel/excel.chartbinoptionsdata#allowoverflow)|Especifica se o estouro de bin está ativado ou não em um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsdata#allowunderflow)|Especifica se o estouro negativo está ou não ativado em um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[Count](/javascript/api/excel/excel.chartbinoptionsdata#count)|Retorna ou define a contagem de bin de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsdata#overflowvalue)|Retorna ou define o valor de estouro de bin de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[tipo](/javascript/api/excel/excel.chartbinoptionsdata#type)|Retorna ou define o tipo de bin para um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsdata#underflowvalue)|Retorna ou define o valor de caixa insuficiente de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[width](/javascript/api/excel/excel.chartbinoptionsdata#width)|Retorna ou define o valor de largura de bin de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
|[ChartBinOptionsLoadOptions](/javascript/api/excel/excel.chartbinoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartbinoptionsloadoptions#$all)||
||[allowOverflow](/javascript/api/excel/excel.chartbinoptionsloadoptions#allowoverflow)|Especifica se o estouro de bin está ativado ou não em um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsloadoptions#allowunderflow)|Especifica se o estouro negativo está ou não ativado em um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[Count](/javascript/api/excel/excel.chartbinoptionsloadoptions#count)|Retorna ou define a contagem de bin de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsloadoptions#overflowvalue)|Retorna ou define o valor de estouro de bin de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[tipo](/javascript/api/excel/excel.chartbinoptionsloadoptions#type)|Retorna ou define o tipo de bin para um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsloadoptions#underflowvalue)|Retorna ou define o valor de caixa insuficiente de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[width](/javascript/api/excel/excel.chartbinoptionsloadoptions#width)|Retorna ou define o valor de largura de bin de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
|[ChartBinOptionsUpdateData](/javascript/api/excel/excel.chartbinoptionsupdatedata)|[allowOverflow](/javascript/api/excel/excel.chartbinoptionsupdatedata#allowoverflow)|Especifica se o estouro de bin está ativado ou não em um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[allowUnderflow](/javascript/api/excel/excel.chartbinoptionsupdatedata#allowunderflow)|Especifica se o estouro negativo está ou não ativado em um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[Count](/javascript/api/excel/excel.chartbinoptionsupdatedata#count)|Retorna ou define a contagem de bin de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[overflowValue](/javascript/api/excel/excel.chartbinoptionsupdatedata#overflowvalue)|Retorna ou define o valor de estouro de bin de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[tipo](/javascript/api/excel/excel.chartbinoptionsupdatedata#type)|Retorna ou define o tipo de bin para um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[underflowValue](/javascript/api/excel/excel.chartbinoptionsupdatedata#underflowvalue)|Retorna ou define o valor de caixa insuficiente de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
||[width](/javascript/api/excel/excel.chartbinoptionsupdatedata#width)|Retorna ou define o valor de largura de bin de um gráfico de histograma ou gráfico de pareto. Leitura/gravação.|
|[ChartBoxwhiskerOptions](/javascript/api/excel/excel.chartboxwhiskeroptions)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptions#quartilecalculation)|Retorna ou define o tipo de cálculo quartil de um gráfico de caixa estreita. Leitura/gravação.|
||[Set (Propriedades: Excel. ChartBoxwhiskerOptions)](/javascript/api/excel/excel.chartboxwhiskeroptions#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartBoxwhiskerOptionsUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.chartboxwhiskeroptions#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showinnerpoints)|Especifica se os pontos internos são mostrados ou não em um gráfico de caixa estreita. Leitura/gravação.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanline)|Especifica se a linha média é mostrada ou não em um gráfico de caixa estreita. Leitura/gravação.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptions#showmeanmarker)|Especifica se o marcador de média é ou não mostrado em um gráfico de caixa estreita. Leitura/gravação.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptions#showoutlierpoints)|Especifica se os pontos discrepantes são mostrados ou não em um gráfico de caixa estreita. Leitura/gravação.|
|[ChartBoxwhiskerOptionsData](/javascript/api/excel/excel.chartboxwhiskeroptionsdata)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#quartilecalculation)|Retorna ou define o tipo de cálculo quartil de um gráfico de caixa estreita. Leitura/gravação.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showinnerpoints)|Especifica se os pontos internos são mostrados ou não em um gráfico de caixa estreita. Leitura/gravação.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showmeanline)|Especifica se a linha média é mostrada ou não em um gráfico de caixa estreita. Leitura/gravação.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showmeanmarker)|Especifica se o marcador de média é ou não mostrado em um gráfico de caixa estreita. Leitura/gravação.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsdata#showoutlierpoints)|Especifica se os pontos discrepantes são mostrados ou não em um gráfico de caixa estreita. Leitura/gravação.|
|[ChartBoxwhiskerOptionsLoadOptions](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions)|[$all](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#$all)||
||[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#quartilecalculation)|Retorna ou define o tipo de cálculo quartil de um gráfico de caixa estreita. Leitura/gravação.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showinnerpoints)|Especifica se os pontos internos são mostrados ou não em um gráfico de caixa estreita. Leitura/gravação.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showmeanline)|Especifica se a linha média é mostrada ou não em um gráfico de caixa estreita. Leitura/gravação.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showmeanmarker)|Especifica se o marcador de média é ou não mostrado em um gráfico de caixa estreita. Leitura/gravação.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsloadoptions#showoutlierpoints)|Especifica se os pontos discrepantes são mostrados ou não em um gráfico de caixa estreita. Leitura/gravação.|
|[ChartBoxwhiskerOptionsUpdateData](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata)|[quartileCalculation](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#quartilecalculation)|Retorna ou define o tipo de cálculo quartil de um gráfico de caixa estreita. Leitura/gravação.|
||[showInnerPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showinnerpoints)|Especifica se os pontos internos são mostrados ou não em um gráfico de caixa estreita. Leitura/gravação.|
||[showMeanLine](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showmeanline)|Especifica se a linha média é mostrada ou não em um gráfico de caixa estreita. Leitura/gravação.|
||[showMeanMarker](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showmeanmarker)|Especifica se o marcador de média é ou não mostrado em um gráfico de caixa estreita. Leitura/gravação.|
||[showOutlierPoints](/javascript/api/excel/excel.chartboxwhiskeroptionsupdatedata#showoutlierpoints)|Especifica se os pontos discrepantes são mostrados ou não em um gráfico de caixa estreita. Leitura/gravação.|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[pivotOptions](/javascript/api/excel/excel.chartcollectionloadoptions#pivotoptions)|Para cada ITEM na coleção: encapsula as opções de um gráfico dinâmico.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[pivotOptions](/javascript/api/excel/excel.chartdata#pivotoptions)|Encapsula as opções para um gráfico dinâmico. Somente leitura.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabel#linknumberformat)|Valor booliano que representa se o formato de número está vinculado às células (de modo que o formato de número mude nos rótulos quando for alterado nas células).|
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabeldata#linknumberformat)|Valor booliano que representa se o formato de número está vinculado às células (de modo que o formato de número mude nos rótulos quando for alterado nas células).|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelloadoptions#linknumberformat)|Valor booliano que representa se o formato de número está vinculado às células (de modo que o formato de número mude nos rótulos quando for alterado nas células).|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelupdatedata#linknumberformat)|Valor booliano que representa se o formato de número está vinculado às células (de modo que o formato de número mude nos rótulos quando for alterado nas células).|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabels#linknumberformat)|Representa se o formato numérico está vinculado ou não às células. Se verdadeiro, o formato numérico será alterado nos rótulos quando ele for alterado nas células|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsdata#linknumberformat)|Representa se o formato numérico está vinculado ou não às células. Se verdadeiro, o formato numérico será alterado nos rótulos quando ele for alterado nas células|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsloadoptions#linknumberformat)|Representa se o formato numérico está vinculado ou não às células. Se verdadeiro, o formato numérico será alterado nos rótulos quando ele for alterado nas células|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.chartdatalabelsupdatedata#linknumberformat)|Representa se o formato numérico está vinculado ou não às células. Se verdadeiro, o formato numérico será alterado nos rótulos quando ele for alterado nas células|
|[ChartErrorBars](/javascript/api/excel/excel.charterrorbars)|[endStyleCap](/javascript/api/excel/excel.charterrorbars#endstylecap)|Especifica se as barras de erro possuem ou não um limite de estilo final.|
||[include](/javascript/api/excel/excel.charterrorbars#include)|Especifica quais partes das barras de erro devem ser incluídas.|
||[format](/javascript/api/excel/excel.charterrorbars#format)|Especifica o tipo de formatação das barras de erro.|
||[Set (Propriedades: Excel. ChartErrorBars)](/javascript/api/excel/excel.charterrorbars#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartErrorBarsUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.charterrorbars#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[tipo](/javascript/api/excel/excel.charterrorbars#type)|O tipo de intervalo marcado pelas barras de erro.|
||[visible](/javascript/api/excel/excel.charterrorbars#visible)|Especifica se as barras de erro são exibidas ou não.|
|[ChartErrorBarsData](/javascript/api/excel/excel.charterrorbarsdata)|[endStyleCap](/javascript/api/excel/excel.charterrorbarsdata#endstylecap)|Especifica se as barras de erro possuem ou não um limite de estilo final.|
||[format](/javascript/api/excel/excel.charterrorbarsdata#format)|Especifica o tipo de formatação das barras de erro.|
||[include](/javascript/api/excel/excel.charterrorbarsdata#include)|Especifica quais partes das barras de erro devem ser incluídas.|
||[tipo](/javascript/api/excel/excel.charterrorbarsdata#type)|O tipo de intervalo marcado pelas barras de erro.|
||[visible](/javascript/api/excel/excel.charterrorbarsdata#visible)|Especifica se as barras de erro são exibidas ou não.|
|[ChartErrorBarsFormat](/javascript/api/excel/excel.charterrorbarsformat)|[line](/javascript/api/excel/excel.charterrorbarsformat#line)|Representa a formatação de linha do gráfico.|
||[Set (Propriedades: Excel. ChartErrorBarsFormat)](/javascript/api/excel/excel.charterrorbarsformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartErrorBarsFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.charterrorbarsformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[ChartErrorBarsFormatData](/javascript/api/excel/excel.charterrorbarsformatdata)|[line](/javascript/api/excel/excel.charterrorbarsformatdata#line)|Representa a formatação de linha do gráfico.|
|[ChartErrorBarsFormatLoadOptions](/javascript/api/excel/excel.charterrorbarsformatloadoptions)|[$all](/javascript/api/excel/excel.charterrorbarsformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.charterrorbarsformatloadoptions#line)|Representa a formatação de linha do gráfico.|
|[ChartErrorBarsFormatUpdateData](/javascript/api/excel/excel.charterrorbarsformatupdatedata)|[line](/javascript/api/excel/excel.charterrorbarsformatupdatedata#line)|Representa a formatação de linha do gráfico.|
|[ChartErrorBarsLoadOptions](/javascript/api/excel/excel.charterrorbarsloadoptions)|[$all](/javascript/api/excel/excel.charterrorbarsloadoptions#$all)||
||[endStyleCap](/javascript/api/excel/excel.charterrorbarsloadoptions#endstylecap)|Especifica se as barras de erro possuem ou não um limite de estilo final.|
||[format](/javascript/api/excel/excel.charterrorbarsloadoptions#format)|Especifica o tipo de formatação das barras de erro.|
||[include](/javascript/api/excel/excel.charterrorbarsloadoptions#include)|Especifica quais partes das barras de erro devem ser incluídas.|
||[tipo](/javascript/api/excel/excel.charterrorbarsloadoptions#type)|O tipo de intervalo marcado pelas barras de erro.|
||[visible](/javascript/api/excel/excel.charterrorbarsloadoptions#visible)|Especifica se as barras de erro são exibidas ou não.|
|[ChartErrorBarsUpdateData](/javascript/api/excel/excel.charterrorbarsupdatedata)|[endStyleCap](/javascript/api/excel/excel.charterrorbarsupdatedata#endstylecap)|Especifica se as barras de erro possuem ou não um limite de estilo final.|
||[format](/javascript/api/excel/excel.charterrorbarsupdatedata#format)|Especifica o tipo de formatação das barras de erro.|
||[include](/javascript/api/excel/excel.charterrorbarsupdatedata#include)|Especifica quais partes das barras de erro devem ser incluídas.|
||[tipo](/javascript/api/excel/excel.charterrorbarsupdatedata#type)|O tipo de intervalo marcado pelas barras de erro.|
||[visible](/javascript/api/excel/excel.charterrorbarsupdatedata#visible)|Especifica se as barras de erro são exibidas ou não.|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[pivotOptions](/javascript/api/excel/excel.chartloadoptions#pivotoptions)|Encapsula as opções para um gráfico dinâmico.|
|[ChartMapOptions](/javascript/api/excel/excel.chartmapoptions)|[labelStrategy](/javascript/api/excel/excel.chartmapoptions#labelstrategy)|Retorna ou define a estratégia de rótulos de mapa da série de um gráfico de mapa de região. Leitura/gravação.|
||[level](/javascript/api/excel/excel.chartmapoptions#level)|Retorna ou define o nível de mapeamento de série de um gráfico de mapa de região. Leitura/gravação.|
||[projectionType](/javascript/api/excel/excel.chartmapoptions#projectiontype)|Retorna ou define o tipo de projeção em série de um gráfico de mapa de região. Leitura/gravação.|
||[Set (Propriedades: Excel. ChartMapOptions)](/javascript/api/excel/excel.chartmapoptions#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartMapOptionsUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.chartmapoptions#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[ChartMapOptionsData](/javascript/api/excel/excel.chartmapoptionsdata)|[labelStrategy](/javascript/api/excel/excel.chartmapoptionsdata#labelstrategy)|Retorna ou define a estratégia de rótulos de mapa da série de um gráfico de mapa de região. Leitura/gravação.|
||[level](/javascript/api/excel/excel.chartmapoptionsdata#level)|Retorna ou define o nível de mapeamento de série de um gráfico de mapa de região. Leitura/gravação.|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsdata#projectiontype)|Retorna ou define o tipo de projeção em série de um gráfico de mapa de região. Leitura/gravação.|
|[ChartMapOptionsLoadOptions](/javascript/api/excel/excel.chartmapoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartmapoptionsloadoptions#$all)||
||[labelStrategy](/javascript/api/excel/excel.chartmapoptionsloadoptions#labelstrategy)|Retorna ou define a estratégia de rótulos de mapa da série de um gráfico de mapa de região. Leitura/gravação.|
||[level](/javascript/api/excel/excel.chartmapoptionsloadoptions#level)|Retorna ou define o nível de mapeamento de série de um gráfico de mapa de região. Leitura/gravação.|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsloadoptions#projectiontype)|Retorna ou define o tipo de projeção em série de um gráfico de mapa de região. Leitura/gravação.|
|[ChartMapOptionsUpdateData](/javascript/api/excel/excel.chartmapoptionsupdatedata)|[labelStrategy](/javascript/api/excel/excel.chartmapoptionsupdatedata#labelstrategy)|Retorna ou define a estratégia de rótulos de mapa da série de um gráfico de mapa de região. Leitura/gravação.|
||[level](/javascript/api/excel/excel.chartmapoptionsupdatedata#level)|Retorna ou define o nível de mapeamento de série de um gráfico de mapa de região. Leitura/gravação.|
||[projectionType](/javascript/api/excel/excel.chartmapoptionsupdatedata#projectiontype)|Retorna ou define o tipo de projeção em série de um gráfico de mapa de região. Leitura/gravação.|
|[ChartPivotOptions](/javascript/api/excel/excel.chartpivotoptions)|[Set (Propriedades: Excel. ChartPivotOptions)](/javascript/api/excel/excel.chartpivotoptions#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartPivotOptionsUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.chartpivotoptions#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showaxisfieldbuttons)|Especifica se deve ou não exibir os botões de campo de eixo em um gráfico dinâmico. A propriedade ShowAxisFieldButtons corresponde ao comando "Exibir Botões de Campo de Eixo" na lista suspensa "Botões de Campo" da guia "Analisar", que está disponível quando um gráfico dinâmico é selecionado.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showlegendfieldbuttons)|Especifica se deve ou não exibir os botões de campo de legenda em um gráfico dinâmico|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showreportfilterfieldbuttons)|Especifica se deve ou não exibir os botões de campo do filtro de relatório em um gráfico dinâmico.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptions#showvaluefieldbuttons)|Especifica se deve ou não exibir os botões de exibir campo de valor em um gráfico dinâmico|
|[ChartPivotOptionsData](/javascript/api/excel/excel.chartpivotoptionsdata)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showaxisfieldbuttons)|Especifica se deve ou não exibir os botões de campo de eixo em um gráfico dinâmico. A propriedade ShowAxisFieldButtons corresponde ao comando "Exibir Botões de Campo de Eixo" na lista suspensa "Botões de Campo" da guia "Analisar", que está disponível quando um gráfico dinâmico é selecionado.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showlegendfieldbuttons)|Especifica se deve ou não exibir os botões de campo de legenda em um gráfico dinâmico|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showreportfilterfieldbuttons)|Especifica se deve ou não exibir os botões de campo do filtro de relatório em um gráfico dinâmico.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsdata#showvaluefieldbuttons)|Especifica se deve ou não exibir os botões de exibir campo de valor em um gráfico dinâmico|
|[ChartPivotOptionsLoadOptions](/javascript/api/excel/excel.chartpivotoptionsloadoptions)|[$all](/javascript/api/excel/excel.chartpivotoptionsloadoptions#$all)||
||[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showaxisfieldbuttons)|Especifica se deve ou não exibir os botões de campo de eixo em um gráfico dinâmico. A propriedade ShowAxisFieldButtons corresponde ao comando "Exibir Botões de Campo de Eixo" na lista suspensa "Botões de Campo" da guia "Analisar", que está disponível quando um gráfico dinâmico é selecionado.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showlegendfieldbuttons)|Especifica se deve ou não exibir os botões de campo de legenda em um gráfico dinâmico|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showreportfilterfieldbuttons)|Especifica se deve ou não exibir os botões de campo do filtro de relatório em um gráfico dinâmico.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsloadoptions#showvaluefieldbuttons)|Especifica se deve ou não exibir os botões de exibir campo de valor em um gráfico dinâmico|
|[ChartPivotOptionsUpdateData](/javascript/api/excel/excel.chartpivotoptionsupdatedata)|[showAxisFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showaxisfieldbuttons)|Especifica se deve ou não exibir os botões de campo de eixo em um gráfico dinâmico. A propriedade ShowAxisFieldButtons corresponde ao comando "Exibir Botões de Campo de Eixo" na lista suspensa "Botões de Campo" da guia "Analisar", que está disponível quando um gráfico dinâmico é selecionado.|
||[showLegendFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showlegendfieldbuttons)|Especifica se deve ou não exibir os botões de campo de legenda em um gráfico dinâmico|
||[showReportFilterFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showreportfilterfieldbuttons)|Especifica se deve ou não exibir os botões de campo do filtro de relatório em um gráfico dinâmico.|
||[showValueFieldButtons](/javascript/api/excel/excel.chartpivotoptionsupdatedata#showvaluefieldbuttons)|Especifica se deve ou não exibir os botões de exibir campo de valor em um gráfico dinâmico|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[bubbleScale](/javascript/api/excel/excel.chartseries#bubblescale)|Este pode ser um valor inteiro de 0 (zero) a 300, representando a porcentagem do tamanho padrão. Esta propriedade só se aplica a gráficos de bolhas. Leitura/gravação.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseries#gradientmaximumcolor)|Retorna ou define a cor para o valor máximo de uma série de gráficos de mapa de região. Leitura/gravação.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseries#gradientmaximumtype)|Retorna ou define o tipo para o valor máximo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseries#gradientmaximumvalue)|Retorna ou define o valor máximo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseries#gradientmidpointcolor)|Retorna ou define a cor do valor do ponto médio de uma série de gráficos de mapa de região. Leitura/gravação.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseries#gradientmidpointtype)|Retorna ou define o tipo para o valor médio de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseries#gradientmidpointvalue)|Retorna ou define o valor médio de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseries#gradientminimumcolor)|Retorna ou define a cor para o valor mínimo de uma série de gráficos de mapa de região. Leitura/gravação.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseries#gradientminimumtype)|Retorna ou define o tipo para o valor mínimo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseries#gradientminimumvalue)|Retorna ou define o valor mínimo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientStyle](/javascript/api/excel/excel.chartseries#gradientstyle)|Retorna ou define o estilo de gradiente da série de um gráfico de mapa da região. Leitura/gravação.|
||[invertColor](/javascript/api/excel/excel.chartseries#invertcolor)|Retorna ou define a cor de preenchimento para pontos de dados negativo de uma série. Leitura/gravação.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseries#parentlabelstrategy)|Retorna ou define a área de estratégia de rótulo pai da série para um gráfico de mapa de árvore. Leitura/gravação.|
||[binOptions](/javascript/api/excel/excel.chartseries#binoptions)|Encapsula as opções de bin para gráficos de histograma e gráficos de pareto. Somente leitura.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseries#boxwhiskeroptions)|Encapsula as opções para os gráficos de caixa estreita. Somente leitura.|
||[mapOptions](/javascript/api/excel/excel.chartseries#mapoptions)|Encapsula as opções para um gráfico de mapa de região. Somente leitura.|
||[xErrorBars](/javascript/api/excel/excel.chartseries#xerrorbars)|Representa o objeto da barra de erros de uma série de gráficos.|
||[yErrorBars](/javascript/api/excel/excel.chartseries#yerrorbars)|Representa o objeto da barra de erros de uma série de gráficos.|
||[showConnectorLines](/javascript/api/excel/excel.chartseries#showconnectorlines)|Especifica se as linhas de conexão são mostradas ou não nos gráficos em cascata. Leitura/gravação.|
||[showLeaderLines](/javascript/api/excel/excel.chartseries#showleaderlines)|Especifica se as linhas de preenchimento são exibidas ou não para cada rótulo de dados na série. Leitura/gravação.|
||[splitValue](/javascript/api/excel/excel.chartseries#splitvalue)|Retorna ou define o valor de limite que separa duas seções de um gráfico de pizza de pizza ou gráfico de barra de pizza. Leitura/gravação.|
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[binOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#binoptions)|Para cada ITEM na coleção: encapsula as opções de compartimento para gráficos de histograma e gráficos de Pareto.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#boxwhiskeroptions)|Para cada ITEM na coleção: encapsula as opções para os gráficos caixa e estreita.|
||[bubbleScale](/javascript/api/excel/excel.chartseriescollectionloadoptions#bubblescale)|Para cada ITEM na coleção: pode ser um valor inteiro de 0 (zero) a 300, representando a porcentagem do tamanho padrão. Esta propriedade só se aplica a gráficos de bolhas. Leitura/gravação.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumcolor)|Para cada ITEM na coleção: Retorna ou define a cor do valor máximo de uma série de gráficos do mapa de região. Leitura/gravação.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumtype)|Para cada ITEM na coleção: Retorna ou define o tipo para o valor máximo de uma série de gráficos do mapa de região. Leitura/gravação.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmaximumvalue)|Para cada ITEM na coleção: Retorna ou define o valor máximo de uma série de gráfico do mapa de região. Leitura/gravação.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointcolor)|Para cada ITEM na coleção: Retorna ou define a cor do valor intermediário de uma série de gráfico do mapa de região. Leitura/gravação.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointtype)|Para cada ITEM na coleção: Retorna ou define o tipo de valor de ponto médio de uma série de gráfico do mapa de região. Leitura/gravação.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientmidpointvalue)|Para cada ITEM na coleção: Retorna ou define o valor de ponto médio de uma série de gráficos do mapa de região. Leitura/gravação.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumcolor)|Para cada ITEM na coleção: Retorna ou define a cor do valor mínimo de uma série de gráficos do mapa de região. Leitura/gravação.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumtype)|Para cada ITEM na coleção: Retorna ou define o tipo para o valor mínimo de uma série de gráficos do mapa de região. Leitura/gravação.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientminimumvalue)|Para cada ITEM na coleção: Retorna ou define o valor mínimo de uma série de gráfico do mapa de região. Leitura/gravação.|
||[gradientStyle](/javascript/api/excel/excel.chartseriescollectionloadoptions#gradientstyle)|Para cada ITEM na coleção: Retorna ou define o estilo de gradiente de série de um gráfico de mapa de região. Leitura/gravação.|
||[invertColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#invertcolor)|Para cada ITEM na coleção: Retorna ou define a cor de preenchimento de pontos de dados negativos em uma série. Leitura/gravação.|
||[mapOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions#mapoptions)|Para cada ITEM na coleção: encapsula as opções de um gráfico de mapa de região.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriescollectionloadoptions#parentlabelstrategy)|Para cada ITEM na coleção: Retorna ou define a área de estratégia de rótulo pai da série para um gráfico de mapa de região. Leitura/gravação.|
||[showConnectorLines](/javascript/api/excel/excel.chartseriescollectionloadoptions#showconnectorlines)|Para cada ITEM na coleção: especifica se as linhas de conexão são ou não mostradas em gráficos de cascata. Leitura/gravação.|
||[showLeaderLines](/javascript/api/excel/excel.chartseriescollectionloadoptions#showleaderlines)|Para cada ITEM na coleção: especifica se as linhas de preenchimento são exibidas para cada rótulo de dados na série. Leitura/gravação.|
||[splitValue](/javascript/api/excel/excel.chartseriescollectionloadoptions#splitvalue)|Para cada ITEM na coleção: Retorna ou define o valor de limite que separa duas seções de um gráfico de pizza de pizza ou um gráfico de barra de pizza. Leitura/gravação.|
||[xErrorBars](/javascript/api/excel/excel.chartseriescollectionloadoptions#xerrorbars)|Para cada ITEM na coleção: representa o objeto de barra de erro de uma série de gráfico.|
||[yErrorBars](/javascript/api/excel/excel.chartseriescollectionloadoptions#yerrorbars)|Para cada ITEM na coleção: representa o objeto de barra de erro de uma série de gráfico.|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[binOptions](/javascript/api/excel/excel.chartseriesdata#binoptions)|Encapsula as opções de bin para gráficos de histograma e gráficos de pareto. Somente leitura.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesdata#boxwhiskeroptions)|Encapsula as opções para os gráficos de caixa estreita. Somente leitura.|
||[bubbleScale](/javascript/api/excel/excel.chartseriesdata#bubblescale)|Este pode ser um valor inteiro de 0 (zero) a 300, representando a porcentagem do tamanho padrão. Esta propriedade só se aplica a gráficos de bolhas. Leitura/gravação.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesdata#gradientmaximumcolor)|Retorna ou define a cor para o valor máximo de uma série de gráficos de mapa de região. Leitura/gravação.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesdata#gradientmaximumtype)|Retorna ou define o tipo para o valor máximo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesdata#gradientmaximumvalue)|Retorna ou define o valor máximo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesdata#gradientmidpointcolor)|Retorna ou define a cor do valor do ponto médio de uma série de gráficos de mapa de região. Leitura/gravação.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesdata#gradientmidpointtype)|Retorna ou define o tipo para o valor médio de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesdata#gradientmidpointvalue)|Retorna ou define o valor médio de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesdata#gradientminimumcolor)|Retorna ou define a cor para o valor mínimo de uma série de gráficos de mapa de região. Leitura/gravação.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesdata#gradientminimumtype)|Retorna ou define o tipo para o valor mínimo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesdata#gradientminimumvalue)|Retorna ou define o valor mínimo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientStyle](/javascript/api/excel/excel.chartseriesdata#gradientstyle)|Retorna ou define o estilo de gradiente da série de um gráfico de mapa da região. Leitura/gravação.|
||[invertColor](/javascript/api/excel/excel.chartseriesdata#invertcolor)|Retorna ou define a cor de preenchimento para pontos de dados negativo de uma série. Leitura/gravação.|
||[mapOptions](/javascript/api/excel/excel.chartseriesdata#mapoptions)|Encapsula as opções para um gráfico de mapa de região. Somente leitura.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesdata#parentlabelstrategy)|Retorna ou define a área de estratégia de rótulo pai da série para um gráfico de mapa de árvore. Leitura/gravação.|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesdata#showconnectorlines)|Especifica se as linhas de conexão são mostradas ou não nos gráficos em cascata. Leitura/gravação.|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesdata#showleaderlines)|Especifica se as linhas de preenchimento são exibidas ou não para cada rótulo de dados na série. Leitura/gravação.|
||[splitValue](/javascript/api/excel/excel.chartseriesdata#splitvalue)|Retorna ou define o valor de limite que separa duas seções de um gráfico de pizza de pizza ou gráfico de barra de pizza. Leitura/gravação.|
||[xErrorBars](/javascript/api/excel/excel.chartseriesdata#xerrorbars)|Representa o objeto da barra de erros de uma série de gráficos.|
||[yErrorBars](/javascript/api/excel/excel.chartseriesdata#yerrorbars)|Representa o objeto da barra de erros de uma série de gráficos.|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[binOptions](/javascript/api/excel/excel.chartseriesloadoptions#binoptions)|Encapsula as opções de bin para gráficos de histograma e gráficos de pareto.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesloadoptions#boxwhiskeroptions)|Encapsula as opções para os gráficos de caixa estreita.|
||[bubbleScale](/javascript/api/excel/excel.chartseriesloadoptions#bubblescale)|Este pode ser um valor inteiro de 0 (zero) a 300, representando a porcentagem do tamanho padrão. Esta propriedade só se aplica a gráficos de bolhas. Leitura/gravação.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumcolor)|Retorna ou define a cor para o valor máximo de uma série de gráficos de mapa de região. Leitura/gravação.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumtype)|Retorna ou define o tipo para o valor máximo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientmaximumvalue)|Retorna ou define o valor máximo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointcolor)|Retorna ou define a cor do valor do ponto médio de uma série de gráficos de mapa de região. Leitura/gravação.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointtype)|Retorna ou define o tipo para o valor médio de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientmidpointvalue)|Retorna ou define o valor médio de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumcolor)|Retorna ou define a cor para o valor mínimo de uma série de gráficos de mapa de região. Leitura/gravação.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumtype)|Retorna ou define o tipo para o valor mínimo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesloadoptions#gradientminimumvalue)|Retorna ou define o valor mínimo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientStyle](/javascript/api/excel/excel.chartseriesloadoptions#gradientstyle)|Retorna ou define o estilo de gradiente da série de um gráfico de mapa da região. Leitura/gravação.|
||[invertColor](/javascript/api/excel/excel.chartseriesloadoptions#invertcolor)|Retorna ou define a cor de preenchimento para pontos de dados negativo de uma série. Leitura/gravação.|
||[mapOptions](/javascript/api/excel/excel.chartseriesloadoptions#mapoptions)|Encapsula as opções para um gráfico de mapa de região.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesloadoptions#parentlabelstrategy)|Retorna ou define a área de estratégia de rótulo pai da série para um gráfico de mapa de árvore. Leitura/gravação.|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesloadoptions#showconnectorlines)|Especifica se as linhas de conexão são mostradas ou não nos gráficos em cascata. Leitura/gravação.|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesloadoptions#showleaderlines)|Especifica se as linhas de preenchimento são exibidas ou não para cada rótulo de dados na série. Leitura/gravação.|
||[splitValue](/javascript/api/excel/excel.chartseriesloadoptions#splitvalue)|Retorna ou define o valor de limite que separa duas seções de um gráfico de pizza de pizza ou gráfico de barra de pizza. Leitura/gravação.|
||[xErrorBars](/javascript/api/excel/excel.chartseriesloadoptions#xerrorbars)|Representa o objeto da barra de erros de uma série de gráficos.|
||[yErrorBars](/javascript/api/excel/excel.chartseriesloadoptions#yerrorbars)|Representa o objeto da barra de erros de uma série de gráficos.|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[binOptions](/javascript/api/excel/excel.chartseriesupdatedata#binoptions)|Encapsula as opções de bin para gráficos de histograma e gráficos de pareto.|
||[boxwhiskerOptions](/javascript/api/excel/excel.chartseriesupdatedata#boxwhiskeroptions)|Encapsula as opções para os gráficos de caixa estreita.|
||[bubbleScale](/javascript/api/excel/excel.chartseriesupdatedata#bubblescale)|Este pode ser um valor inteiro de 0 (zero) a 300, representando a porcentagem do tamanho padrão. Esta propriedade só se aplica a gráficos de bolhas. Leitura/gravação.|
||[gradientMaximumColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumcolor)|Retorna ou define a cor para o valor máximo de uma série de gráficos de mapa de região. Leitura/gravação.|
||[gradientMaximumType](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumtype)|Retorna ou define o tipo para o valor máximo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMaximumValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientmaximumvalue)|Retorna ou define o valor máximo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMidpointColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointcolor)|Retorna ou define a cor do valor do ponto médio de uma série de gráficos de mapa de região. Leitura/gravação.|
||[gradientMidpointType](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointtype)|Retorna ou define o tipo para o valor médio de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMidpointValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientmidpointvalue)|Retorna ou define o valor médio de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMinimumColor](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumcolor)|Retorna ou define a cor para o valor mínimo de uma série de gráficos de mapa de região. Leitura/gravação.|
||[gradientMinimumType](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumtype)|Retorna ou define o tipo para o valor mínimo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientMinimumValue](/javascript/api/excel/excel.chartseriesupdatedata#gradientminimumvalue)|Retorna ou define o valor mínimo de uma série de gráficos de mapa da região. Leitura/gravação.|
||[gradientStyle](/javascript/api/excel/excel.chartseriesupdatedata#gradientstyle)|Retorna ou define o estilo de gradiente da série de um gráfico de mapa da região. Leitura/gravação.|
||[invertColor](/javascript/api/excel/excel.chartseriesupdatedata#invertcolor)|Retorna ou define a cor de preenchimento para pontos de dados negativo de uma série. Leitura/gravação.|
||[mapOptions](/javascript/api/excel/excel.chartseriesupdatedata#mapoptions)|Encapsula as opções para um gráfico de mapa de região.|
||[parentLabelStrategy](/javascript/api/excel/excel.chartseriesupdatedata#parentlabelstrategy)|Retorna ou define a área de estratégia de rótulo pai da série para um gráfico de mapa de árvore. Leitura/gravação.|
||[showConnectorLines](/javascript/api/excel/excel.chartseriesupdatedata#showconnectorlines)|Especifica se as linhas de conexão são mostradas ou não nos gráficos em cascata. Leitura/gravação.|
||[showLeaderLines](/javascript/api/excel/excel.chartseriesupdatedata#showleaderlines)|Especifica se as linhas de preenchimento são exibidas ou não para cada rótulo de dados na série. Leitura/gravação.|
||[splitValue](/javascript/api/excel/excel.chartseriesupdatedata#splitvalue)|Retorna ou define o valor de limite que separa duas seções de um gráfico de pizza de pizza ou gráfico de barra de pizza. Leitura/gravação.|
||[xErrorBars](/javascript/api/excel/excel.chartseriesupdatedata#xerrorbars)|Representa o objeto da barra de erros de uma série de gráficos.|
||[yErrorBars](/javascript/api/excel/excel.chartseriesupdatedata#yerrorbars)|Representa o objeto da barra de erros de uma série de gráficos.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabel#linknumberformat)|Valor booliano que representa se o formato de número está vinculado às células (de modo que o formato de número mude nos rótulos quando for alterado nas células).|
|[ChartTrendlineLabelData](/javascript/api/excel/excel.charttrendlinelabeldata)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabeldata#linknumberformat)|Valor booliano que representa se o formato de número está vinculado às células (de modo que o formato de número mude nos rótulos quando for alterado nas células).|
|[ChartTrendlineLabelLoadOptions](/javascript/api/excel/excel.charttrendlinelabelloadoptions)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabelloadoptions#linknumberformat)|Valor booliano que representa se o formato de número está vinculado às células (de modo que o formato de número mude nos rótulos quando for alterado nas células).|
|[ChartTrendlineLabelUpdateData](/javascript/api/excel/excel.charttrendlinelabelupdatedata)|[linkNumberFormat](/javascript/api/excel/excel.charttrendlinelabelupdatedata#linknumberformat)|Valor booliano que representa se o formato de número está vinculado às células (de modo que o formato de número mude nos rótulos quando for alterado nas células).|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[pivotOptions](/javascript/api/excel/excel.chartupdatedata#pivotoptions)|Encapsula as opções para um gráfico dinâmico.|
|[ColumnProperties](/javascript/api/excel/excel.columnproperties)|[address](/javascript/api/excel/excel.columnproperties#address)|Representa a propriedade `address`.|
||[addressLocal](/javascript/api/excel/excel.columnproperties#addresslocal)|Representa a propriedade `addressLocal`.|
||[columnIndex](/javascript/api/excel/excel.columnproperties#columnindex)|Representa a propriedade `columnIndex`.|
|[ColumnPropertiesLoadOptions](/javascript/api/excel/excel.columnpropertiesloadoptions)|[columnHidden](/javascript/api/excel/excel.columnpropertiesloadoptions#columnhidden)|Especifica se a `columnHidden` propriedade deve ser carregada.|
||[columnIndex](/javascript/api/excel/excel.columnpropertiesloadoptions#columnindex)|Especifica se a `columnIndex` propriedade deve ser carregada.|
||[columnWidth](/javascript/api/excel/excel.columnpropertiesloadoptions#columnwidth)||
||[formato: Excel. CellPropertiesFormatLoadOptions & {
            columnWidth?] (formato/JavaScript/API/Excel/Excel.columnpropertiesloadoptions #)|Especifica se a `format` propriedade deve ser carregada.|
|[ConditionalFormat](/javascript/api/excel/excel.conditionalformat)|[getRanges()](/javascript/api/excel/excel.conditionalformat#getranges--)|Retorna o RangeAreas, compreendendo um ou mais intervalos retangulares, ao qual o formato condicional é aplicado. Somente leitura.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[getInvalidCells()](/javascript/api/excel/excel.datavalidation#getinvalidcells--)|Retorna um RangeAreas, que consiste em um ou mais intervalos retangulares, com valores inválidos de célula. Se todos os valores de célula forem válidos, essa função gerará um erro ItemNotFound.|
||[getInvalidCellsOrNullObject()](/javascript/api/excel/excel.datavalidation#getinvalidcellsornullobject--)|Retorna um RangeAreas, que consiste em um ou mais intervalos retangulares, com valores inválidos de célula. Se todos os valores de célula forem válidos, essa função retornará null.|
|[FilterCriteria](/javascript/api/excel/excel.filtercriteria)|[subField](/javascript/api/excel/excel.filtercriteria#subfield)|A propriedade usada pelo filtro para realizar a filtragem avançada em richvalues.|
|[GeometricShape](/javascript/api/excel/excel.geometricshape)|[id](/javascript/api/excel/excel.geometricshape#id)|Retorna o identificador de forma. Somente leitura.|
||[shape](/javascript/api/excel/excel.geometricshape#shape)|Retorna o objeto de Forma para a forma geométrica. Somente leitura.|
|[GeometricShapeData](/javascript/api/excel/excel.geometricshapedata)|[id](/javascript/api/excel/excel.geometricshapedata#id)|Retorna o identificador de forma. Somente leitura.|
|[GeometricShapeLoadOptions](/javascript/api/excel/excel.geometricshapeloadoptions)|[$all](/javascript/api/excel/excel.geometricshapeloadoptions#$all)||
||[id](/javascript/api/excel/excel.geometricshapeloadoptions#id)|Retorna o identificador de forma. Somente leitura.|
||[shape](/javascript/api/excel/excel.geometricshapeloadoptions#shape)|Retorna o objeto de Forma para a forma geométrica.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getCount()](/javascript/api/excel/excel.groupshapecollection#getcount--)|Retorna o número de formas no grupo de forma. Somente leitura.|
||[getItem(key: string)](/javascript/api/excel/excel.groupshapecollection#getitem-key-)|Obtém uma forma usando seu nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.groupshapecollection#getitemat-index-)|Obtém uma forma com base em sua posição na coleção.|
||[items](/javascript/api/excel/excel.groupshapecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[GroupShapeCollectionLoadOptions](/javascript/api/excel/excel.groupshapecollectionloadoptions)|[$all](/javascript/api/excel/excel.groupshapecollectionloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.groupshapecollectionloadoptions#alttextdescription)|Para cada ITEM na coleção: Retorna ou define o texto de descrição alternativa para um objeto Shape.|
||[altTextTitle](/javascript/api/excel/excel.groupshapecollectionloadoptions#alttexttitle)|Para cada ITEM na coleção: Retorna ou define o texto de título alternativo para um objeto Shape.|
||[connectionSiteCount](/javascript/api/excel/excel.groupshapecollectionloadoptions#connectionsitecount)|Para cada ITEM na coleção: retorna o número de sites de conexão nesta forma. Somente leitura.|
||[fill](/javascript/api/excel/excel.groupshapecollectionloadoptions#fill)|Para cada ITEM na coleção: retorna a formatação de preenchimento dessa forma.|
||[geometricShape](/javascript/api/excel/excel.groupshapecollectionloadoptions#geometricshape)|Para cada ITEM na coleção: retorna a forma geométrica associada à forma. Um erro será lançado, se o tipo de forma não for "GeometricShape".|
||[geometricShapeType](/javascript/api/excel/excel.groupshapecollectionloadoptions#geometricshapetype)|Para cada ITEM na coleção: representa o tipo de forma geométrica dessa forma geométrica. Para saber detalhes, confira Excel.GeometricShapeType. Retorna nulo se o tipo de forma não for "GeometricShape".|
||[group](/javascript/api/excel/excel.groupshapecollectionloadoptions#group)|Para cada ITEM na coleção: retorna o grupo de formas associado à forma. Um erro será lançado, se o tipo de forma não for "GroupShape".|
||[height](/javascript/api/excel/excel.groupshapecollectionloadoptions#height)|Para cada ITEM na coleção: representa a altura, em pontos, da forma.|
||[id](/javascript/api/excel/excel.groupshapecollectionloadoptions#id)|Para cada ITEM na coleção: representa o identificador da forma. Somente leitura.|
||[image](/javascript/api/excel/excel.groupshapecollectionloadoptions#image)|Para cada ITEM na coleção: retorna a imagem associada à forma. Um erro será lançado, se o tipo de forma não for "Imagem".|
||[left](/javascript/api/excel/excel.groupshapecollectionloadoptions#left)|Para cada ITEM na coleção: a distância, em pontos, do lado esquerdo da forma até o lado esquerdo da planilha.|
||[level](/javascript/api/excel/excel.groupshapecollectionloadoptions#level)|Para cada ITEM na coleção: representa o nível da forma especificada. Por exemplo, um nível de 0 significa que a forma não faz parte de nenhum grupo, um nível de 1 significa que a forma é parte de um grupo de nível superior e um nível 2 significa que a forma faz parte de um subgrupo do nível superior.|
||[line](/javascript/api/excel/excel.groupshapecollectionloadoptions#line)|Para cada ITEM na coleção: retorna a linha associada à forma. Um erro será lançado, se o tipo de forma não for "Linha".|
||[lineFormat](/javascript/api/excel/excel.groupshapecollectionloadoptions#lineformat)|Para cada ITEM na coleção: retorna a formatação de linha dessa forma.|
||[lockAspectRatio](/javascript/api/excel/excel.groupshapecollectionloadoptions#lockaspectratio)|Para cada ITEM na coleção: especifica se a taxa de proporção dessa forma será ou não bloqueada.|
||[name](/javascript/api/excel/excel.groupshapecollectionloadoptions#name)|Para cada ITEM na coleção: representa o nome da forma.|
||[parentGroup](/javascript/api/excel/excel.groupshapecollectionloadoptions#parentgroup)|Para cada ITEM na coleção: representa o grupo pai desta forma.|
||[rotation](/javascript/api/excel/excel.groupshapecollectionloadoptions#rotation)|Para cada ITEM na coleção: representa a rotação, em graus, da forma.|
||[textFrame](/javascript/api/excel/excel.groupshapecollectionloadoptions#textframe)|Para cada ITEM na coleção: retorna o objeto de quadro de texto desta forma. Somente leitura.|
||[top](/javascript/api/excel/excel.groupshapecollectionloadoptions#top)|Para cada ITEM na coleção: a distância, em pontos, da borda superior da forma até a borda superior da planilha.|
||[tipo](/javascript/api/excel/excel.groupshapecollectionloadoptions#type)|Para cada ITEM na coleção: retorna o tipo dessa forma. Para saber detalhes, confira Excel.ShapeType. Somente leitura.|
||[visible](/javascript/api/excel/excel.groupshapecollectionloadoptions#visible)|Para cada ITEM na coleção: representa a visibilidade dessa forma.|
||[width](/javascript/api/excel/excel.groupshapecollectionloadoptions#width)|Para cada ITEM na coleção: representa a largura, em pontos, da forma.|
||[zOrderPosition](/javascript/api/excel/excel.groupshapecollectionloadoptions#zorderposition)|Para cada ITEM na coleção: retorna a posição da forma especificada na ordem z, com 0 que representa a parte inferior da pilha da ordem. Somente leitura.|
|[HeaderFooter](/javascript/api/excel/excel.headerfooter)|[centerFooter](/javascript/api/excel/excel.headerfooter#centerfooter)|Obtém ou define o rodapé central da planilha.|
||[centerHeader](/javascript/api/excel/excel.headerfooter#centerheader)|Obtém ou define o cabeçalho central da planilha.|
||[leftFooter](/javascript/api/excel/excel.headerfooter#leftfooter)|Obtém ou define o rodapé esquerdo da planilha.|
||[leftHeader](/javascript/api/excel/excel.headerfooter#leftheader)|Obtém ou define o cabeçalho esquerdo da planilha.|
||[rightFooter](/javascript/api/excel/excel.headerfooter#rightfooter)|Obtém ou define o rodapé direito da planilha.|
||[rightHeader](/javascript/api/excel/excel.headerfooter#rightheader)|Obtém ou define o cabeçalho direito da planilha.|
||[Set (Propriedades: Excel. HeaderFooter)](/javascript/api/excel/excel.headerfooter#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. HeaderFooterUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.headerfooter#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[HeaderFooterData](/javascript/api/excel/excel.headerfooterdata)|[centerFooter](/javascript/api/excel/excel.headerfooterdata#centerfooter)|Obtém ou define o rodapé central da planilha.|
||[centerHeader](/javascript/api/excel/excel.headerfooterdata#centerheader)|Obtém ou define o cabeçalho central da planilha.|
||[leftFooter](/javascript/api/excel/excel.headerfooterdata#leftfooter)|Obtém ou define o rodapé esquerdo da planilha.|
||[leftHeader](/javascript/api/excel/excel.headerfooterdata#leftheader)|Obtém ou define o cabeçalho esquerdo da planilha.|
||[rightFooter](/javascript/api/excel/excel.headerfooterdata#rightfooter)|Obtém ou define o rodapé direito da planilha.|
||[rightHeader](/javascript/api/excel/excel.headerfooterdata#rightheader)|Obtém ou define o cabeçalho direito da planilha.|
|[HeaderFooterGroup](/javascript/api/excel/excel.headerfootergroup)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroup#defaultforallpages)|O cabeçalho/rodapé geral, usado em todas as páginas, a menos que seja especificada a página par/ímpar ou a primeira página.|
||[evenPages](/javascript/api/excel/excel.headerfootergroup#evenpages)|O cabeçalho/rodapé a ser usado para páginas pares, o cabeçalho/rodapé ímpar deve ser especificado para páginas ímpares.|
||[firstPage](/javascript/api/excel/excel.headerfootergroup#firstpage)|O cabeçalho/rodapé da primeira página. Para todas as outras páginas, geral ou par/ímpar é usado.|
||[oddPages](/javascript/api/excel/excel.headerfootergroup#oddpages)|O cabeçalho/rodapé a ser usado para páginas ímpares, o cabeçalho/rodapé par deve ser especificado para páginas pares.|
||[Set (Propriedades: Excel. HeaderFooterGroup)](/javascript/api/excel/excel.headerfootergroup#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. HeaderFooterGroupUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.headerfootergroup#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[state](/javascript/api/excel/excel.headerfootergroup#state)|Obtém ou define o estado do qual os cabeçalhos/rodapés são definidos. Para saber detalhes, confira Excel.HeaderFooterState.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroup#usesheetmargins)|Obtém ou define um sinalizador indicando se os cabeçalhos/rodapés estão alinhados com as margens da página que foram definidas nas opções de layout de página da planilha.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroup#usesheetscale)|Obtém ou define um sinalizador que indica se os cabeçalhos/rodapés devem ser dimensionados pela escala de porcentagem da página definida nas opções de layout de página da planilha.|
|[HeaderFooterGroupData](/javascript/api/excel/excel.headerfootergroupdata)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroupdata#defaultforallpages)|O cabeçalho/rodapé geral, usado em todas as páginas, a menos que seja especificada a página par/ímpar ou a primeira página.|
||[evenPages](/javascript/api/excel/excel.headerfootergroupdata#evenpages)|O cabeçalho/rodapé a ser usado para páginas pares, o cabeçalho/rodapé ímpar deve ser especificado para páginas ímpares.|
||[firstPage](/javascript/api/excel/excel.headerfootergroupdata#firstpage)|O cabeçalho/rodapé da primeira página. Para todas as outras páginas, geral ou par/ímpar é usado.|
||[oddPages](/javascript/api/excel/excel.headerfootergroupdata#oddpages)|O cabeçalho/rodapé a ser usado para páginas ímpares, o cabeçalho/rodapé par deve ser especificado para páginas pares.|
||[state](/javascript/api/excel/excel.headerfootergroupdata#state)|Obtém ou define o estado do qual os cabeçalhos/rodapés são definidos. Para saber detalhes, confira Excel.HeaderFooterState.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroupdata#usesheetmargins)|Obtém ou define um sinalizador indicando se os cabeçalhos/rodapés estão alinhados com as margens da página que foram definidas nas opções de layout de página da planilha.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroupdata#usesheetscale)|Obtém ou define um sinalizador que indica se os cabeçalhos/rodapés devem ser dimensionados pela escala de porcentagem da página definida nas opções de layout de página da planilha.|
|[HeaderFooterGroupLoadOptions](/javascript/api/excel/excel.headerfootergrouploadoptions)|[$all](/javascript/api/excel/excel.headerfootergrouploadoptions#$all)||
||[defaultForAllPages](/javascript/api/excel/excel.headerfootergrouploadoptions#defaultforallpages)|O cabeçalho/rodapé geral, usado em todas as páginas, a menos que seja especificada a página par/ímpar ou a primeira página.|
||[evenPages](/javascript/api/excel/excel.headerfootergrouploadoptions#evenpages)|O cabeçalho/rodapé a ser usado para páginas pares, o cabeçalho/rodapé ímpar deve ser especificado para páginas ímpares.|
||[firstPage](/javascript/api/excel/excel.headerfootergrouploadoptions#firstpage)|O cabeçalho/rodapé da primeira página. Para todas as outras páginas, geral ou par/ímpar é usado.|
||[oddPages](/javascript/api/excel/excel.headerfootergrouploadoptions#oddpages)|O cabeçalho/rodapé a ser usado para páginas ímpares, o cabeçalho/rodapé par deve ser especificado para páginas pares.|
||[state](/javascript/api/excel/excel.headerfootergrouploadoptions#state)|Obtém ou define o estado do qual os cabeçalhos/rodapés são definidos. Para saber detalhes, confira Excel.HeaderFooterState.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergrouploadoptions#usesheetmargins)|Obtém ou define um sinalizador indicando se os cabeçalhos/rodapés estão alinhados com as margens da página que foram definidas nas opções de layout de página da planilha.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergrouploadoptions#usesheetscale)|Obtém ou define um sinalizador que indica se os cabeçalhos/rodapés devem ser dimensionados pela escala de porcentagem da página definida nas opções de layout de página da planilha.|
|[HeaderFooterGroupUpdateData](/javascript/api/excel/excel.headerfootergroupupdatedata)|[defaultForAllPages](/javascript/api/excel/excel.headerfootergroupupdatedata#defaultforallpages)|O cabeçalho/rodapé geral, usado em todas as páginas, a menos que seja especificada a página par/ímpar ou a primeira página.|
||[evenPages](/javascript/api/excel/excel.headerfootergroupupdatedata#evenpages)|O cabeçalho/rodapé a ser usado para páginas pares, o cabeçalho/rodapé ímpar deve ser especificado para páginas ímpares.|
||[firstPage](/javascript/api/excel/excel.headerfootergroupupdatedata#firstpage)|O cabeçalho/rodapé da primeira página. Para todas as outras páginas, geral ou par/ímpar é usado.|
||[oddPages](/javascript/api/excel/excel.headerfootergroupupdatedata#oddpages)|O cabeçalho/rodapé a ser usado para páginas ímpares, o cabeçalho/rodapé par deve ser especificado para páginas pares.|
||[state](/javascript/api/excel/excel.headerfootergroupupdatedata#state)|Obtém ou define o estado do qual os cabeçalhos/rodapés são definidos. Para saber detalhes, confira Excel.HeaderFooterState.|
||[useSheetMargins](/javascript/api/excel/excel.headerfootergroupupdatedata#usesheetmargins)|Obtém ou define um sinalizador indicando se os cabeçalhos/rodapés estão alinhados com as margens da página que foram definidas nas opções de layout de página da planilha.|
||[useSheetScale](/javascript/api/excel/excel.headerfootergroupupdatedata#usesheetscale)|Obtém ou define um sinalizador que indica se os cabeçalhos/rodapés devem ser dimensionados pela escala de porcentagem da página definida nas opções de layout de página da planilha.|
|[HeaderFooterLoadOptions](/javascript/api/excel/excel.headerfooterloadoptions)|[$all](/javascript/api/excel/excel.headerfooterloadoptions#$all)||
||[centerFooter](/javascript/api/excel/excel.headerfooterloadoptions#centerfooter)|Obtém ou define o rodapé central da planilha.|
||[centerHeader](/javascript/api/excel/excel.headerfooterloadoptions#centerheader)|Obtém ou define o cabeçalho central da planilha.|
||[leftFooter](/javascript/api/excel/excel.headerfooterloadoptions#leftfooter)|Obtém ou define o rodapé esquerdo da planilha.|
||[leftHeader](/javascript/api/excel/excel.headerfooterloadoptions#leftheader)|Obtém ou define o cabeçalho esquerdo da planilha.|
||[rightFooter](/javascript/api/excel/excel.headerfooterloadoptions#rightfooter)|Obtém ou define o rodapé direito da planilha.|
||[rightHeader](/javascript/api/excel/excel.headerfooterloadoptions#rightheader)|Obtém ou define o cabeçalho direito da planilha.|
|[HeaderFooterUpdateData](/javascript/api/excel/excel.headerfooterupdatedata)|[centerFooter](/javascript/api/excel/excel.headerfooterupdatedata#centerfooter)|Obtém ou define o rodapé central da planilha.|
||[centerHeader](/javascript/api/excel/excel.headerfooterupdatedata#centerheader)|Obtém ou define o cabeçalho central da planilha.|
||[leftFooter](/javascript/api/excel/excel.headerfooterupdatedata#leftfooter)|Obtém ou define o rodapé esquerdo da planilha.|
||[leftHeader](/javascript/api/excel/excel.headerfooterupdatedata#leftheader)|Obtém ou define o cabeçalho esquerdo da planilha.|
||[rightFooter](/javascript/api/excel/excel.headerfooterupdatedata#rightfooter)|Obtém ou define o rodapé direito da planilha.|
||[rightHeader](/javascript/api/excel/excel.headerfooterupdatedata#rightheader)|Obtém ou define o cabeçalho direito da planilha.|
|[Image](/javascript/api/excel/excel.image)|[format](/javascript/api/excel/excel.image#format)|Retorna o formato da imagem. Somente leitura.|
||[id](/javascript/api/excel/excel.image#id)|Representa o identificador de forma para o objeto de imagem. Somente leitura.|
||[shape](/javascript/api/excel/excel.image#shape)|Retorna o objeto de forma associado à imagem. Somente leitura.|
|[ImageData](/javascript/api/excel/excel.imagedata)|[format](/javascript/api/excel/excel.imagedata#format)|Retorna o formato da imagem. Somente leitura.|
||[id](/javascript/api/excel/excel.imagedata#id)|Representa o identificador de forma para o objeto de imagem. Somente leitura.|
|[ImageLoadOptions](/javascript/api/excel/excel.imageloadoptions)|[$all](/javascript/api/excel/excel.imageloadoptions#$all)||
||[format](/javascript/api/excel/excel.imageloadoptions#format)|Retorna o formato da imagem. Somente leitura.|
||[id](/javascript/api/excel/excel.imageloadoptions#id)|Representa o identificador de forma para o objeto de imagem. Somente leitura.|
||[shape](/javascript/api/excel/excel.imageloadoptions#shape)|Retorna o objeto de forma associado à imagem.|
|[IterativeCalculation](/javascript/api/excel/excel.iterativecalculation)|[enabled](/javascript/api/excel/excel.iterativecalculation#enabled)|True se o Excel usará a interação para resolver referências circulares.|
||[maxChange](/javascript/api/excel/excel.iterativecalculation#maxchange)|Retorna ou define a quantidade máxima de alteração entre cada iteração conforme o Excel resolve referências circulares.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculation#maxiteration)|Retorna ou define o número máximo de iterações que o Excel pode usar para resolver uma referência circular.|
||[Set (Propriedades: Excel. IterativeCalculation)](/javascript/api/excel/excel.iterativecalculation#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. IterativeCalculationUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.iterativecalculation#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[IterativeCalculationData](/javascript/api/excel/excel.iterativecalculationdata)|[enabled](/javascript/api/excel/excel.iterativecalculationdata#enabled)|True se o Excel usará a interação para resolver referências circulares.|
||[maxChange](/javascript/api/excel/excel.iterativecalculationdata#maxchange)|Retorna ou define a quantidade máxima de alteração entre cada iteração conforme o Excel resolve referências circulares.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationdata#maxiteration)|Retorna ou define o número máximo de iterações que o Excel pode usar para resolver uma referência circular.|
|[IterativeCalculationLoadOptions](/javascript/api/excel/excel.iterativecalculationloadoptions)|[$all](/javascript/api/excel/excel.iterativecalculationloadoptions#$all)||
||[enabled](/javascript/api/excel/excel.iterativecalculationloadoptions#enabled)|True se o Excel usará a interação para resolver referências circulares.|
||[maxChange](/javascript/api/excel/excel.iterativecalculationloadoptions#maxchange)|Retorna ou define a quantidade máxima de alteração entre cada iteração conforme o Excel resolve referências circulares.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationloadoptions#maxiteration)|Retorna ou define o número máximo de iterações que o Excel pode usar para resolver uma referência circular.|
|[IterativeCalculationUpdateData](/javascript/api/excel/excel.iterativecalculationupdatedata)|[enabled](/javascript/api/excel/excel.iterativecalculationupdatedata#enabled)|True se o Excel usará a interação para resolver referências circulares.|
||[maxChange](/javascript/api/excel/excel.iterativecalculationupdatedata#maxchange)|Retorna ou define a quantidade máxima de alteração entre cada iteração conforme o Excel resolve referências circulares.|
||[maxIteration](/javascript/api/excel/excel.iterativecalculationupdatedata#maxiteration)|Retorna ou define o número máximo de iterações que o Excel pode usar para resolver uma referência circular.|
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
||[beginConnectedShape](/javascript/api/excel/excel.line#beginconnectedshape)|Representa a forma na qual o início da linha especificada está conectado. Somente leitura.|
||[beginConnectedSite](/javascript/api/excel/excel.line#beginconnectedsite)|Representa o site de conexão ao qual o início de um conector está conectado. Somente leitura. Retorna nulo quando o início da linha não está conectado a qualquer forma.|
||[endConnectedShape](/javascript/api/excel/excel.line#endconnectedshape)|Representa a forma na qual o final da linha especificada está conectado. Somente leitura.|
||[endConnectedSite](/javascript/api/excel/excel.line#endconnectedsite)|Representa o site de conexão ao qual o final de um conector está conectado. Somente leitura. Retorna nulo quando o final da linha não está conectado a qualquer forma.|
||[id](/javascript/api/excel/excel.line#id)|Representa o identificador de forma. Somente leitura.|
||[isBeginConnected](/javascript/api/excel/excel.line#isbeginconnected)|Especifica se o início do conector especificado está conectado ou não a uma forma. Somente leitura.|
||[isEndConnected](/javascript/api/excel/excel.line#isendconnected)|Especifica se o final do conector especificado está conectado ou não a uma forma. Somente leitura.|
||[shape](/javascript/api/excel/excel.line#shape)|Retorna o objeto de forma associado à linha. Somente leitura.|
||[Set (Propriedades: Excel. line)](/javascript/api/excel/excel.line#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. LineUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.line#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[LineData](/javascript/api/excel/excel.linedata)|[beginArrowheadLength](/javascript/api/excel/excel.linedata#beginarrowheadlength)|Representa o comprimento da ponta da seta no início da linha especificada.|
||[beginArrowheadStyle](/javascript/api/excel/excel.linedata#beginarrowheadstyle)|Representa o estilo da ponta de seta no início da linha especificada.|
||[BeginArrowheadWidth](/javascript/api/excel/excel.linedata#beginarrowheadwidth)|Representa a largura da ponta da seta no início da linha especificada.|
||[beginConnectedSite](/javascript/api/excel/excel.linedata#beginconnectedsite)|Representa o site de conexão ao qual o início de um conector está conectado. Somente leitura. Retorna nulo quando o início da linha não está conectado a qualquer forma.|
||[connectorType](/javascript/api/excel/excel.linedata#connectortype)|Representa o tipo de conector de linha.|
||[endArrowheadLength](/javascript/api/excel/excel.linedata#endarrowheadlength)|Representa o comprimento da ponta de seta no final da linha especificada.|
||[endArrowheadStyle](/javascript/api/excel/excel.linedata#endarrowheadstyle)|Representa o estilo da ponta de seta no final da linha especificada.|
||[endArrowheadWidth](/javascript/api/excel/excel.linedata#endarrowheadwidth)|Representa a largura da ponta de seta no final da linha especificada.|
||[endConnectedSite](/javascript/api/excel/excel.linedata#endconnectedsite)|Representa o site de conexão ao qual o final de um conector está conectado. Somente leitura. Retorna nulo quando o final da linha não está conectado a qualquer forma.|
||[id](/javascript/api/excel/excel.linedata#id)|Representa o identificador de forma. Somente leitura.|
||[isBeginConnected](/javascript/api/excel/excel.linedata#isbeginconnected)|Especifica se o início do conector especificado está conectado ou não a uma forma. Somente leitura.|
||[isEndConnected](/javascript/api/excel/excel.linedata#isendconnected)|Especifica se o final do conector especificado está conectado ou não a uma forma. Somente leitura.|
|[LineLoadOptions](/javascript/api/excel/excel.lineloadoptions)|[$all](/javascript/api/excel/excel.lineloadoptions#$all)||
||[beginArrowheadLength](/javascript/api/excel/excel.lineloadoptions#beginarrowheadlength)|Representa o comprimento da ponta da seta no início da linha especificada.|
||[beginArrowheadStyle](/javascript/api/excel/excel.lineloadoptions#beginarrowheadstyle)|Representa o estilo da ponta de seta no início da linha especificada.|
||[BeginArrowheadWidth](/javascript/api/excel/excel.lineloadoptions#beginarrowheadwidth)|Representa a largura da ponta da seta no início da linha especificada.|
||[beginConnectedShape](/javascript/api/excel/excel.lineloadoptions#beginconnectedshape)|Representa a forma na qual o início da linha especificada está conectado.|
||[beginConnectedSite](/javascript/api/excel/excel.lineloadoptions#beginconnectedsite)|Representa o site de conexão ao qual o início de um conector está conectado. Somente leitura. Retorna nulo quando o início da linha não está conectado a qualquer forma.|
||[connectorType](/javascript/api/excel/excel.lineloadoptions#connectortype)|Representa o tipo de conector de linha.|
||[endArrowheadLength](/javascript/api/excel/excel.lineloadoptions#endarrowheadlength)|Representa o comprimento da ponta de seta no final da linha especificada.|
||[endArrowheadStyle](/javascript/api/excel/excel.lineloadoptions#endarrowheadstyle)|Representa o estilo da ponta de seta no final da linha especificada.|
||[endArrowheadWidth](/javascript/api/excel/excel.lineloadoptions#endarrowheadwidth)|Representa a largura da ponta de seta no final da linha especificada.|
||[endConnectedShape](/javascript/api/excel/excel.lineloadoptions#endconnectedshape)|Representa a forma na qual o final da linha especificada está conectado.|
||[endConnectedSite](/javascript/api/excel/excel.lineloadoptions#endconnectedsite)|Representa o site de conexão ao qual o final de um conector está conectado. Somente leitura. Retorna nulo quando o final da linha não está conectado a qualquer forma.|
||[id](/javascript/api/excel/excel.lineloadoptions#id)|Representa o identificador de forma. Somente leitura.|
||[isBeginConnected](/javascript/api/excel/excel.lineloadoptions#isbeginconnected)|Especifica se o início do conector especificado está conectado ou não a uma forma. Somente leitura.|
||[isEndConnected](/javascript/api/excel/excel.lineloadoptions#isendconnected)|Especifica se o final do conector especificado está conectado ou não a uma forma. Somente leitura.|
||[shape](/javascript/api/excel/excel.lineloadoptions#shape)|Retorna o objeto de forma associado à linha.|
|[LineUpdateData](/javascript/api/excel/excel.lineupdatedata)|[beginArrowheadLength](/javascript/api/excel/excel.lineupdatedata#beginarrowheadlength)|Representa o comprimento da ponta da seta no início da linha especificada.|
||[beginArrowheadStyle](/javascript/api/excel/excel.lineupdatedata#beginarrowheadstyle)|Representa o estilo da ponta de seta no início da linha especificada.|
||[BeginArrowheadWidth](/javascript/api/excel/excel.lineupdatedata#beginarrowheadwidth)|Representa a largura da ponta da seta no início da linha especificada.|
||[connectorType](/javascript/api/excel/excel.lineupdatedata#connectortype)|Representa o tipo de conector de linha.|
||[endArrowheadLength](/javascript/api/excel/excel.lineupdatedata#endarrowheadlength)|Representa o comprimento da ponta de seta no final da linha especificada.|
||[endArrowheadStyle](/javascript/api/excel/excel.lineupdatedata#endarrowheadstyle)|Representa o estilo da ponta de seta no final da linha especificada.|
||[endArrowheadWidth](/javascript/api/excel/excel.lineupdatedata#endarrowheadwidth)|Representa a largura da ponta de seta no final da linha especificada.|
|[PageBreak](/javascript/api/excel/excel.pagebreak)|[delete()](/javascript/api/excel/excel.pagebreak#delete--)|Exclui um objeto de quebra de página.|
||[getCellAfterBreak()](/javascript/api/excel/excel.pagebreak#getcellafterbreak--)|Obtém a primeira célula após a quebra de página.|
||[columnIndex](/javascript/api/excel/excel.pagebreak#columnindex)|Representa o índice de coluna para a quebra de página|
||[rowIndex](/javascript/api/excel/excel.pagebreak#rowindex)|Representa o índice de linha para a quebra de página|
|[PageBreakCollection](/javascript/api/excel/excel.pagebreakcollection)|[add(pageBreakRange: Range \| string)](/javascript/api/excel/excel.pagebreakcollection#add-pagebreakrange-)|Adiciona uma quebra de página antes da célula superior esquerda do intervalo especificado.|
||[getCount()](/javascript/api/excel/excel.pagebreakcollection#getcount--)|Obtém o número de quebras de página na coleção.|
||[getItem(index: number)](/javascript/api/excel/excel.pagebreakcollection#getitem-index-)|Obtém um objeto de quebra de página através do índice.|
||[items](/javascript/api/excel/excel.pagebreakcollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[removePageBreaks()](/javascript/api/excel/excel.pagebreakcollection#removepagebreaks--)|Redefine todas as quebras de página manuais na coleção.|
|[PageBreakCollectionLoadOptions](/javascript/api/excel/excel.pagebreakcollectionloadoptions)|[$all](/javascript/api/excel/excel.pagebreakcollectionloadoptions#$all)||
||[columnIndex](/javascript/api/excel/excel.pagebreakcollectionloadoptions#columnindex)|Para cada ITEM na coleção: representa o índice de coluna da quebra de página|
||[rowIndex](/javascript/api/excel/excel.pagebreakcollectionloadoptions#rowindex)|Para cada ITEM na coleção: representa o índice de linha da quebra de página|
|[PageBreakData](/javascript/api/excel/excel.pagebreakdata)|[columnIndex](/javascript/api/excel/excel.pagebreakdata#columnindex)|Representa o índice de coluna para a quebra de página|
||[rowIndex](/javascript/api/excel/excel.pagebreakdata#rowindex)|Representa o índice de linha para a quebra de página|
|[PageBreakLoadOptions](/javascript/api/excel/excel.pagebreakloadoptions)|[$all](/javascript/api/excel/excel.pagebreakloadoptions#$all)||
||[columnIndex](/javascript/api/excel/excel.pagebreakloadoptions#columnindex)|Representa o índice de coluna para a quebra de página|
||[rowIndex](/javascript/api/excel/excel.pagebreakloadoptions#rowindex)|Representa o índice de linha para a quebra de página|
|[PageLayout](/javascript/api/excel/excel.pagelayout)|[blackAndWhite](/javascript/api/excel/excel.pagelayout#blackandwhite)|Obtém ou define a opção de impressão em preto e branco da planilha.|
||[bottomMargin](/javascript/api/excel/excel.pagelayout#bottommargin)|Obtém ou define a margem de página inferior da planilha para impressão em pontos.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayout#centerhorizontally)|Obtém ou define o sinalizador de centralização horizontal da planilha. Esse sinalizador determina se a planilha será centralizada horizontalmente quando for impressa.|
||[centerVertically](/javascript/api/excel/excel.pagelayout#centervertically)|Obtém ou define o sinalizador de centralização vertical da planilha. Esse sinalizador determina se a planilha será centralizada verticalmente quando for impressa.|
||[draftMode](/javascript/api/excel/excel.pagelayout#draftmode)|Obtém ou define a opção de modo de rascunho da planilha. Se for true, a planilha será impressa sem gráficos.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayout#firstpagenumber)|Obtém ou define o primeiro número de página da planilha a ser impressa. O valor null representa a numeração "automática" de páginas.|
||[footerMargin](/javascript/api/excel/excel.pagelayout#footermargin)|Obtém ou define a margem do rodapé da planilha, em pontos, para usar durante a impressão.|
||[getPrintArea()](/javascript/api/excel/excel.pagelayout#getprintarea--)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos retangulares, que representa a área de impressão da planilha. Se não houver uma área de impressão, um erro ItemNotFound será gerado.|
||[getPrintAreaOrNullObject()](/javascript/api/excel/excel.pagelayout#getprintareaornullobject--)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos retangulares, que representa a área de impressão da planilha. Se não houver uma área de impressão, um objeto null será retornado.|
||[getPrintTitleColumns()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumns--)|Obtém o objeto range que representa as colunas de título.|
||[getPrintTitleColumnsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlecolumnsornullobject--)|Obtém o objeto range que representa as colunas de título. Se não estiver configurado, retornará um objeto null.|
||[getPrintTitleRows()](/javascript/api/excel/excel.pagelayout#getprinttitlerows--)|Obtém o objeto range representando as linhas do título.|
||[getPrintTitleRowsOrNullObject()](/javascript/api/excel/excel.pagelayout#getprinttitlerowsornullobject--)|Obtém o objeto range representando as linhas do título. Se não estiver configurado, retornará um objeto null.|
||[headerMargin](/javascript/api/excel/excel.pagelayout#headermargin)|Obtém ou define a margem do cabeçalho da planilha, em pontos, para usar durante a impressão.|
||[leftMargin](/javascript/api/excel/excel.pagelayout#leftmargin)|Obtém ou define a margem esquerda da planilha, em pontos, para usar durante a impressão.|
||[orientation](/javascript/api/excel/excel.pagelayout#orientation)|Obtém ou define a orientação de página da planilha.|
||[paperSize](/javascript/api/excel/excel.pagelayout#papersize)|Obtém ou define o tamanho do papel da página da planilha.|
||[printComments](/javascript/api/excel/excel.pagelayout#printcomments)|Obtém ou define se os comentários da planilha deverão ser exibidos durante a impressão.|
||[printErrors](/javascript/api/excel/excel.pagelayout#printerrors)|Obtém ou define a opção de erros de impressão da planilha.|
||[printGridlines](/javascript/api/excel/excel.pagelayout#printgridlines)|Obtém ou define um sinalizador de linhas de grade de impressão da planilha. Esse sinalizador determina se as linhas de grade serão impressas ou não.|
||[printHeadings](/javascript/api/excel/excel.pagelayout#printheadings)|Obtém ou define um sinalizador de cabeçalhos de impressão da planilha. Esse sinalizador determina se os cabeçalhos serão impressos ou não.|
||[printOrder](/javascript/api/excel/excel.pagelayout#printorder)|Obtém ou define a opção de ordem de impressão da página da planilha. Isso especifica a ordem que será usada para processar o número de página impresso.|
||[headersFooters](/javascript/api/excel/excel.pagelayout#headersfooters)|Configuração de cabeçalho e rodapé da planilha.|
||[rightMargin](/javascript/api/excel/excel.pagelayout#rightmargin)|Obtém ou define a margem direita da planilha, em pontos, para usar durante a impressão.|
||[Set (Propriedades: Excel. PageLayout)](/javascript/api/excel/excel.pagelayout#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. PageLayoutUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.pagelayout#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[setPrintArea(printArea: Range \| RangeAreas \| string)](/javascript/api/excel/excel.pagelayout#setprintarea-printarea-)|Define a área de impressão da planilha.|
||[setPrintMargins(unit: "Points" \| "Inches" \| "Centimeters", marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Define as margens das páginas da planilha com unidades.|
||[setPrintMargins(unit: Excel.PrintMarginUnit, marginOptions: Excel.PageLayoutMarginOptions)](/javascript/api/excel/excel.pagelayout#setprintmargins-unit--marginoptions-)|Define as margens das páginas da planilha com unidades.|
||[setPrintTitleColumns(printTitleColumns: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlecolumns-printtitlecolumns-)|Define as colunas que contêm as células que serão repetidas à esquerda de cada página da planilha para impressão.|
||[setPrintTitleRows(printTitleRows: Range \| string)](/javascript/api/excel/excel.pagelayout#setprinttitlerows-printtitlerows-)|Define as linhas que contêm as células que serão repetidas na parte de cada página da planilha para impressão.|
||[topMargin](/javascript/api/excel/excel.pagelayout#topmargin)|Obtém ou define a margem superior da planilha, em pontos, para usar durante a impressão.|
||[zoom](/javascript/api/excel/excel.pagelayout#zoom)|Obtém ou define as opções de zoom de impressão da planilha.|
|[PageLayoutData](/javascript/api/excel/excel.pagelayoutdata)|[blackAndWhite](/javascript/api/excel/excel.pagelayoutdata#blackandwhite)|Obtém ou define a opção de impressão em preto e branco da planilha.|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutdata#bottommargin)|Obtém ou define a margem de página inferior da planilha para impressão em pontos.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutdata#centerhorizontally)|Obtém ou define o sinalizador de centralização horizontal da planilha. Esse sinalizador determina se a planilha será centralizada horizontalmente quando for impressa.|
||[centerVertically](/javascript/api/excel/excel.pagelayoutdata#centervertically)|Obtém ou define o sinalizador de centralização vertical da planilha. Esse sinalizador determina se a planilha será centralizada verticalmente quando for impressa.|
||[draftMode](/javascript/api/excel/excel.pagelayoutdata#draftmode)|Obtém ou define a opção de modo de rascunho da planilha. Se for true, a planilha será impressa sem gráficos.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutdata#firstpagenumber)|Obtém ou define o primeiro número de página da planilha a ser impressa. O valor null representa a numeração "automática" de páginas.|
||[footerMargin](/javascript/api/excel/excel.pagelayoutdata#footermargin)|Obtém ou define a margem do rodapé da planilha, em pontos, para usar durante a impressão.|
||[headerMargin](/javascript/api/excel/excel.pagelayoutdata#headermargin)|Obtém ou define a margem do cabeçalho da planilha, em pontos, para usar durante a impressão.|
||[headersFooters](/javascript/api/excel/excel.pagelayoutdata#headersfooters)|Configuração de cabeçalho e rodapé da planilha.|
||[leftMargin](/javascript/api/excel/excel.pagelayoutdata#leftmargin)|Obtém ou define a margem esquerda da planilha, em pontos, para usar durante a impressão.|
||[orientation](/javascript/api/excel/excel.pagelayoutdata#orientation)|Obtém ou define a orientação de página da planilha.|
||[paperSize](/javascript/api/excel/excel.pagelayoutdata#papersize)|Obtém ou define o tamanho do papel da página da planilha.|
||[printComments](/javascript/api/excel/excel.pagelayoutdata#printcomments)|Obtém ou define se os comentários da planilha deverão ser exibidos durante a impressão.|
||[printErrors](/javascript/api/excel/excel.pagelayoutdata#printerrors)|Obtém ou define a opção de erros de impressão da planilha.|
||[printGridlines](/javascript/api/excel/excel.pagelayoutdata#printgridlines)|Obtém ou define um sinalizador de linhas de grade de impressão da planilha. Esse sinalizador determina se as linhas de grade serão impressas ou não.|
||[printHeadings](/javascript/api/excel/excel.pagelayoutdata#printheadings)|Obtém ou define um sinalizador de cabeçalhos de impressão da planilha. Esse sinalizador determina se os cabeçalhos serão impressos ou não.|
||[printOrder](/javascript/api/excel/excel.pagelayoutdata#printorder)|Obtém ou define a opção de ordem de impressão da página da planilha. Isso especifica a ordem que será usada para processar o número de página impresso.|
||[rightMargin](/javascript/api/excel/excel.pagelayoutdata#rightmargin)|Obtém ou define a margem direita da planilha, em pontos, para usar durante a impressão.|
||[topMargin](/javascript/api/excel/excel.pagelayoutdata#topmargin)|Obtém ou define a margem superior da planilha, em pontos, para usar durante a impressão.|
||[zoom](/javascript/api/excel/excel.pagelayoutdata#zoom)|Obtém ou define as opções de zoom de impressão da planilha.|
|[PageLayoutLoadOptions](/javascript/api/excel/excel.pagelayoutloadoptions)|[$all](/javascript/api/excel/excel.pagelayoutloadoptions#$all)||
||[blackAndWhite](/javascript/api/excel/excel.pagelayoutloadoptions#blackandwhite)|Obtém ou define a opção de impressão em preto e branco da planilha.|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutloadoptions#bottommargin)|Obtém ou define a margem de página inferior da planilha para impressão em pontos.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutloadoptions#centerhorizontally)|Obtém ou define o sinalizador de centralização horizontal da planilha. Esse sinalizador determina se a planilha será centralizada horizontalmente quando for impressa.|
||[centerVertically](/javascript/api/excel/excel.pagelayoutloadoptions#centervertically)|Obtém ou define o sinalizador de centralização vertical da planilha. Esse sinalizador determina se a planilha será centralizada verticalmente quando for impressa.|
||[draftMode](/javascript/api/excel/excel.pagelayoutloadoptions#draftmode)|Obtém ou define a opção de modo de rascunho da planilha. Se for true, a planilha será impressa sem gráficos.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutloadoptions#firstpagenumber)|Obtém ou define o primeiro número de página da planilha a ser impressa. O valor null representa a numeração "automática" de páginas.|
||[footerMargin](/javascript/api/excel/excel.pagelayoutloadoptions#footermargin)|Obtém ou define a margem do rodapé da planilha, em pontos, para usar durante a impressão.|
||[headerMargin](/javascript/api/excel/excel.pagelayoutloadoptions#headermargin)|Obtém ou define a margem do cabeçalho da planilha, em pontos, para usar durante a impressão.|
||[headersFooters](/javascript/api/excel/excel.pagelayoutloadoptions#headersfooters)|Configuração de cabeçalho e rodapé da planilha.|
||[leftMargin](/javascript/api/excel/excel.pagelayoutloadoptions#leftmargin)|Obtém ou define a margem esquerda da planilha, em pontos, para usar durante a impressão.|
||[orientation](/javascript/api/excel/excel.pagelayoutloadoptions#orientation)|Obtém ou define a orientação de página da planilha.|
||[paperSize](/javascript/api/excel/excel.pagelayoutloadoptions#papersize)|Obtém ou define o tamanho do papel da página da planilha.|
||[printComments](/javascript/api/excel/excel.pagelayoutloadoptions#printcomments)|Obtém ou define se os comentários da planilha deverão ser exibidos durante a impressão.|
||[printErrors](/javascript/api/excel/excel.pagelayoutloadoptions#printerrors)|Obtém ou define a opção de erros de impressão da planilha.|
||[printGridlines](/javascript/api/excel/excel.pagelayoutloadoptions#printgridlines)|Obtém ou define um sinalizador de linhas de grade de impressão da planilha. Esse sinalizador determina se as linhas de grade serão impressas ou não.|
||[printHeadings](/javascript/api/excel/excel.pagelayoutloadoptions#printheadings)|Obtém ou define um sinalizador de cabeçalhos de impressão da planilha. Esse sinalizador determina se os cabeçalhos serão impressos ou não.|
||[printOrder](/javascript/api/excel/excel.pagelayoutloadoptions#printorder)|Obtém ou define a opção de ordem de impressão da página da planilha. Isso especifica a ordem que será usada para processar o número de página impresso.|
||[rightMargin](/javascript/api/excel/excel.pagelayoutloadoptions#rightmargin)|Obtém ou define a margem direita da planilha, em pontos, para usar durante a impressão.|
||[topMargin](/javascript/api/excel/excel.pagelayoutloadoptions#topmargin)|Obtém ou define a margem superior da planilha, em pontos, para usar durante a impressão.|
||[zoom](/javascript/api/excel/excel.pagelayoutloadoptions#zoom)|Obtém ou define as opções de zoom de impressão da planilha.|
|[PageLayoutMarginOptions](/javascript/api/excel/excel.pagelayoutmarginoptions)|[bottom](/javascript/api/excel/excel.pagelayoutmarginoptions#bottom)|Representa a margem inferior do layout de página na unidade especificada para usar na impressão.|
||[footer](/javascript/api/excel/excel.pagelayoutmarginoptions#footer)|Representa a margem do rodapé do layout de página na unidade especificada para usar na impressão.|
||[header](/javascript/api/excel/excel.pagelayoutmarginoptions#header)|Representa a margem do cabeçalho do layout de página na unidade especificada para usar na impressão.|
||[left](/javascript/api/excel/excel.pagelayoutmarginoptions#left)|Representa a margem esquerda do layout de página na unidade especificada para usar na impressão.|
||[direita](/javascript/api/excel/excel.pagelayoutmarginoptions#right)|Representa a margem direita do layout de página na unidade especificada para usar na impressão.|
||[top](/javascript/api/excel/excel.pagelayoutmarginoptions#top)|Representa a margem superior do layout de página na unidade especificada para usar na impressão.|
|[PageLayoutUpdateData](/javascript/api/excel/excel.pagelayoutupdatedata)|[blackAndWhite](/javascript/api/excel/excel.pagelayoutupdatedata#blackandwhite)|Obtém ou define a opção de impressão em preto e branco da planilha.|
||[bottomMargin](/javascript/api/excel/excel.pagelayoutupdatedata#bottommargin)|Obtém ou define a margem de página inferior da planilha para impressão em pontos.|
||[centerHorizontally](/javascript/api/excel/excel.pagelayoutupdatedata#centerhorizontally)|Obtém ou define o sinalizador de centralização horizontal da planilha. Esse sinalizador determina se a planilha será centralizada horizontalmente quando for impressa.|
||[centerVertically](/javascript/api/excel/excel.pagelayoutupdatedata#centervertically)|Obtém ou define o sinalizador de centralização vertical da planilha. Esse sinalizador determina se a planilha será centralizada verticalmente quando for impressa.|
||[draftMode](/javascript/api/excel/excel.pagelayoutupdatedata#draftmode)|Obtém ou define a opção de modo de rascunho da planilha. Se for true, a planilha será impressa sem gráficos.|
||[firstPageNumber](/javascript/api/excel/excel.pagelayoutupdatedata#firstpagenumber)|Obtém ou define o primeiro número de página da planilha a ser impressa. O valor null representa a numeração "automática" de páginas.|
||[footerMargin](/javascript/api/excel/excel.pagelayoutupdatedata#footermargin)|Obtém ou define a margem do rodapé da planilha, em pontos, para usar durante a impressão.|
||[headerMargin](/javascript/api/excel/excel.pagelayoutupdatedata#headermargin)|Obtém ou define a margem do cabeçalho da planilha, em pontos, para usar durante a impressão.|
||[headersFooters](/javascript/api/excel/excel.pagelayoutupdatedata#headersfooters)|Configuração de cabeçalho e rodapé da planilha.|
||[leftMargin](/javascript/api/excel/excel.pagelayoutupdatedata#leftmargin)|Obtém ou define a margem esquerda da planilha, em pontos, para usar durante a impressão.|
||[orientation](/javascript/api/excel/excel.pagelayoutupdatedata#orientation)|Obtém ou define a orientação de página da planilha.|
||[paperSize](/javascript/api/excel/excel.pagelayoutupdatedata#papersize)|Obtém ou define o tamanho do papel da página da planilha.|
||[printComments](/javascript/api/excel/excel.pagelayoutupdatedata#printcomments)|Obtém ou define se os comentários da planilha deverão ser exibidos durante a impressão.|
||[printErrors](/javascript/api/excel/excel.pagelayoutupdatedata#printerrors)|Obtém ou define a opção de erros de impressão da planilha.|
||[printGridlines](/javascript/api/excel/excel.pagelayoutupdatedata#printgridlines)|Obtém ou define um sinalizador de linhas de grade de impressão da planilha. Esse sinalizador determina se as linhas de grade serão impressas ou não.|
||[printHeadings](/javascript/api/excel/excel.pagelayoutupdatedata#printheadings)|Obtém ou define um sinalizador de cabeçalhos de impressão da planilha. Esse sinalizador determina se os cabeçalhos serão impressos ou não.|
||[printOrder](/javascript/api/excel/excel.pagelayoutupdatedata#printorder)|Obtém ou define a opção de ordem de impressão da página da planilha. Isso especifica a ordem que será usada para processar o número de página impresso.|
||[rightMargin](/javascript/api/excel/excel.pagelayoutupdatedata#rightmargin)|Obtém ou define a margem direita da planilha, em pontos, para usar durante a impressão.|
||[topMargin](/javascript/api/excel/excel.pagelayoutupdatedata#topmargin)|Obtém ou define a margem superior da planilha, em pontos, para usar durante a impressão.|
||[zoom](/javascript/api/excel/excel.pagelayoutupdatedata#zoom)|Obtém ou define as opções de zoom de impressão da planilha.|
|[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)|[horizontalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#horizontalfittopages)|Número de páginas a ser horizontalmente ajustado. Esse valor pode ser null se o dimensionamento por porcentagem for usado.|
||[scale](/javascript/api/excel/excel.pagelayoutzoomoptions#scale)|O valor do dimensionamento da página de impressão pode estar entre 10 e 400. Esse valor poderá ser null se o ajuste da altura ou largura da página for especificado.|
||[verticalFitToPages](/javascript/api/excel/excel.pagelayoutzoomoptions#verticalfittopages)|Número de páginas a ser verticalmente ajustado. Esse valor pode ser null se o dimensionamento por porcentagem for usado.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[sortByValues (sortBy: "Ascending" \| "Descending", ValuesHierarchy: Excel. DataPivotHierarchy, pivotItemScope?: matriz de cadeia de \| caracteres<PivotItem>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Classifica o Campo dinâmico por valores especificados em um determinado escopo. O escopo define quais valores específicos serão usados na classificação quando|
||[sortByValues(sortBy: Excel.SortBy, valuesHierarchy: Excel.DataPivotHierarchy, pivotItemScope?: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotfield#sortbyvalues-sortby--valueshierarchy--pivotitemscope-)|Classifica o Campo dinâmico por valores especificados em um determinado escopo. O escopo define quais valores específicos serão usados na classificação quando|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[autoFormat](/javascript/api/excel/excel.pivotlayout#autoformat)|Especifica se a formatação será formatada automaticamente quando for atualizada ou quando os campos forem movidos|
||[getDataHierarchy(cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getdatahierarchy-cell-)|Obtém o DataHierarchy que é usado para calcular o valor em um intervalo especificado dentro da Tabela dinâmica.|
||[getPivotItems(axis: "Unknown" \| "Row" \| "Column" \| "Data" \| "Filter", cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Obtém os Itens dinâmicos de um eixo que compõem o valor em um intervalo especificado dentro da Tabela dinâmica.|
||[getPivotItems(axis: Excel.PivotAxis, cell: Range \| string)](/javascript/api/excel/excel.pivotlayout#getpivotitems-axis--cell-)|Obtém os Itens dinâmicos de um eixo que compõem o valor em um intervalo especificado dentro da Tabela dinâmica.|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayout#preserveformatting)|Especifica se a formatação será preservada quando o relatório for atualizado ou recalculado por operações como giro, classificação ou alteração de itens de campo da página.|
||[setAutoSortOnCell (célula: cadeia \| de caracteres de intervalo, sortBy: " \| crescente" "descendente")](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Define a Tabela Dinâmica para classificar automaticamente usando a célula especificada para selecionar automaticamente todos os critérios e contextos necessários. Funciona de maneira idêntica à aplicação de uma autoclassificação da interface do usuário.|
||[setAutoSortOnCell(cell: Range \| string, sortBy: Excel.SortBy)](/javascript/api/excel/excel.pivotlayout#setautosortoncell-cell--sortby-)|Define a Tabela Dinâmica para classificar automaticamente usando a célula especificada para selecionar automaticamente todos os critérios e contextos necessários. Funciona de maneira idêntica à aplicação de uma autoclassificação da interface do usuário.|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[autoFormat](/javascript/api/excel/excel.pivotlayoutdata#autoformat)|Especifica se a formatação será formatada automaticamente quando for atualizada ou quando os campos forem movidos|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutdata#preserveformatting)|Especifica se a formatação será preservada quando o relatório for atualizado ou recalculado por operações como giro, classificação ou alteração de itens de campo da página.|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[autoFormat](/javascript/api/excel/excel.pivotlayoutloadoptions#autoformat)|Especifica se a formatação será formatada automaticamente quando for atualizada ou quando os campos forem movidos|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutloadoptions#preserveformatting)|Especifica se a formatação será preservada quando o relatório for atualizado ou recalculado por operações como giro, classificação ou alteração de itens de campo da página.|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[autoFormat](/javascript/api/excel/excel.pivotlayoutupdatedata#autoformat)|Especifica se a formatação será formatada automaticamente quando for atualizada ou quando os campos forem movidos|
||[preserveFormatting](/javascript/api/excel/excel.pivotlayoutupdatedata#preserveformatting)|Especifica se a formatação será preservada quando o relatório for atualizado ou recalculado por operações como giro, classificação ou alteração de itens de campo da página.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[enableDataValueEditing](/javascript/api/excel/excel.pivottable#enabledatavalueediting)|Especifica se a Tabela Dinâmica permite que os valores no corpo de dados sejam editados pelo usuário.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottable#usecustomsortlists)|Especifica se a tabela dinâmica usa listas personalizadas ao classificar.|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[enableDataValueEditing](/javascript/api/excel/excel.pivottablecollectionloadoptions#enabledatavalueediting)|Para cada ITEM na coleção: especifica se a tabela dinâmica permite que valores no corpo de dados sejam editados pelo usuário.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottablecollectionloadoptions#usecustomsortlists)|Para cada ITEM na coleção: especifica se a tabela dinâmica usa listas personalizadas ao classificar.|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[enableDataValueEditing](/javascript/api/excel/excel.pivottabledata#enabledatavalueediting)|Especifica se a Tabela Dinâmica permite que os valores no corpo de dados sejam editados pelo usuário.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottabledata#usecustomsortlists)|Especifica se a tabela dinâmica usa listas personalizadas ao classificar.|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[enableDataValueEditing](/javascript/api/excel/excel.pivottableloadoptions#enabledatavalueediting)|Especifica se a Tabela Dinâmica permite que os valores no corpo de dados sejam editados pelo usuário.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottableloadoptions#usecustomsortlists)|Especifica se a tabela dinâmica usa listas personalizadas ao classificar.|
|[PivotTableUpdateData](/javascript/api/excel/excel.pivottableupdatedata)|[enableDataValueEditing](/javascript/api/excel/excel.pivottableupdatedata#enabledatavalueediting)|Especifica se a Tabela Dinâmica permite que os valores no corpo de dados sejam editados pelo usuário.|
||[useCustomSortLists](/javascript/api/excel/excel.pivottableupdatedata#usecustomsortlists)|Especifica se a tabela dinâmica usa listas personalizadas ao classificar.|
|[Range](/javascript/api/excel/excel.range)|[autoFill(destinationRange: Range \| string, autoFillType?: "FillDefault" \| "FillCopy" \| "FillSeries" \| "FillFormats" \| "FillValues" \| "FillDays" \| "FillWeekdays" \| "FillMonths" \| "FillYears" \| "LinearTrend" \| "GrowthTrend" \| "FlashFill")](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Os preenchimentos variam do intervalo atual até o intervalo de destino.|
||[autoFill(destinationRange: Range \| string, autoFillType?: Excel.AutoFillType)](/javascript/api/excel/excel.range#autofill-destinationrange--autofilltype-)|Os preenchimentos variam do intervalo atual até o intervalo de destino.|
||[convertDataTypeToText()](/javascript/api/excel/excel.range#convertdatatypetotext--)|Converte o intervalo de células com tipos de dados em texto.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.range#converttolinkeddatatype-serviceid--languageculture-)|Converte as células de intervalo em um tipo de dados vinculado na planilha.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copia a formatação ou dados da célula do intervalo de origem ou de RangeAreas para o intervalo atual.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copia a formatação ou dados da célula do intervalo de origem ou de RangeAreas para o intervalo atual.|
||[find(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#find-text--criteria-)|Localiza certa cadeia de caracteres com base em critérios especificados.|
||[findOrNullObject(text: string, criteria: Excel.SearchCriteria)](/javascript/api/excel/excel.range#findornullobject-text--criteria-)|Localiza certa cadeia de caracteres com base em critérios especificados.|
||[flashFill()](/javascript/api/excel/excel.range#flashfill--)|Faz o preenchimento relâmpago no intervalo atual. O preenchimento relâmpago preenche automaticamente dados quando detecta um padrão. Portanto, o intervalo deve ser de coluna única e ter dados em torno para encontrar o padrão.|
||[getCellProperties(cellPropertiesLoadOptions: CellPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcellproperties-cellpropertiesloadoptions-)|Retorna uma matriz 2D encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada célula.|
||[getColumnProperties(columnPropertiesLoadOptions: ColumnPropertiesLoadOptions)](/javascript/api/excel/excel.range#getcolumnproperties-columnpropertiesloadoptions-)|Retorna uma única matriz dimensional encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada coluna.  Para propriedades que não são consistentes nas células de uma determinada coluna, será retornado null.|
||[getRowProperties(rowPropertiesLoadOptions: RowPropertiesLoadOptions)](/javascript/api/excel/excel.range#getrowproperties-rowpropertiesloadoptions-)|Retorna uma única matriz dimensional encapsulando os dados de fonte, preenchimento, bordas, alinhamento e outras propriedades de cada célula.  Para propriedades que não são consistentes nas células de uma determinada linha, será retornado null.|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas"\|  "SameConditionalFormat" \| "SameDataValidation" \|  "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \|  "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos retangulares, que representa todas as células que correspondem ao tipo e valor especificado.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos retangulares, que representa todas as células que correspondem ao tipo e valor especificado.|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \|"Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \|"LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \|"Text")](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos, que representa todas as células que correspondem ao tipo e valor especificado.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)|Obtém o objeto RangeAreas, compreendendo um ou mais intervalos, que representa todas as células que correspondem ao tipo e valor especificado.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.range#gettables-fullycontained-)|Obtém uma coleção de tabelas com escopo que se sobrepõe ao intervalo.|
||[linkedDataTypeState](/javascript/api/excel/excel.range#linkeddatatypestate)|Representa o estado do tipo de dados de cada célula. Somente leitura.|
||[removeDuplicates(columns: number[], includesHeader: boolean)](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)|Remove valores duplicados do intervalo especificado pelas colunas.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.range#replaceall-text--replacement--criteria-)|Localiza e substitui a cadeia de caracteres fornecida com base nos critérios especificados no intervalo atual.|
||[setCellProperties(cellPropertiesData: SettableCellProperties[][])](/javascript/api/excel/excel.range#setcellproperties-cellpropertiesdata-)|Atualiza o intervalo com base em uma matriz 2D de propriedades da célula, encapsulando itens como fonte, preenchimento, bordas, alinhamento e assim por diante.|
||[setColumnProperties(columnPropertiesData: SettableColumnProperties[])](/javascript/api/excel/excel.range#setcolumnproperties-columnpropertiesdata-)|Atualiza o intervalo com base em uma única matriz dimensional de propriedades da coluna, encapsulando itens como fonte, preenchimento, bordas, alinhamento e assim por diante.|
||[setDirty()](/javascript/api/excel/excel.range#setdirty--)|Define um intervalo a ser recalculado quando o próximo recálculo ocorrer.|
||[setRowProperties(rowPropertiesData: SettableRowProperties[])](/javascript/api/excel/excel.range#setrowproperties-rowpropertiesdata-)|Atualiza o intervalo com base em uma única matriz dimensional de propriedades da linha, encapsulando itens como fonte, preenchimento, bordas, alinhamento e assim por diante.|
|[RangeAreas](/javascript/api/excel/excel.rangeareas)|[calculate()](/javascript/api/excel/excel.rangeareas#calculate--)|Calcula todas as células no RangeAreas.|
||[clear(applyTo?: "All" \| "Formats" \| "Contents" \| "Hyperlinks" \| "RemoveHyperlinks")](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Limpa valores, formato, preenchimento, borda, etc. em cada uma das áreas que compõe este objeto RangeAreas.|
||[clear(applyTo?: Excel.ClearApplyTo)](/javascript/api/excel/excel.rangeareas#clear-applyto-)|Limpa valores, formato, preenchimento, borda, etc. em cada uma das áreas que compõe este objeto RangeAreas.|
||[convertDataTypeToText()](/javascript/api/excel/excel.rangeareas#convertdatatypetotext--)|Converte todas as células de RangeAreas com tipos de dados em texto.|
||[convertToLinkedDataType(serviceID: number, languageCulture: string)](/javascript/api/excel/excel.rangeareas#converttolinkeddatatype-serviceid--languageculture-)|Converte todas as células de RangeAreas em tipos de dados vinculados.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: "All" \| "Formulas" \| "Values" \| "Formats", skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copia a formatação ou dados da célula do intervalo de origem ou de RangeAreas para o RangeAreas atual.|
||[copyFrom(sourceRange: Range \| RangeAreas \| string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean)](/javascript/api/excel/excel.rangeareas#copyfrom-sourcerange--copytype--skipblanks--transpose-)|Copia a formatação ou dados da célula do intervalo de origem ou de RangeAreas para o RangeAreas atual.|
||[getEntireColumn()](/javascript/api/excel/excel.rangeareas#getentirecolumn--)|Retorna um objeto RangeAreas que representa as colunas inteiras dos objetos RangeAreas (por exemplo, se o RangeAreas atual representa as células "B4:E11, H2", ele retorna um RangeAreas que representa as colunas "B:E, H:H").|
||[getEntireRow()](/javascript/api/excel/excel.rangeareas#getentirerow--)|Retorna um objeto RangeAreas que representa as linhas inteiras dos objetos RangeAreas (por exemplo, se o RangeAreas atual representa as células "B4:E11", ele retorna um RangeAreas que representa as linhas "4:11").|
||[getIntersection(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersection-anotherrange-)|Retorna o objeto RangeAreas que representa a interseção dos intervalos fornecidos ou RangeAreas. Se nenhuma interseção for encontrada, um erro ItemNotFound será gerado.|
||[getIntersectionOrNullObject(anotherRange: Range \| RangeAreas \| string)](/javascript/api/excel/excel.rangeareas#getintersectionornullobject-anotherrange-)|Retorna o objeto RangeAreas que representa a interseção dos intervalos fornecidos ou RangeAreas. Se nenhuma interseção for encontrada, um objeto null será retornado.|
||[getOffsetRangeAreas(rowOffset: number, columnOffset: number)](/javascript/api/excel/excel.rangeareas#getoffsetrangeareas-rowoffset--columnoffset-)|Retorna um objeto RangeAreas que é deslocado pelo deslocamento de linha e coluna específico. A dimensão do RangeAreas retornado corresponderá ao objeto original. Se o RangeAreas resultante for imposto para fora dos limites da grade da planilha, o sistema gerará um erro.|
||[getSpecialCells(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \| "Constants" \| "Formulas"\| "SameConditionalFormat" \| "SameDataValidation" \|  "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \|  "ErrorsLogicalNumber" \| "ErrorsLogicalText" \|  "ErrorsNumberText" \| "Logical" \| "LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \| "Text")](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Retorna um objeto RangeAreas que representa todas as células que correspondem ao tipo e valor especificados. Gera um erro se nenhuma célula especial que corresponda aos critérios for encontrada.|
||[getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcells-celltype--cellvaluetype-)|Retorna um objeto RangeAreas que representa todas as células que correspondem ao tipo e valor especificados. Gera um erro se nenhuma célula especial que corresponda aos critérios for encontrada.|
||[getSpecialCellsOrNullObject(cellType: "ConditionalFormats" \| "DataValidations" \| "Blanks" \|"Constants" \| "Formulas" \| "SameConditionalFormat" \| "SameDataValidation" \| "Visible", cellValueType?: "All" \| "Errors" \| "ErrorsLogical" \| "ErrorsNumbers" \| "ErrorsText" \| "ErrorsLogicalNumber" \| "ErrorsLogicalText" \| "ErrorsNumberText" \| "Logical" \|"LogicalNumbers" \| "LogicalText" \| "LogicalNumbersText" \| "Numbers" \| "NumbersText" \|"Text")](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Retorna um objeto RangeAreas que representa todas as células que correspondem ao tipo e valor especificados. Retorna um objeto null se nenhuma célula especial que corresponda ao critério for encontrada.|
||[getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType)](/javascript/api/excel/excel.rangeareas#getspecialcellsornullobject-celltype--cellvaluetype-)|Retorna um objeto RangeAreas que representa todas as células que correspondem ao tipo e valor especificados. Retorna um objeto null se nenhuma célula especial que corresponda ao critério for encontrada.|
||[getTables(fullyContained?: boolean)](/javascript/api/excel/excel.rangeareas#gettables-fullycontained-)|Retorna uma coleção de tabelas com escopo que se sobrepõe a qualquer intervalo neste objeto RangeAreas.|
||[getUsedRangeAreas(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareas-valuesonly-)|Retorna o RangeAreas usado que compreende todas as áreas utilizadas de intervalos retangulares individuais no objeto RangeAreas.|
||[getUsedRangeAreasOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.rangeareas#getusedrangeareasornullobject-valuesonly-)|Retorna o RangeAreas usado que compreende todas as áreas utilizadas de intervalos retangulares individuais no objeto RangeAreas.|
||[address](/javascript/api/excel/excel.rangeareas#address)|Retorna a referência RageAreas no estilo A1. O valor do endereço conterá o nome da planilha para cada bloco retangular de células (por exemplo, "Sheet1!A1:B4, Sheet1!D1:D4"). Somente leitura.|
||[addressLocal](/javascript/api/excel/excel.rangeareas#addresslocal)|Retorna a referência RageAreas na localidade do usuário.  Somente leitura.|
||[areaCount](/javascript/api/excel/excel.rangeareas#areacount)|Retorna o número de intervalos retangulares que compõem este objeto RangeAreas.|
||[areas](/javascript/api/excel/excel.rangeareas#areas)|Retorna uma coleção de intervalos retangulares que compõem este objeto RangeAreas.|
||[cellCount](/javascript/api/excel/excel.rangeareas#cellcount)|Retorna o número de células no objeto RangeAreas somando as contagens de células de todos os intervalos retangulares individuais. Retornará -1 se a contagem de células exceder 2^31-1 (2.147.483.647). Somente leitura.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareas#conditionalformats)|Retorna uma coleção de ConditionalFormats que se cruza com qualquer célula nesse objeto RangeAreas. Somente leitura.|
||[dataValidation](/javascript/api/excel/excel.rangeareas#datavalidation)|Retorna um objeto dataValidation para todos os intervalos no RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareas#format)|Retorna um objeto rangeFormat encapsulando a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades para todos os intervalos no objeto RangeAreas. Somente leitura.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareas#isentirecolumn)|Indica se todos os intervalos neste objeto RangeAreas representam colunas inteiras (por exemplo, "A:C, Q:Z"). Somente leitura.|
||[isEntireRow](/javascript/api/excel/excel.rangeareas#isentirerow)|Indica se todos os intervalos neste objeto RangeAreas representam linhas inteiras (por exemplo, "1:3, 5:7"). Somente leitura.|
||[worksheet](/javascript/api/excel/excel.rangeareas#worksheet)|Retorna a planilha para o RangeAreas atual. Somente leitura.|
||[Set (Propriedades: Excel. RangeAreas)](/javascript/api/excel/excel.rangeareas#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. RangeAreasUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.rangeareas#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[setDirty()](/javascript/api/excel/excel.rangeareas#setdirty--)|Define o RangeAreas que será recalculado quando o próximo recálculo ocorrer.|
||[style](/javascript/api/excel/excel.rangeareas#style)|Representa o estilo de todos os intervalos nesse objeto RangeAreas.|
||[track()](/javascript/api/excel/excel.rangeareas#track--)|Acompanha o objeto para ajuste automático com base nas alterações adjacentes no documento. Essa chamada é uma abreviação de context.trackedObjects.add(thisObject). Se você estiver usando esse objeto em chamadas ".sync" e fora da execução sequencial de um lote ".run" e receber um erro "InvalidObjectPath" ao definir uma propriedade ou invocar um método no objeto, era necessário ter adicionado o objeto à coleção de objetos rastreados quando o objeto foi criado pela primeira vez.|
||[untrack()](/javascript/api/excel/excel.rangeareas#untrack--)|Libere a memória associada a este objeto, se ele já tiver sido rastreado anteriormente. Essa chamada é uma abreviação de context.trackedObjects.remove(thisObject). Ter muitos objetos rastreados desacelera o aplicativo host, por isso, lembre-se de liberar todos os objetos adicionados após usá-los. Você precisa chamar "context.sync()" antes da liberação da memória entrar em vigor.|
|[RangeAreasData](/javascript/api/excel/excel.rangeareasdata)|[address](/javascript/api/excel/excel.rangeareasdata#address)|Retorna a referência RageAreas no estilo A1. O valor do endereço conterá o nome da planilha para cada bloco retangular de células (por exemplo, "Sheet1!A1:B4, Sheet1!D1:D4"). Somente leitura.|
||[addressLocal](/javascript/api/excel/excel.rangeareasdata#addresslocal)|Retorna a referência RageAreas na localidade do usuário.  Somente leitura.|
||[areaCount](/javascript/api/excel/excel.rangeareasdata#areacount)|Retorna o número de intervalos retangulares que compõem este objeto RangeAreas.|
||[areas](/javascript/api/excel/excel.rangeareasdata#areas)|Retorna uma coleção de intervalos retangulares que compõem este objeto RangeAreas.|
||[cellCount](/javascript/api/excel/excel.rangeareasdata#cellcount)|Retorna o número de células no objeto RangeAreas somando as contagens de células de todos os intervalos retangulares individuais. Retornará -1 se a contagem de células exceder 2^31-1 (2.147.483.647). Somente leitura.|
||[conditionalFormats](/javascript/api/excel/excel.rangeareasdata#conditionalformats)|Retorna uma coleção de ConditionalFormats que se cruza com qualquer célula nesse objeto RangeAreas. Somente leitura.|
||[dataValidation](/javascript/api/excel/excel.rangeareasdata#datavalidation)|Retorna um objeto dataValidation para todos os intervalos no RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareasdata#format)|Retorna um objeto rangeFormat encapsulando a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades para todos os intervalos no objeto RangeAreas. Somente leitura.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareasdata#isentirecolumn)|Indica se todos os intervalos neste objeto RangeAreas representam colunas inteiras (por exemplo, "A:C, Q:Z"). Somente leitura.|
||[isEntireRow](/javascript/api/excel/excel.rangeareasdata#isentirerow)|Indica se todos os intervalos neste objeto RangeAreas representam linhas inteiras (por exemplo, "1:3, 5:7"). Somente leitura.|
||[style](/javascript/api/excel/excel.rangeareasdata#style)|Representa o estilo de todos os intervalos nesse objeto RangeAreas.|
|[RangeAreasLoadOptions](/javascript/api/excel/excel.rangeareasloadoptions)|[$all](/javascript/api/excel/excel.rangeareasloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangeareasloadoptions#address)|Retorna a referência RageAreas no estilo A1. O valor do endereço conterá o nome da planilha para cada bloco retangular de células (por exemplo, "Sheet1!A1:B4, Sheet1!D1:D4"). Somente leitura.|
||[addressLocal](/javascript/api/excel/excel.rangeareasloadoptions#addresslocal)|Retorna a referência RageAreas na localidade do usuário.  Somente leitura.|
||[areaCount](/javascript/api/excel/excel.rangeareasloadoptions#areacount)|Retorna o número de intervalos retangulares que compõem este objeto RangeAreas.|
||[cellCount](/javascript/api/excel/excel.rangeareasloadoptions#cellcount)|Retorna o número de células no objeto RangeAreas somando as contagens de células de todos os intervalos retangulares individuais. Retornará -1 se a contagem de células exceder 2^31-1 (2.147.483.647). Somente leitura.|
||[dataValidation](/javascript/api/excel/excel.rangeareasloadoptions#datavalidation)|Retorna um objeto dataValidation para todos os intervalos no RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareasloadoptions#format)|Retorna um objeto rangeFormat encapsulando a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades para todos os intervalos no objeto RangeAreas.|
||[isEntireColumn](/javascript/api/excel/excel.rangeareasloadoptions#isentirecolumn)|Indica se todos os intervalos neste objeto RangeAreas representam colunas inteiras (por exemplo, "A:C, Q:Z"). Somente leitura.|
||[isEntireRow](/javascript/api/excel/excel.rangeareasloadoptions#isentirerow)|Indica se todos os intervalos neste objeto RangeAreas representam linhas inteiras (por exemplo, "1:3, 5:7"). Somente leitura.|
||[style](/javascript/api/excel/excel.rangeareasloadoptions#style)|Representa o estilo de todos os intervalos nesse objeto RangeAreas.|
||[worksheet](/javascript/api/excel/excel.rangeareasloadoptions#worksheet)|Retorna a planilha para o RangeAreas atual.|
|[RangeAreasUpdateData](/javascript/api/excel/excel.rangeareasupdatedata)|[dataValidation](/javascript/api/excel/excel.rangeareasupdatedata#datavalidation)|Retorna um objeto dataValidation para todos os intervalos no RangeAreas.|
||[format](/javascript/api/excel/excel.rangeareasupdatedata#format)|Retorna um objeto rangeFormat encapsulando a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades para todos os intervalos no objeto RangeAreas.|
||[style](/javascript/api/excel/excel.rangeareasupdatedata#style)|Representa o estilo de todos os intervalos nesse objeto RangeAreas.|
|[RangeBorder](/javascript/api/excel/excel.rangeborder)|[tintAndShade](/javascript/api/excel/excel.rangeborder#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece a cor para a Borda do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeBorderCollection](/javascript/api/excel/excel.rangebordercollection)|[tintAndShade](/javascript/api/excel/excel.rangebordercollection#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece a cor para as Bordas do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeBorderCollectionLoadOptions](/javascript/api/excel/excel.rangebordercollectionloadoptions)|[tintAndShade](/javascript/api/excel/excel.rangebordercollectionloadoptions#tintandshade)|Para cada ITEM na coleção: Retorna ou define um duplo que clareia ou escurece uma cor para a borda do intervalo, o valor é entre-1 (mais escuro) e 1 (mais brilhante), com 0 para a cor original.|
|[RangeBorderCollectionUpdateData](/javascript/api/excel/excel.rangebordercollectionupdatedata)|[tintAndShade](/javascript/api/excel/excel.rangebordercollectionupdatedata#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece a cor para as Bordas do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeBorderData](/javascript/api/excel/excel.rangeborderdata)|[tintAndShade](/javascript/api/excel/excel.rangeborderdata#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece a cor para a Borda do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeBorderLoadOptions](/javascript/api/excel/excel.rangeborderloadoptions)|[tintAndShade](/javascript/api/excel/excel.rangeborderloadoptions#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece a cor para a Borda do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeBorderUpdateData](/javascript/api/excel/excel.rangeborderupdatedata)|[tintAndShade](/javascript/api/excel/excel.rangeborderupdatedata#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece a cor para a Borda do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeCollection](/javascript/api/excel/excel.rangecollection)|[getCount()](/javascript/api/excel/excel.rangecollection#getcount--)|Retorna o número de intervalos no RangeCollection.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangecollection#getitemat-index-)|Retorna o objeto range com base em sua posição no RangeCollection.|
||[items](/javascript/api/excel/excel.rangecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[RangeCollectionLoadOptions](/javascript/api/excel/excel.rangecollectionloadoptions)|[$all](/javascript/api/excel/excel.rangecollectionloadoptions#$all)||
||[address](/javascript/api/excel/excel.rangecollectionloadoptions#address)|Para cada ITEM na coleção: representa a referência de intervalo no estilo a1. O valor de endereço conterá a referência de planilha (por exemplo, "Planilha1! A1: B4 "). Somente leitura.|
||[addressLocal](/javascript/api/excel/excel.rangecollectionloadoptions#addresslocal)|Para cada ITEM na coleção: representa a referência de intervalo para o intervalo especificado no idioma do usuário. Somente leitura.|
||[cellCount](/javascript/api/excel/excel.rangecollectionloadoptions#cellcount)|Para cada ITEM na coleção: número de células no intervalo. Essa API retornará -1 se a contagem de células exceder 2^31-1 (2.147.483.647). Somente leitura.|
||[columnCount](/javascript/api/excel/excel.rangecollectionloadoptions#columncount)|Para cada ITEM na coleção: representa o número total de colunas no intervalo. Somente leitura.|
||[columnHidden](/javascript/api/excel/excel.rangecollectionloadoptions#columnhidden)|Para cada ITEM na coleção: representa se todas as colunas do intervalo atual estão ocultas.|
||[columnIndex](/javascript/api/excel/excel.rangecollectionloadoptions#columnindex)|Para cada ITEM na coleção: representa o número de coluna da primeira célula do intervalo. Indexados com zero. Somente leitura.|
||[dataValidation](/javascript/api/excel/excel.rangecollectionloadoptions#datavalidation)|Para cada ITEM na coleção: retorna um objeto de validação de dados.|
||[format](/javascript/api/excel/excel.rangecollectionloadoptions#format)|Para cada ITEM na coleção: retorna um objeto Format, encapsulando a fonte, o preenchimento, as bordas, o alinhamento e outras propriedades do intervalo.|
||[fórmulas](/javascript/api/excel/excel.rangecollectionloadoptions#formulas)|Para cada ITEM na coleção: representa a fórmula em notação de estilo a1.|
||[formulasLocal](/javascript/api/excel/excel.rangecollectionloadoptions#formulaslocal)|Para cada ITEM na coleção: representa a fórmula em notação de estilo a1, no idioma do usuário e na localidade de formatação de números.  Por exemplo, a fórmula "=SUM(A1, 1.5)" em inglês seria "=SOMA(A1; 1,5)" em português.|
||[formulasR1C1](/javascript/api/excel/excel.rangecollectionloadoptions#formulasr1c1)|Para cada ITEM na coleção: representa a fórmula em notação de estilo L1C1.|
||[hidden](/javascript/api/excel/excel.rangecollectionloadoptions#hidden)|Para cada ITEM na coleção: representa se todas as células do intervalo atual estão ocultas. Somente leitura.|
||[hiperlink](/javascript/api/excel/excel.rangecollectionloadoptions#hyperlink)|Para cada ITEM na coleção: representa o hiperlink para o intervalo atual.|
||[isEntireColumn](/javascript/api/excel/excel.rangecollectionloadoptions#isentirecolumn)|Para cada ITEM na coleção: representa se o intervalo atual é uma coluna inteira. Somente leitura.|
||[isEntireRow](/javascript/api/excel/excel.rangecollectionloadoptions#isentirerow)|Para cada ITEM na coleção: representa se o intervalo atual é uma linha inteira. Somente leitura.|
||[linkedDataTypeState](/javascript/api/excel/excel.rangecollectionloadoptions#linkeddatatypestate)|Para cada ITEM na coleção: representa o estado do tipo de dados de cada célula. Somente leitura.|
||[numberFormat](/javascript/api/excel/excel.rangecollectionloadoptions#numberformat)|Para cada ITEM na coleção: representa o código de formato de número do Excel para o intervalo especificado.|
||[numberFormatLocal](/javascript/api/excel/excel.rangecollectionloadoptions#numberformatlocal)|Para cada ITEM na coleção: representa o código de formato de número do Excel para o intervalo determinado como uma cadeia de caracteres no idioma do usuário.|
||[Validação](/javascript/api/excel/excel.rangecollectionloadoptions#rowcount)|Para cada ITEM na coleção: retorna o número total de linhas no intervalo. Somente leitura.|
||[rowHidden](/javascript/api/excel/excel.rangecollectionloadoptions#rowhidden)|Para cada ITEM na coleção: representa se todas as linhas do intervalo atual estão ocultas.|
||[rowIndex](/javascript/api/excel/excel.rangecollectionloadoptions#rowindex)|Para cada ITEM na coleção: retorna o número de linha da primeira célula do intervalo. Indexados com zero. Somente leitura.|
||[style](/javascript/api/excel/excel.rangecollectionloadoptions#style)|Para cada ITEM na coleção: representa o estilo do intervalo atual.|
||[text](/javascript/api/excel/excel.rangecollectionloadoptions#text)|Para cada ITEM na coleção: valores de texto do intervalo especificado. O valor de texto não depende da largura da célula. A substituição pelo sinal #, que ocorre na interface de usuário do Excel, não afeta o valor de texto retornado pela API. Somente leitura.|
||[valueTypes](/javascript/api/excel/excel.rangecollectionloadoptions#valuetypes)|Para cada ITEM na coleção: representa o tipo de dados de cada célula. Somente leitura.|
||[values](/javascript/api/excel/excel.rangecollectionloadoptions#values)|Para cada ITEM na coleção: representa os valores brutos do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|
||[worksheet](/javascript/api/excel/excel.rangecollectionloadoptions#worksheet)|Para cada ITEM na coleção: a planilha que contém o intervalo atual.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[linkedDataTypeState](/javascript/api/excel/excel.rangedata#linkeddatatypestate)|Representa o estado do tipo de dados de cada célula. Somente leitura.|
|[RangeFill](/javascript/api/excel/excel.rangefill)|[padrão](/javascript/api/excel/excel.rangefill#pattern)|Obtém ou define o padrão de um intervalo. Para saber detalhes, confira Excel.FillPattern. LinearGradient e RectangularGradient não são compatíveis.|
||[patternColor](/javascript/api/excel/excel.rangefill#patterncolor)|Define o código de cor HTML que representa a cor do padrão Range, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefill#patterntintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor padrão para o Preenchimento do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
||[tintAndShade](/javascript/api/excel/excel.rangefill#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor para o Preenchimento do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeFillData](/javascript/api/excel/excel.rangefilldata)|[padrão](/javascript/api/excel/excel.rangefilldata#pattern)|Obtém ou define o padrão de um intervalo. Para saber detalhes, confira Excel.FillPattern. LinearGradient e RectangularGradient não são compatíveis.|
||[patternColor](/javascript/api/excel/excel.rangefilldata#patterncolor)|Define o código de cor HTML que representa a cor do padrão Range, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefilldata#patterntintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor padrão para o Preenchimento do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
||[tintAndShade](/javascript/api/excel/excel.rangefilldata#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor para o Preenchimento do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeFillLoadOptions](/javascript/api/excel/excel.rangefillloadoptions)|[padrão](/javascript/api/excel/excel.rangefillloadoptions#pattern)|Obtém ou define o padrão de um intervalo. Para saber detalhes, confira Excel.FillPattern. LinearGradient e RectangularGradient não são compatíveis.|
||[patternColor](/javascript/api/excel/excel.rangefillloadoptions#patterncolor)|Define o código de cor HTML que representa a cor do padrão Range, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefillloadoptions#patterntintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor padrão para o Preenchimento do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
||[tintAndShade](/javascript/api/excel/excel.rangefillloadoptions#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor para o Preenchimento do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeFillUpdateData](/javascript/api/excel/excel.rangefillupdatedata)|[padrão](/javascript/api/excel/excel.rangefillupdatedata#pattern)|Obtém ou define o padrão de um intervalo. Para saber detalhes, confira Excel.FillPattern. LinearGradient e RectangularGradient não são compatíveis.|
||[patternColor](/javascript/api/excel/excel.rangefillupdatedata#patterncolor)|Define o código de cor HTML que representa a cor do padrão Range, do formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja").|
||[patternTintAndShade](/javascript/api/excel/excel.rangefillupdatedata#patterntintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor padrão para o Preenchimento do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
||[tintAndShade](/javascript/api/excel/excel.rangefillupdatedata#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor para o Preenchimento do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeFont](/javascript/api/excel/excel.rangefont)|[strikethrough](/javascript/api/excel/excel.rangefont#strikethrough)|Representa o status da fonte em tachado. Um valor nulo indica que todo o intervalo não tem configuração de tachado uniforme.|
||[subscript](/javascript/api/excel/excel.rangefont#subscript)|Representa o status da fonte em subscrito.|
||[superscript](/javascript/api/excel/excel.rangefont#superscript)|Representa o status da fonte em sobrescrito.|
||[tintAndShade](/javascript/api/excel/excel.rangefont#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor para a Fonte do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeFontData](/javascript/api/excel/excel.rangefontdata)|[strikethrough](/javascript/api/excel/excel.rangefontdata#strikethrough)|Representa o status da fonte em tachado. Um valor nulo indica que todo o intervalo não tem configuração de tachado uniforme.|
||[subscript](/javascript/api/excel/excel.rangefontdata#subscript)|Representa o status da fonte em subscrito.|
||[superscript](/javascript/api/excel/excel.rangefontdata#superscript)|Representa o status da fonte em sobrescrito.|
||[tintAndShade](/javascript/api/excel/excel.rangefontdata#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor para a Fonte do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeFontLoadOptions](/javascript/api/excel/excel.rangefontloadoptions)|[strikethrough](/javascript/api/excel/excel.rangefontloadoptions#strikethrough)|Representa o status da fonte em tachado. Um valor nulo indica que todo o intervalo não tem configuração de tachado uniforme.|
||[subscript](/javascript/api/excel/excel.rangefontloadoptions#subscript)|Representa o status da fonte em subscrito.|
||[superscript](/javascript/api/excel/excel.rangefontloadoptions#superscript)|Representa o status da fonte em sobrescrito.|
||[tintAndShade](/javascript/api/excel/excel.rangefontloadoptions#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor para a Fonte do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeFontUpdateData](/javascript/api/excel/excel.rangefontupdatedata)|[strikethrough](/javascript/api/excel/excel.rangefontupdatedata#strikethrough)|Representa o status da fonte em tachado. Um valor nulo indica que todo o intervalo não tem configuração de tachado uniforme.|
||[subscript](/javascript/api/excel/excel.rangefontupdatedata#subscript)|Representa o status da fonte em subscrito.|
||[superscript](/javascript/api/excel/excel.rangefontupdatedata#superscript)|Representa o status da fonte em sobrescrito.|
||[tintAndShade](/javascript/api/excel/excel.rangefontupdatedata#tintandshade)|Retorna ou define um valor em dobro que clareia ou escurece uma cor para a Fonte do intervalo, o valor fica entre -1 (mais escuro) e 1 (mais claro), sendo 0 a cor original.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[autoIndent](/javascript/api/excel/excel.rangeformat#autoindent)|Indica se o texto é automaticamente recuado quando o alinhamento de texto é definido como distribuição igual.|
||[indentLevel](/javascript/api/excel/excel.rangeformat#indentlevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo.|
||[readingOrder](/javascript/api/excel/excel.rangeformat#readingorder)|A ordem de leitura para o intervalo.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformat#shrinktofit)|Indica se o texto é automaticamente reduzido para caber na largura da coluna disponível.|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[autoIndent](/javascript/api/excel/excel.rangeformatdata#autoindent)|Indica se o texto é automaticamente recuado quando o alinhamento de texto é definido como distribuição igual.|
||[indentLevel](/javascript/api/excel/excel.rangeformatdata#indentlevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo.|
||[readingOrder](/javascript/api/excel/excel.rangeformatdata#readingorder)|A ordem de leitura para o intervalo.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatdata#shrinktofit)|Indica se o texto é automaticamente reduzido para caber na largura da coluna disponível.|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[autoIndent](/javascript/api/excel/excel.rangeformatloadoptions#autoindent)|Indica se o texto é automaticamente recuado quando o alinhamento de texto é definido como distribuição igual.|
||[indentLevel](/javascript/api/excel/excel.rangeformatloadoptions#indentlevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo.|
||[readingOrder](/javascript/api/excel/excel.rangeformatloadoptions#readingorder)|A ordem de leitura para o intervalo.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatloadoptions#shrinktofit)|Indica se o texto é automaticamente reduzido para caber na largura da coluna disponível.|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[autoIndent](/javascript/api/excel/excel.rangeformatupdatedata#autoindent)|Indica se o texto é automaticamente recuado quando o alinhamento de texto é definido como distribuição igual.|
||[indentLevel](/javascript/api/excel/excel.rangeformatupdatedata#indentlevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo.|
||[readingOrder](/javascript/api/excel/excel.rangeformatupdatedata#readingorder)|A ordem de leitura para o intervalo.|
||[shrinkToFit](/javascript/api/excel/excel.rangeformatupdatedata#shrinktofit)|Indica se o texto é automaticamente reduzido para caber na largura da coluna disponível.|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[linkedDataTypeState](/javascript/api/excel/excel.rangeloadoptions#linkeddatatypestate)|Representa o estado do tipo de dados de cada célula. Somente leitura.|
|[RemoveDuplicatesResult](/javascript/api/excel/excel.removeduplicatesresult)|[removed](/javascript/api/excel/excel.removeduplicatesresult#removed)|Número de linhas duplicadas removidas pela operação.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresult#uniqueremaining)|Número de linhas restantes exclusivas presentes no intervalo resultante.|
|[RemoveDuplicatesResultData](/javascript/api/excel/excel.removeduplicatesresultdata)|[removed](/javascript/api/excel/excel.removeduplicatesresultdata#removed)|Número de linhas duplicadas removidas pela operação.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresultdata#uniqueremaining)|Número de linhas restantes exclusivas presentes no intervalo resultante.|
|[RemoveDuplicatesResultLoadOptions](/javascript/api/excel/excel.removeduplicatesresultloadoptions)|[$all](/javascript/api/excel/excel.removeduplicatesresultloadoptions#$all)||
||[removed](/javascript/api/excel/excel.removeduplicatesresultloadoptions#removed)|Número de linhas duplicadas removidas pela operação.|
||[uniqueRemaining](/javascript/api/excel/excel.removeduplicatesresultloadoptions#uniqueremaining)|Número de linhas restantes exclusivas presentes no intervalo resultante.|
|[ReplaceCriteria](/javascript/api/excel/excel.replacecriteria)|[completeMatch](/javascript/api/excel/excel.replacecriteria#completematch)|Especifica se a correspondência deve ser completa ou parcial. O padrão é false (parcial).|
||[matchCase](/javascript/api/excel/excel.replacecriteria#matchcase)|Especifica se a correspondência diferencia maiúsculas de minúsculas. O padrão é false (não diferencia maiúsculas de minúsculas).|
|[RowProperties](/javascript/api/excel/excel.rowproperties)|[address](/javascript/api/excel/excel.rowproperties#address)|Representa a propriedade `address`.|
||[addressLocal](/javascript/api/excel/excel.rowproperties#addresslocal)|Representa a propriedade `addressLocal`.|
||[rowIndex](/javascript/api/excel/excel.rowproperties#rowindex)|Representa a propriedade `rowIndex`.|
|[RowPropertiesLoadOptions](/javascript/api/excel/excel.rowpropertiesloadoptions)|[formato: Excel. CellPropertiesFormatLoadOptions & {
            AlturaDaLinha?] (formato/JavaScript/API/Excel/Excel.rowpropertiesloadoptions #)|Especifica se a `format` propriedade deve ser carregada.|
||[rowHeight](/javascript/api/excel/excel.rowpropertiesloadoptions#rowheight)||
||[rowHidden](/javascript/api/excel/excel.rowpropertiesloadoptions#rowhidden)|Especifica se a `rowHidden` propriedade deve ser carregada.|
||[rowIndex](/javascript/api/excel/excel.rowpropertiesloadoptions#rowindex)|Especifica se a `rowIndex` propriedade deve ser carregada.|
|[SearchCriteria](/javascript/api/excel/excel.searchcriteria)|[completeMatch](/javascript/api/excel/excel.searchcriteria#completematch)|Especifica se a correspondência deve ser completa ou parcial. Uma correspondência completa corresponde a todo o conteúdo da célula. O padrão é false (parcial).|
||[matchCase](/javascript/api/excel/excel.searchcriteria#matchcase)|Especifica se a correspondência diferencia maiúsculas de minúsculas. O padrão é false (não diferencia maiúsculas de minúsculas).|
||[searchDirection](/javascript/api/excel/excel.searchcriteria#searchdirection)|Especifica a direção da pesquisa. O padrão é para frente. Confira Excel.SearchDirection.|
|[SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)|[format](/javascript/api/excel/excel.settablecellproperties#format)|Representa a propriedade `format`.|
||[hiperlink](/javascript/api/excel/excel.settablecellproperties#hyperlink)|Representa a propriedade `hyperlink`.|
||[style](/javascript/api/excel/excel.settablecellproperties#style)|Representa a propriedade `style`.|
|[SettableColumnProperties](/javascript/api/excel/excel.settablecolumnproperties)|[columnHidden](/javascript/api/excel/excel.settablecolumnproperties#columnhidden)|Representa a propriedade `columnHidden`.|
||[columnWidth](/javascript/api/excel/excel.settablecolumnproperties#columnwidth)||
||[formato: Excel. CellPropertiesFormat & {
            columnWidth?] (formato/JavaScript/API/Excel/Excel.settablecolumnproperties #)|Representa a propriedade `format`.|
|[SettableRowProperties](/javascript/api/excel/excel.settablerowproperties)|[formato: Excel. CellPropertiesFormat & {
            AlturaDaLinha?] (formato/JavaScript/API/Excel/Excel.settablerowproperties #)|Representa a propriedade `format`.|
||[rowHeight](/javascript/api/excel/excel.settablerowproperties#rowheight)||
||[rowHidden](/javascript/api/excel/excel.settablerowproperties#rowhidden)|Representa a propriedade `rowHidden`.|
|[Shape](/javascript/api/excel/excel.shape)|[altTextDescription](/javascript/api/excel/excel.shape#alttextdescription)|Retorna ou define o texto da descrição alternativa de um objeto de forma.|
||[altTextTitle](/javascript/api/excel/excel.shape#alttexttitle)|Retorna ou define o texto do título alternativo de um objeto de forma.|
||[delete()](/javascript/api/excel/excel.shape#delete--)|Remove a forma da planilha.|
||[geometricShapeType](/javascript/api/excel/excel.shape#geometricshapetype)|Representa o tipo de forma geométricas da forma geométrica. Para saber detalhes, confira Excel.GeometricShapeType. Retorna nulo se o tipo de forma não for "GeometricShape".|
||[getAsImage(format: "UNKNOWN" \| "BMP" \| "JPEG" \| "GIF" \| "PNG" \| "SVG")](/javascript/api/excel/excel.shape#getasimage-format-)|Converte a forma em uma imagem e retorna a imagem como uma cadeia de caracteres de base 64. O DPI é 96. Os formatos com suporte apenas são `Excel.PictureFormat.BMP`, `Excel.PictureFormat.PNG`, `Excel.PictureFormat.JPEG`, e `Excel.PictureFormat.GIF`.|
||[getAsImage(format: Excel.PictureFormat)](/javascript/api/excel/excel.shape#getasimage-format-)|Converte a forma em uma imagem e retorna a imagem como uma cadeia de caracteres de base 64. O DPI é 96. Os formatos com suporte apenas são `Excel.PictureFormat.BMP`, `Excel.PictureFormat.PNG`, `Excel.PictureFormat.JPEG`, e `Excel.PictureFormat.GIF`.|
||[height](/javascript/api/excel/excel.shape#height)|Representa a altura, em pontos, da forma.|
||[incrementLeft(increment: number)](/javascript/api/excel/excel.shape#incrementleft-increment-)|Move a forma horizontalmente pelo número especificado de pontos.|
||[incrementRotation(increment: number)](/javascript/api/excel/excel.shape#incrementrotation-increment-)|O formato é girado em sentido horário ao redor do eixo z pelo número especificado de graus.|
||[incrementTop(increment: number)](/javascript/api/excel/excel.shape#incrementtop-increment-)|Move a forma verticalmente pelo número especificado de pontos.|
||[left](/javascript/api/excel/excel.shape#left)|A distância, em pontos, da lateral esquerda da forma do lado  esquerdo da planilha.|
||[lockAspectRatio](/javascript/api/excel/excel.shape#lockaspectratio)|Especifica se a taxa de proporção dessa forma está bloqueada ou não.|
||[name](/javascript/api/excel/excel.shape#name)|Representa o nome da forma.|
||[connectionSiteCount](/javascript/api/excel/excel.shape#connectionsitecount)|Retorna o número de locais de conexão nessa forma. Somente leitura.|
||[fill](/javascript/api/excel/excel.shape#fill)|Retorna a formatação de preenchimento dessa forma. Somente leitura.|
||[geometricShape](/javascript/api/excel/excel.shape#geometricshape)|Retorna a forma geométrica associada à forma. Um erro será lançado, se o tipo de forma não for "GeometricShape".|
||[group](/javascript/api/excel/excel.shape#group)|Retorna o grupo de forma associado à forma. Um erro será lançado, se o tipo de forma não for "GroupShape".|
||[id](/javascript/api/excel/excel.shape#id)|Representa o identificador de forma. Somente leitura.|
||[image](/javascript/api/excel/excel.shape#image)|Retorna a imagem associada à forma. Um erro será lançado, se o tipo de forma não for "Imagem".|
||[level](/javascript/api/excel/excel.shape#level)|Representa o nível da forma especificada. Por exemplo, um nível de 0 significa que a forma não faz parte de nenhum grupo, um nível de 1 significa que a forma é parte de um grupo de nível superior e um nível 2 significa que a forma faz parte de um subgrupo do nível superior.|
||[line](/javascript/api/excel/excel.shape#line)|Retorna a linha associada à forma. Um erro será lançado, se o tipo de forma não for "Linha".|
||[lineFormat](/javascript/api/excel/excel.shape#lineformat)|Retorna a formatação de linha do objeto de forma. Somente leitura.|
||[onActivated](/javascript/api/excel/excel.shape#onactivated)|Ocorre quando a forma é ativada.|
||[onDeactivated](/javascript/api/excel/excel.shape#ondeactivated)|Ocorre quando a forma é desativada.|
||[parentGroup](/javascript/api/excel/excel.shape#parentgroup)|Representa o grupo pai dessa forma.|
||[textFrame](/javascript/api/excel/excel.shape#textframe)|Retorna o objeto text frame de uma forma. Somente leitura.|
||[tipo](/javascript/api/excel/excel.shape#type)|Retorna o tipo dessa forma. Para saber detalhes, confira Excel.ShapeType. Somente leitura.|
||[zOrderPosition](/javascript/api/excel/excel.shape#zorderposition)|Retorna a posição da forma especificada na ordem z, com 0 representando a parte inferior da pilha do pedido. Somente leitura.|
||[rotation](/javascript/api/excel/excel.shape#rotation)|Representa a rotação, em graus, da forma.|
||[scaleHeight(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Dimensiona a altura da forma por um fator especificado. Para imagens, é possível indicar se você deseja dimensionar a forma em relação ao tamanho original ou ao tamanho atual. As formas que não são figuras serão sempre dimensionadas em relação à sua altura atual.|
||[scaleHeight(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scaleheight-scalefactor--scaletype--scalefrom-)|Dimensiona a altura da forma por um fator especificado. Para imagens, é possível indicar se você deseja dimensionar a forma em relação ao tamanho original ou ao tamanho atual. As formas que não são figuras serão sempre dimensionadas em relação à sua altura atual.|
||[scaleWidth(scaleFactor: number, scaleType: "CurrentSize" \| "OriginalSize", scaleFrom?: "ScaleFromTopLeft" \| "ScaleFromMiddle" \| "ScaleFromBottomRight")](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Dimensiona a largura da forma por um fator especificado. Para imagens, é possível indicar se você deseja dimensionar a forma em relação ao tamanho original ou ao tamanho atual. As formas que não são figuras serão sempre dimensionadas em relação à sua largura atual.|
||[scaleWidth(scaleFactor: number, scaleType: Excel.ShapeScaleType, scaleFrom?: Excel.ShapeScaleFrom)](/javascript/api/excel/excel.shape#scalewidth-scalefactor--scaletype--scalefrom-)|Dimensiona a largura da forma por um fator especificado. Para imagens, é possível indicar se você deseja dimensionar a forma em relação ao tamanho original ou ao tamanho atual. As formas que não são figuras serão sempre dimensionadas em relação à sua largura atual.|
||[Set (Propriedades: Excel. Shape)](/javascript/api/excel/excel.shape#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ShapeUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.shape#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[setZOrder(position: "BringToFront" \| "BringForward" \| "SendToBack" \| "SendBackward")](/javascript/api/excel/excel.shape#setzorder-position-)|Move a forma especificada para cima ou para baixo na ordem z da coleção, que a desloca para frente ou para trás de outras formas.|
||[setZOrder(position: Excel.ShapeZOrder)](/javascript/api/excel/excel.shape#setzorder-position-)|Move a forma especificada para cima ou para baixo na ordem z da coleção, que a desloca para frente ou para trás de outras formas.|
||[top](/javascript/api/excel/excel.shape#top)|A distância, em pontos, da borda superior da forma até a borda superior da planilha.|
||[visible](/javascript/api/excel/excel.shape#visible)|Representa a visibilidade essa forma.|
||[width](/javascript/api/excel/excel.shape#width)|Representa a largura, em pontos, da forma.|
|[ShapeActivatedEventArgs](/javascript/api/excel/excel.shapeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapeactivatedeventargs#shapeid)|Obtém o id da forma ativada.|
||[tipo](/javascript/api/excel/excel.shapeactivatedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.shapeactivatedeventargs#worksheetid)|Obtém a id da planilha na qual a forma está ativada.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addGeometricShape(geometricShapeType: "LineInverse" \| "Triangle" \| "RightTriangle" \| "Rectangle" \| "Diamond" \| "Parallelogram" \| "Trapezoid" \| "NonIsoscelesTrapezoid" \| "Pentagon" \| "Hexagon" \| "Heptagon" \| "Octagon" \| "Decagon" \| "Dodecagon" \| "Star4" \| "Star5" \| "Star6" \| "Star7" \| "Star8" \| "Star10" \| "Star12" \| "Star16" \| "Star24" \| "Star32" \| "RoundRectangle" \| "Round1Rectangle" \| "Round2SameRectangle" \| "Round2DiagonalRectangle" \| "SnipRoundRectangle" \| "Snip1Rectangle" \| "Snip2SameRectangle" \| "Snip2DiagonalRectangle" \| "Plaque" \| "Ellipse" \| "Teardrop" \| "HomePlate" \| "Chevron" \| "PieWedge" \| "Pie" \| "BlockArc" \| "Donut" \| "NoSmoking" \| "RightArrow" \| "LeftArrow" \| "UpArrow" \| "DownArrow" \| "StripedRightArrow" \| "NotchedRightArrow" \| "BentUpArrow" \| "LeftRightArrow" \| "UpDownArrow" \| "LeftUpArrow" \| "LeftRightUpArrow" \| "QuadArrow" \| "LeftArrowCallout" \| "RightArrowCallout" \| "UpArrowCallout" \| "DownArrowCallout" \| "LeftRightArrowCallout" \| "UpDownArrowCallout" \| "QuadArrowCallout" \| "BentArrow" \| "UturnArrow" \| "CircularArrow" \| "LeftCircularArrow" \| "LeftRightCircularArrow" \| "CurvedRightArrow" \| "CurvedLeftArrow" \| "CurvedUpArrow" \| "CurvedDownArrow" \| "SwooshArrow" \| "Cube" \| "Can" \| "LightningBolt" \| "Heart" \| "Sun" \| "Moon" \| "SmileyFace" \| "IrregularSeal1" \| "IrregularSeal2" \| "FoldedCorner" \| "Bevel" \| "Frame" \| "HalfFrame" \| "Corner" \| "DiagonalStripe" \| "Chord" \| "Arc" \| "LeftBracket" \| "RightBracket" \| "LeftBrace" \| "RightBrace" \| "BracketPair" \| "BracePair" \| "Callout1" \| "Callout2" \| "Callout3" \| "AccentCallout1" \| "AccentCallout2" \| "AccentCallout3" \| "BorderCallout1" \| "BorderCallout2" \| "BorderCallout3" \| "AccentBorderCallout1" \| "AccentBorderCallout2" \| "AccentBorderCallout3" \| "WedgeRectCallout" \| "WedgeRRectCallout" \| "WedgeEllipseCallout" \| "CloudCallout" \| "Cloud" \| "Ribbon" \| "Ribbon2" \| "EllipseRibbon" \| "EllipseRibbon2" \| "LeftRightRibbon" \| "VerticalScroll" \| "HorizontalScroll" \| "Wave" \| "DoubleWave" \| "Plus" \| "FlowChartProcess" \| "FlowChartDecision" \| "FlowChartInputOutput" \| "FlowChartPredefinedProcess" \| "FlowChartInternalStorage" \| "FlowChartDocument" \| "FlowChartMultidocument" \| "FlowChartTerminator" \| "FlowChartPreparation" \| "FlowChartManualInput" \| "FlowChartManualOperation" \| "FlowChartConnector" \| "FlowChartPunchedCard" \| "FlowChartPunchedTape" \| "FlowChartSummingJunction" \| "FlowChartOr" \| "FlowChartCollate" \| "FlowChartSort" \| "FlowChartExtract" \| "FlowChartMerge" \| "FlowChartOfflineStorage" \| "FlowChartOnlineStorage" \| "FlowChartMagneticTape" \| "FlowChartMagneticDisk" \| "FlowChartMagneticDrum" \| "FlowChartDisplay" \| "FlowChartDelay" \| "FlowChartAlternateProcess" \| "FlowChartOffpageConnector" \| "ActionButtonBlank" \| "ActionButtonHome" \| "ActionButtonHelp" \| "ActionButtonInformation" \| "ActionButtonForwardNext" \| "ActionButtonBackPrevious" \| "ActionButtonEnd" \| "ActionButtonBeginning" \| "ActionButtonReturn" \| "ActionButtonDocument" \| "ActionButtonSound" \| "ActionButtonMovie" \| "Gear6" \| "Gear9" \| "Funnel" \| "MathPlus" \| "MathMinus" \| "MathMultiply" \| "MathDivide" \| "MathEqual" \| "MathNotEqual" \| "CornerTabs" \| "SquareTabs" \| "PlaqueTabs" \| "ChartX" \| "ChartStar" \| "ChartPlus”)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Adiciona uma forma geométrica à planilha. Retorna um objeto Shape que representa a nova forma.|
||[addGeometricShape(geometricShapeType: Excel.GeometricShapeType)](/javascript/api/excel/excel.shapecollection#addgeometricshape-geometricshapetype-)|Adiciona uma forma geométrica à planilha. Retorna um objeto Shape que representa a nova forma.|
||[addGroup(values: Array<string \| Shape>)](/javascript/api/excel/excel.shapecollection#addgroup-values-)|Um subconjunto de formas na planilha do conjunto de grupos. Retorna um objeto Shape que representa o novo grupo de formas.|
||[addImage(base64ImageString: string)](/javascript/api/excel/excel.shapecollection#addimage-base64imagestring-)|Cria uma imagem de uma cadeia de caracteres na base 64 e a adiciona à planilha. Retorna o objeto Shape que representa a nova imagem.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: "Straight" \| "Elbow" \| "Curve")](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Adiciona uma linha à planilha. Retorna um objeto Shape que representa a nova linha.|
||[addLine(startLeft: number, startTop: number, endLeft: number, endTop: number, connectorType?: Excel.ConnectorType)](/javascript/api/excel/excel.shapecollection#addline-startleft--starttop--endleft--endtop--connectortype-)|Adiciona uma linha à planilha. Retorna um objeto Shape que representa a nova linha.|
||[addTextBox(text?: string)](/javascript/api/excel/excel.shapecollection#addtextbox-text-)|Adiciona uma caixa de texto na planilha com o texto fornecido como conteúdo. Retorna um objeto Shape que representa a nova caixa de texto.|
||[getCount()](/javascript/api/excel/excel.shapecollection#getcount--)|Retorna o número de formas da planilha. Somente leitura.|
||[getItem(key: string)](/javascript/api/excel/excel.shapecollection#getitem-key-)|Obtém uma forma usando seu nome ou ID.|
||[getItemAt(index: number)](/javascript/api/excel/excel.shapecollection#getitemat-index-)|Obtém uma forma usando sua posição na coleção.|
||[items](/javascript/api/excel/excel.shapecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ShapeCollectionLoadOptions](/javascript/api/excel/excel.shapecollectionloadoptions)|[$all](/javascript/api/excel/excel.shapecollectionloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.shapecollectionloadoptions#alttextdescription)|Para cada ITEM na coleção: Retorna ou define o texto de descrição alternativa para um objeto Shape.|
||[altTextTitle](/javascript/api/excel/excel.shapecollectionloadoptions#alttexttitle)|Para cada ITEM na coleção: Retorna ou define o texto de título alternativo para um objeto Shape.|
||[connectionSiteCount](/javascript/api/excel/excel.shapecollectionloadoptions#connectionsitecount)|Para cada ITEM na coleção: retorna o número de sites de conexão nesta forma. Somente leitura.|
||[fill](/javascript/api/excel/excel.shapecollectionloadoptions#fill)|Para cada ITEM na coleção: retorna a formatação de preenchimento dessa forma.|
||[geometricShape](/javascript/api/excel/excel.shapecollectionloadoptions#geometricshape)|Para cada ITEM na coleção: retorna a forma geométrica associada à forma. Um erro será lançado, se o tipo de forma não for "GeometricShape".|
||[geometricShapeType](/javascript/api/excel/excel.shapecollectionloadoptions#geometricshapetype)|Para cada ITEM na coleção: representa o tipo de forma geométrica dessa forma geométrica. Para saber detalhes, confira Excel.GeometricShapeType. Retorna nulo se o tipo de forma não for "GeometricShape".|
||[group](/javascript/api/excel/excel.shapecollectionloadoptions#group)|Para cada ITEM na coleção: retorna o grupo de formas associado à forma. Um erro será lançado, se o tipo de forma não for "GroupShape".|
||[height](/javascript/api/excel/excel.shapecollectionloadoptions#height)|Para cada ITEM na coleção: representa a altura, em pontos, da forma.|
||[id](/javascript/api/excel/excel.shapecollectionloadoptions#id)|Para cada ITEM na coleção: representa o identificador da forma. Somente leitura.|
||[image](/javascript/api/excel/excel.shapecollectionloadoptions#image)|Para cada ITEM na coleção: retorna a imagem associada à forma. Um erro será lançado, se o tipo de forma não for "Imagem".|
||[left](/javascript/api/excel/excel.shapecollectionloadoptions#left)|Para cada ITEM na coleção: a distância, em pontos, do lado esquerdo da forma até o lado esquerdo da planilha.|
||[level](/javascript/api/excel/excel.shapecollectionloadoptions#level)|Para cada ITEM na coleção: representa o nível da forma especificada. Por exemplo, um nível de 0 significa que a forma não faz parte de nenhum grupo, um nível de 1 significa que a forma é parte de um grupo de nível superior e um nível 2 significa que a forma faz parte de um subgrupo do nível superior.|
||[line](/javascript/api/excel/excel.shapecollectionloadoptions#line)|Para cada ITEM na coleção: retorna a linha associada à forma. Um erro será lançado, se o tipo de forma não for "Linha".|
||[lineFormat](/javascript/api/excel/excel.shapecollectionloadoptions#lineformat)|Para cada ITEM na coleção: retorna a formatação de linha dessa forma.|
||[lockAspectRatio](/javascript/api/excel/excel.shapecollectionloadoptions#lockaspectratio)|Para cada ITEM na coleção: especifica se a taxa de proporção dessa forma será ou não bloqueada.|
||[name](/javascript/api/excel/excel.shapecollectionloadoptions#name)|Para cada ITEM na coleção: representa o nome da forma.|
||[parentGroup](/javascript/api/excel/excel.shapecollectionloadoptions#parentgroup)|Para cada ITEM na coleção: representa o grupo pai desta forma.|
||[rotation](/javascript/api/excel/excel.shapecollectionloadoptions#rotation)|Para cada ITEM na coleção: representa a rotação, em graus, da forma.|
||[textFrame](/javascript/api/excel/excel.shapecollectionloadoptions#textframe)|Para cada ITEM na coleção: retorna o objeto de quadro de texto desta forma. Somente leitura.|
||[top](/javascript/api/excel/excel.shapecollectionloadoptions#top)|Para cada ITEM na coleção: a distância, em pontos, da borda superior da forma até a borda superior da planilha.|
||[tipo](/javascript/api/excel/excel.shapecollectionloadoptions#type)|Para cada ITEM na coleção: retorna o tipo dessa forma. Para saber detalhes, confira Excel.ShapeType. Somente leitura.|
||[visible](/javascript/api/excel/excel.shapecollectionloadoptions#visible)|Para cada ITEM na coleção: representa a visibilidade dessa forma.|
||[width](/javascript/api/excel/excel.shapecollectionloadoptions#width)|Para cada ITEM na coleção: representa a largura, em pontos, da forma.|
||[zOrderPosition](/javascript/api/excel/excel.shapecollectionloadoptions#zorderposition)|Para cada ITEM na coleção: retorna a posição da forma especificada na ordem z, com 0 que representa a parte inferior da pilha da ordem. Somente leitura.|
|[ShapeData](/javascript/api/excel/excel.shapedata)|[altTextDescription](/javascript/api/excel/excel.shapedata#alttextdescription)|Retorna ou define o texto da descrição alternativa de um objeto de forma.|
||[altTextTitle](/javascript/api/excel/excel.shapedata#alttexttitle)|Retorna ou define o texto do título alternativo de um objeto de forma.|
||[connectionSiteCount](/javascript/api/excel/excel.shapedata#connectionsitecount)|Retorna o número de locais de conexão nessa forma. Somente leitura.|
||[fill](/javascript/api/excel/excel.shapedata#fill)|Retorna a formatação de preenchimento dessa forma. Somente leitura.|
||[geometricShapeType](/javascript/api/excel/excel.shapedata#geometricshapetype)|Representa o tipo de forma geométricas da forma geométrica. Para saber detalhes, confira Excel.GeometricShapeType. Retorna nulo se o tipo de forma não for "GeometricShape".|
||[height](/javascript/api/excel/excel.shapedata#height)|Representa a altura, em pontos, da forma.|
||[id](/javascript/api/excel/excel.shapedata#id)|Representa o identificador de forma. Somente leitura.|
||[left](/javascript/api/excel/excel.shapedata#left)|A distância, em pontos, da lateral esquerda da forma do lado  esquerdo da planilha.|
||[level](/javascript/api/excel/excel.shapedata#level)|Representa o nível da forma especificada. Por exemplo, um nível de 0 significa que a forma não faz parte de nenhum grupo, um nível de 1 significa que a forma é parte de um grupo de nível superior e um nível 2 significa que a forma faz parte de um subgrupo do nível superior.|
||[lineFormat](/javascript/api/excel/excel.shapedata#lineformat)|Retorna a formatação de linha do objeto de forma. Somente leitura.|
||[lockAspectRatio](/javascript/api/excel/excel.shapedata#lockaspectratio)|Especifica se a taxa de proporção dessa forma está bloqueada ou não.|
||[name](/javascript/api/excel/excel.shapedata#name)|Representa o nome da forma.|
||[rotation](/javascript/api/excel/excel.shapedata#rotation)|Representa a rotação, em graus, da forma.|
||[top](/javascript/api/excel/excel.shapedata#top)|A distância, em pontos, da borda superior da forma até a borda superior da planilha.|
||[tipo](/javascript/api/excel/excel.shapedata#type)|Retorna o tipo dessa forma. Para saber detalhes, confira Excel.ShapeType. Somente leitura.|
||[visible](/javascript/api/excel/excel.shapedata#visible)|Representa a visibilidade essa forma.|
||[width](/javascript/api/excel/excel.shapedata#width)|Representa a largura, em pontos, da forma.|
||[zOrderPosition](/javascript/api/excel/excel.shapedata#zorderposition)|Retorna a posição da forma especificada na ordem z, com 0 representando a parte inferior da pilha do pedido. Somente leitura.|
|[ShapeDeactivatedEventArgs](/javascript/api/excel/excel.shapedeactivatedeventargs)|[shapeId](/javascript/api/excel/excel.shapedeactivatedeventargs#shapeid)|Obtém o id da forma que está desativada.|
||[tipo](/javascript/api/excel/excel.shapedeactivatedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.shapedeactivatedeventargs#worksheetid)|Obtém a id da planilha na qual a forma está desativada.|
|[ShapeFill](/javascript/api/excel/excel.shapefill)|[clear()](/javascript/api/excel/excel.shapefill#clear--)|Limpa a formatação do preenchimento de um objeto de forma.|
||[foregroundColor](/javascript/api/excel/excel.shapefill#foregroundcolor)|Representa o primeiro plano de preenchimento da forma para cor no formato de cor HTML, no formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[tipo](/javascript/api/excel/excel.shapefill#type)|Retorna o tipo de preenchimento da forma. Somente leitura. Para saber detalhes, confira Excel.ShapeFillType.|
||[Set (Propriedades: Excel. ShapeFill)](/javascript/api/excel/excel.shapefill#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ShapeFillUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.shapefill#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[setSolidColor(color: string)](/javascript/api/excel/excel.shapefill#setsolidcolor-color-)|Define a formatação de preenchimento de um formato com uma cor uniforme. Isso altera o tipo de preenchimento para "Sólido".|
||[transparency](/javascript/api/excel/excel.shapefill#transparency)|Retorna ou define a porcentagem de transparência do preenchimento especificado como um valor de 0,0 (opaco) a 1,0 (transparente). Retorna nulo se o tipo de forma não suportar transparência ou se o preenchimento de forma tiver transparência inconsistente como com um tipo de preenchimento de gradiente.|
|[ShapeFillData](/javascript/api/excel/excel.shapefilldata)|[foregroundColor](/javascript/api/excel/excel.shapefilldata#foregroundcolor)|Representa o primeiro plano de preenchimento da forma para cor no formato de cor HTML, no formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[transparency](/javascript/api/excel/excel.shapefilldata#transparency)|Retorna ou define a porcentagem de transparência do preenchimento especificado como um valor de 0,0 (opaco) a 1,0 (transparente). Retorna nulo se o tipo de forma não suportar transparência ou se o preenchimento de forma tiver transparência inconsistente como com um tipo de preenchimento de gradiente.|
||[tipo](/javascript/api/excel/excel.shapefilldata#type)|Retorna o tipo de preenchimento da forma. Somente leitura. Para saber detalhes, confira Excel.ShapeFillType.|
|[ShapeFillLoadOptions](/javascript/api/excel/excel.shapefillloadoptions)|[$all](/javascript/api/excel/excel.shapefillloadoptions#$all)||
||[foregroundColor](/javascript/api/excel/excel.shapefillloadoptions#foregroundcolor)|Representa o primeiro plano de preenchimento da forma para cor no formato de cor HTML, no formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[transparency](/javascript/api/excel/excel.shapefillloadoptions#transparency)|Retorna ou define a porcentagem de transparência do preenchimento especificado como um valor de 0,0 (opaco) a 1,0 (transparente). Retorna nulo se o tipo de forma não suportar transparência ou se o preenchimento de forma tiver transparência inconsistente como com um tipo de preenchimento de gradiente.|
||[tipo](/javascript/api/excel/excel.shapefillloadoptions#type)|Retorna o tipo de preenchimento da forma. Somente leitura. Para saber detalhes, confira Excel.ShapeFillType.|
|[ShapeFillUpdateData](/javascript/api/excel/excel.shapefillupdatedata)|[foregroundColor](/javascript/api/excel/excel.shapefillupdatedata#foregroundcolor)|Representa o primeiro plano de preenchimento da forma para cor no formato de cor HTML, no formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[transparency](/javascript/api/excel/excel.shapefillupdatedata#transparency)|Retorna ou define a porcentagem de transparência do preenchimento especificado como um valor de 0,0 (opaco) a 1,0 (transparente). Retorna nulo se o tipo de forma não suportar transparência ou se o preenchimento de forma tiver transparência inconsistente como com um tipo de preenchimento de gradiente.|
|[ShapeFont](/javascript/api/excel/excel.shapefont)|[bold](/javascript/api/excel/excel.shapefont#bold)|Representa o status da fonte em negrito. Retornará null se o TextRange incluir fragmentos de texto em negrito e não em negrito.|
||[color](/javascript/api/excel/excel.shapefont#color)|A representação de código de cor HTML para a cor do texto. (Por exemplo, #FF0000 representa vermelho). Retornará null se o TextRange incluir fragmentos de texto com cores diferentes.|
||[italic](/javascript/api/excel/excel.shapefont#italic)|Representa o status da fonte em itálico. Retorna null se o TextRange incluir fragmentos de texto em itálico e que não está em itálico.|
||[name](/javascript/api/excel/excel.shapefont#name)|Representa o nome da fonte (por exemplo, "Calibri"). Se o texto estiver no idioma Script Complexo ou Leste Asiático, esse é o nome da fonte correspondente. Caso contrário, esse é o nome da fonte Latin.|
||[Set (Propriedades: Excel. ShapeFont)](/javascript/api/excel/excel.shapefont#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ShapeFontUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.shapefont#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[size](/javascript/api/excel/excel.shapefont#size)|Representa o tamanho da fonte em pontos (por exemplo, 11). Retorna nulo se o TextRange incluir fragmentos de texto com tamanhos de fontes diferentes.|
||[underline](/javascript/api/excel/excel.shapefont#underline)|Tipo de sublinhado aplicado à fonte. Retorna nulo se o TextRange incluir fragmentos de texto com estilos de sublinhado diferentes. Para saber detalhes, confira Excel.ShapeFontUnderlineStyle.|
|[ShapeFontData](/javascript/api/excel/excel.shapefontdata)|[bold](/javascript/api/excel/excel.shapefontdata#bold)|Representa o status da fonte em negrito. Retornará null se o TextRange incluir fragmentos de texto em negrito e não em negrito.|
||[color](/javascript/api/excel/excel.shapefontdata#color)|A representação de código de cor HTML para a cor do texto. (Por exemplo, #FF0000 representa vermelho). Retornará null se o TextRange incluir fragmentos de texto com cores diferentes.|
||[italic](/javascript/api/excel/excel.shapefontdata#italic)|Representa o status da fonte em itálico. Retorna null se o TextRange incluir fragmentos de texto em itálico e que não está em itálico.|
||[name](/javascript/api/excel/excel.shapefontdata#name)|Representa o nome da fonte (por exemplo, "Calibri"). Se o texto estiver no idioma Script Complexo ou Leste Asiático, esse é o nome da fonte correspondente. Caso contrário, esse é o nome da fonte Latin.|
||[size](/javascript/api/excel/excel.shapefontdata#size)|Representa o tamanho da fonte em pontos (por exemplo, 11). Retorna nulo se o TextRange incluir fragmentos de texto com tamanhos de fontes diferentes.|
||[underline](/javascript/api/excel/excel.shapefontdata#underline)|Tipo de sublinhado aplicado à fonte. Retorna nulo se o TextRange incluir fragmentos de texto com estilos de sublinhado diferentes. Para saber detalhes, confira Excel.ShapeFontUnderlineStyle.|
|[ShapeFontLoadOptions](/javascript/api/excel/excel.shapefontloadoptions)|[$all](/javascript/api/excel/excel.shapefontloadoptions#$all)||
||[bold](/javascript/api/excel/excel.shapefontloadoptions#bold)|Representa o status da fonte em negrito. Retornará null se o TextRange incluir fragmentos de texto em negrito e não em negrito.|
||[color](/javascript/api/excel/excel.shapefontloadoptions#color)|A representação de código de cor HTML para a cor do texto. (Por exemplo, #FF0000 representa vermelho). Retornará null se o TextRange incluir fragmentos de texto com cores diferentes.|
||[italic](/javascript/api/excel/excel.shapefontloadoptions#italic)|Representa o status da fonte em itálico. Retorna null se o TextRange incluir fragmentos de texto em itálico e que não está em itálico.|
||[name](/javascript/api/excel/excel.shapefontloadoptions#name)|Representa o nome da fonte (por exemplo, "Calibri"). Se o texto estiver no idioma Script Complexo ou Leste Asiático, esse é o nome da fonte correspondente. Caso contrário, esse é o nome da fonte Latin.|
||[size](/javascript/api/excel/excel.shapefontloadoptions#size)|Representa o tamanho da fonte em pontos (por exemplo, 11). Retorna nulo se o TextRange incluir fragmentos de texto com tamanhos de fontes diferentes.|
||[underline](/javascript/api/excel/excel.shapefontloadoptions#underline)|Tipo de sublinhado aplicado à fonte. Retorna nulo se o TextRange incluir fragmentos de texto com estilos de sublinhado diferentes. Para saber detalhes, confira Excel.ShapeFontUnderlineStyle.|
|[ShapeFontUpdateData](/javascript/api/excel/excel.shapefontupdatedata)|[bold](/javascript/api/excel/excel.shapefontupdatedata#bold)|Representa o status da fonte em negrito. Retornará null se o TextRange incluir fragmentos de texto em negrito e não em negrito.|
||[color](/javascript/api/excel/excel.shapefontupdatedata#color)|A representação de código de cor HTML para a cor do texto. (Por exemplo, #FF0000 representa vermelho). Retornará null se o TextRange incluir fragmentos de texto com cores diferentes.|
||[italic](/javascript/api/excel/excel.shapefontupdatedata#italic)|Representa o status da fonte em itálico. Retorna null se o TextRange incluir fragmentos de texto em itálico e que não está em itálico.|
||[name](/javascript/api/excel/excel.shapefontupdatedata#name)|Representa o nome da fonte (por exemplo, "Calibri"). Se o texto estiver no idioma Script Complexo ou Leste Asiático, esse é o nome da fonte correspondente. Caso contrário, esse é o nome da fonte Latin.|
||[size](/javascript/api/excel/excel.shapefontupdatedata#size)|Representa o tamanho da fonte em pontos (por exemplo, 11). Retorna nulo se o TextRange incluir fragmentos de texto com tamanhos de fontes diferentes.|
||[underline](/javascript/api/excel/excel.shapefontupdatedata#underline)|Tipo de sublinhado aplicado à fonte. Retorna nulo se o TextRange incluir fragmentos de texto com estilos de sublinhado diferentes. Para saber detalhes, confira Excel.ShapeFontUnderlineStyle.|
|[ShapeGroup](/javascript/api/excel/excel.shapegroup)|[id](/javascript/api/excel/excel.shapegroup#id)|Representa o identificador de forma. Somente leitura.|
||[shape](/javascript/api/excel/excel.shapegroup#shape)|Retorna o objeto de forma associado ao grupo. Somente leitura.|
||[shapes](/javascript/api/excel/excel.shapegroup#shapes)|Retorna uma coleção de objetos de forma. Somente leitura.|
||[ungroup()](/javascript/api/excel/excel.shapegroup#ungroup--)|Desagrupa todas as formas agrupadas no grupo de forma especificado.|
|[ShapeGroupData](/javascript/api/excel/excel.shapegroupdata)|[id](/javascript/api/excel/excel.shapegroupdata#id)|Representa o identificador de forma. Somente leitura.|
||[shapes](/javascript/api/excel/excel.shapegroupdata#shapes)|Retorna uma coleção de objetos de forma. Somente leitura.|
|[ShapeGroupLoadOptions](/javascript/api/excel/excel.shapegrouploadoptions)|[$all](/javascript/api/excel/excel.shapegrouploadoptions#$all)||
||[id](/javascript/api/excel/excel.shapegrouploadoptions#id)|Representa o identificador de forma. Somente leitura.|
||[shape](/javascript/api/excel/excel.shapegrouploadoptions#shape)|Retorna o objeto de forma associado ao grupo.|
|[ShapeLineFormat](/javascript/api/excel/excel.shapelineformat)|[color](/javascript/api/excel/excel.shapelineformat#color)|Representa a cor da linha no formato de cor HTML, no formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[dashStyle](/javascript/api/excel/excel.shapelineformat#dashstyle)|Representa o estilo de linha da forma. Retorna nulo quando a linha não estiver visível ou quando existirem estilos de traços inconsistentes. Para saber detalhes, confira Excel.ShapeLineStyle.|
||[Set (Propriedades: Excel. ShapeLineFormat)](/javascript/api/excel/excel.shapelineformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ShapeLineFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.shapelineformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[style](/javascript/api/excel/excel.shapelineformat#style)|Representa o estilo de linha da forma. Retorna nulo quando a linha não estiver visível ou quando existirem estilos inconsistentes. Para saber detalhes, confira Excel.ShapeLineStyle.|
||[transparency](/javascript/api/excel/excel.shapelineformat#transparency)|Representa o grau de transparência da linha especificada como um valor de 0,0 (opaco) a 1,0 (claro). Retorna nulo quando a forma possui transparências inconsistentes.|
||[visible](/javascript/api/excel/excel.shapelineformat#visible)|Representa se a formatação de linha de um elemento de forma é visível ou não. Retorna nulo quando a forma possui visibilidades inconsistentes.|
||[weight](/javascript/api/excel/excel.shapelineformat#weight)|Representa a espessura da linha, em pontos. Retorna nulo quando não a linha não estiver visível ou existirem espessuras de linha inconsistentes.|
|[ShapeLineFormatData](/javascript/api/excel/excel.shapelineformatdata)|[color](/javascript/api/excel/excel.shapelineformatdata#color)|Representa a cor da linha no formato de cor HTML, no formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[dashStyle](/javascript/api/excel/excel.shapelineformatdata#dashstyle)|Representa o estilo de linha da forma. Retorna nulo quando a linha não estiver visível ou quando existirem estilos de traços inconsistentes. Para saber detalhes, confira Excel.ShapeLineStyle.|
||[style](/javascript/api/excel/excel.shapelineformatdata#style)|Representa o estilo de linha da forma. Retorna nulo quando a linha não estiver visível ou quando existirem estilos inconsistentes. Para saber detalhes, confira Excel.ShapeLineStyle.|
||[transparency](/javascript/api/excel/excel.shapelineformatdata#transparency)|Representa o grau de transparência da linha especificada como um valor de 0,0 (opaco) a 1,0 (claro). Retorna nulo quando a forma possui transparências inconsistentes.|
||[visible](/javascript/api/excel/excel.shapelineformatdata#visible)|Representa se a formatação de linha de um elemento de forma é visível ou não. Retorna nulo quando a forma possui visibilidades inconsistentes.|
||[weight](/javascript/api/excel/excel.shapelineformatdata#weight)|Representa a espessura da linha, em pontos. Retorna nulo quando não a linha não estiver visível ou existirem espessuras de linha inconsistentes.|
|[ShapeLineFormatLoadOptions](/javascript/api/excel/excel.shapelineformatloadoptions)|[$all](/javascript/api/excel/excel.shapelineformatloadoptions#$all)||
||[color](/javascript/api/excel/excel.shapelineformatloadoptions#color)|Representa a cor da linha no formato de cor HTML, no formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[dashStyle](/javascript/api/excel/excel.shapelineformatloadoptions#dashstyle)|Representa o estilo de linha da forma. Retorna nulo quando a linha não estiver visível ou quando existirem estilos de traços inconsistentes. Para saber detalhes, confira Excel.ShapeLineStyle.|
||[style](/javascript/api/excel/excel.shapelineformatloadoptions#style)|Representa o estilo de linha da forma. Retorna nulo quando a linha não estiver visível ou quando existirem estilos inconsistentes. Para saber detalhes, confira Excel.ShapeLineStyle.|
||[transparency](/javascript/api/excel/excel.shapelineformatloadoptions#transparency)|Representa o grau de transparência da linha especificada como um valor de 0,0 (opaco) a 1,0 (claro). Retorna nulo quando a forma possui transparências inconsistentes.|
||[visible](/javascript/api/excel/excel.shapelineformatloadoptions#visible)|Representa se a formatação de linha de um elemento de forma é visível ou não. Retorna nulo quando a forma possui visibilidades inconsistentes.|
||[weight](/javascript/api/excel/excel.shapelineformatloadoptions#weight)|Representa a espessura da linha, em pontos. Retorna nulo quando não a linha não estiver visível ou existirem espessuras de linha inconsistentes.|
|[ShapeLineFormatUpdateData](/javascript/api/excel/excel.shapelineformatupdatedata)|[color](/javascript/api/excel/excel.shapelineformatupdatedata#color)|Representa a cor da linha no formato de cor HTML, no formulário #RRGGBB (por exemplo, "FFA500") ou como uma cor HTML nomeada (por exemplo, "laranja")|
||[dashStyle](/javascript/api/excel/excel.shapelineformatupdatedata#dashstyle)|Representa o estilo de linha da forma. Retorna nulo quando a linha não estiver visível ou quando existirem estilos de traços inconsistentes. Para saber detalhes, confira Excel.ShapeLineStyle.|
||[style](/javascript/api/excel/excel.shapelineformatupdatedata#style)|Representa o estilo de linha da forma. Retorna nulo quando a linha não estiver visível ou quando existirem estilos inconsistentes. Para saber detalhes, confira Excel.ShapeLineStyle.|
||[transparency](/javascript/api/excel/excel.shapelineformatupdatedata#transparency)|Representa o grau de transparência da linha especificada como um valor de 0,0 (opaco) a 1,0 (claro). Retorna nulo quando a forma possui transparências inconsistentes.|
||[visible](/javascript/api/excel/excel.shapelineformatupdatedata#visible)|Representa se a formatação de linha de um elemento de forma é visível ou não. Retorna nulo quando a forma possui visibilidades inconsistentes.|
||[weight](/javascript/api/excel/excel.shapelineformatupdatedata#weight)|Representa a espessura da linha, em pontos. Retorna nulo quando não a linha não estiver visível ou existirem espessuras de linha inconsistentes.|
|[ShapeLoadOptions](/javascript/api/excel/excel.shapeloadoptions)|[$all](/javascript/api/excel/excel.shapeloadoptions#$all)||
||[altTextDescription](/javascript/api/excel/excel.shapeloadoptions#alttextdescription)|Retorna ou define o texto da descrição alternativa de um objeto de forma.|
||[altTextTitle](/javascript/api/excel/excel.shapeloadoptions#alttexttitle)|Retorna ou define o texto do título alternativo de um objeto de forma.|
||[connectionSiteCount](/javascript/api/excel/excel.shapeloadoptions#connectionsitecount)|Retorna o número de locais de conexão nessa forma. Somente leitura.|
||[fill](/javascript/api/excel/excel.shapeloadoptions#fill)|Retorna a formatação de preenchimento dessa forma.|
||[geometricShape](/javascript/api/excel/excel.shapeloadoptions#geometricshape)|Retorna a forma geométrica associada à forma. Um erro será lançado, se o tipo de forma não for "GeometricShape".|
||[geometricShapeType](/javascript/api/excel/excel.shapeloadoptions#geometricshapetype)|Representa o tipo de forma geométricas da forma geométrica. Para saber detalhes, confira Excel.GeometricShapeType. Retorna nulo se o tipo de forma não for "GeometricShape".|
||[group](/javascript/api/excel/excel.shapeloadoptions#group)|Retorna o grupo de forma associado à forma. Um erro será lançado, se o tipo de forma não for "GroupShape".|
||[height](/javascript/api/excel/excel.shapeloadoptions#height)|Representa a altura, em pontos, da forma.|
||[id](/javascript/api/excel/excel.shapeloadoptions#id)|Representa o identificador de forma. Somente leitura.|
||[image](/javascript/api/excel/excel.shapeloadoptions#image)|Retorna a imagem associada à forma. Um erro será lançado, se o tipo de forma não for "Imagem".|
||[left](/javascript/api/excel/excel.shapeloadoptions#left)|A distância, em pontos, da lateral esquerda da forma do lado  esquerdo da planilha.|
||[level](/javascript/api/excel/excel.shapeloadoptions#level)|Representa o nível da forma especificada. Por exemplo, um nível de 0 significa que a forma não faz parte de nenhum grupo, um nível de 1 significa que a forma é parte de um grupo de nível superior e um nível 2 significa que a forma faz parte de um subgrupo do nível superior.|
||[line](/javascript/api/excel/excel.shapeloadoptions#line)|Retorna a linha associada à forma. Um erro será lançado, se o tipo de forma não for "Linha".|
||[lineFormat](/javascript/api/excel/excel.shapeloadoptions#lineformat)|Retorna a formatação de linha do objeto de forma.|
||[lockAspectRatio](/javascript/api/excel/excel.shapeloadoptions#lockaspectratio)|Especifica se a taxa de proporção dessa forma está bloqueada ou não.|
||[name](/javascript/api/excel/excel.shapeloadoptions#name)|Representa o nome da forma.|
||[parentGroup](/javascript/api/excel/excel.shapeloadoptions#parentgroup)|Representa o grupo pai dessa forma.|
||[rotation](/javascript/api/excel/excel.shapeloadoptions#rotation)|Representa a rotação, em graus, da forma.|
||[textFrame](/javascript/api/excel/excel.shapeloadoptions#textframe)|Retorna o objeto text frame de uma forma. Somente leitura.|
||[top](/javascript/api/excel/excel.shapeloadoptions#top)|A distância, em pontos, da borda superior da forma até a borda superior da planilha.|
||[tipo](/javascript/api/excel/excel.shapeloadoptions#type)|Retorna o tipo dessa forma. Para saber detalhes, confira Excel.ShapeType. Somente leitura.|
||[visible](/javascript/api/excel/excel.shapeloadoptions#visible)|Representa a visibilidade essa forma.|
||[width](/javascript/api/excel/excel.shapeloadoptions#width)|Representa a largura, em pontos, da forma.|
||[zOrderPosition](/javascript/api/excel/excel.shapeloadoptions#zorderposition)|Retorna a posição da forma especificada na ordem z, com 0 representando a parte inferior da pilha do pedido. Somente leitura.|
|[ShapeUpdateData](/javascript/api/excel/excel.shapeupdatedata)|[altTextDescription](/javascript/api/excel/excel.shapeupdatedata#alttextdescription)|Retorna ou define o texto da descrição alternativa de um objeto de forma.|
||[altTextTitle](/javascript/api/excel/excel.shapeupdatedata#alttexttitle)|Retorna ou define o texto do título alternativo de um objeto de forma.|
||[fill](/javascript/api/excel/excel.shapeupdatedata#fill)|Retorna a formatação de preenchimento dessa forma.|
||[geometricShapeType](/javascript/api/excel/excel.shapeupdatedata#geometricshapetype)|Representa o tipo de forma geométricas da forma geométrica. Para saber detalhes, confira Excel.GeometricShapeType. Retorna nulo se o tipo de forma não for "GeometricShape".|
||[height](/javascript/api/excel/excel.shapeupdatedata#height)|Representa a altura, em pontos, da forma.|
||[left](/javascript/api/excel/excel.shapeupdatedata#left)|A distância, em pontos, da lateral esquerda da forma do lado  esquerdo da planilha.|
||[lineFormat](/javascript/api/excel/excel.shapeupdatedata#lineformat)|Retorna a formatação de linha do objeto de forma.|
||[lockAspectRatio](/javascript/api/excel/excel.shapeupdatedata#lockaspectratio)|Especifica se a taxa de proporção dessa forma está bloqueada ou não.|
||[name](/javascript/api/excel/excel.shapeupdatedata#name)|Representa o nome da forma.|
||[rotation](/javascript/api/excel/excel.shapeupdatedata#rotation)|Representa a rotação, em graus, da forma.|
||[top](/javascript/api/excel/excel.shapeupdatedata#top)|A distância, em pontos, da borda superior da forma até a borda superior da planilha.|
||[visible](/javascript/api/excel/excel.shapeupdatedata#visible)|Representa a visibilidade essa forma.|
||[width](/javascript/api/excel/excel.shapeupdatedata#width)|Representa a largura, em pontos, da forma.|
|[SortField](/javascript/api/excel/excel.sortfield)|[subField](/javascript/api/excel/excel.sortfield#subfield)|Representa o subcampo que é o nome da propriedade de destino de um valor avançado para classificação.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getCount()](/javascript/api/excel/excel.stylecollection#getcount--)|Obtém o número de estilos na coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.stylecollection#getitemat-index-)|Obtém um estilo com base em sua posição na coleção.|
|[Table](/javascript/api/excel/excel.table)|[autoFilter](/javascript/api/excel/excel.table#autofilter)|Representa o objeto AutoFilter da tabela. Somente Leitura.|
|[TableAddedEventArgs](/javascript/api/excel/excel.tableaddedeventargs)|[source](/javascript/api/excel/excel.tableaddedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tableaddedeventargs#tableid)|Obtém a id da tabela que é adicionada.|
||[tipo](/javascript/api/excel/excel.tableaddedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tableaddedeventargs#worksheetid)|Obtém a id da planilha na qual o gráfico é adicionado.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[detalhes](/javascript/api/excel/excel.tablechangedeventargs#details)|Representa informações sobre os detalhes da alteração|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onAdded](/javascript/api/excel/excel.tablecollection#onadded)|Ocorre quando uma nova tabela é adicionada na pasta de trabalho.|
||[onDeleted](/javascript/api/excel/excel.tablecollection#ondeleted)|Ocorre quando a tabela especificada é excluída em uma pasta de trabalho.|
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[autoFilter](/javascript/api/excel/excel.tablecollectionloadoptions#autofilter)|Para cada ITEM na coleção: representa o objeto AutoFilter da tabela.|
|[TableData](/javascript/api/excel/excel.tabledata)|[autoFilter](/javascript/api/excel/excel.tabledata#autofilter)|Representa o objeto AutoFilter da tabela. Somente Leitura.|
|[TableDeletedEventArgs](/javascript/api/excel/excel.tabledeletedeventargs)|[source](/javascript/api/excel/excel.tabledeletedeventargs#source)|Especifica a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tabledeletedeventargs#tableid)|Especifica a id da tabela que é excluída.|
||[tableName](/javascript/api/excel/excel.tabledeletedeventargs#tablename)|Especifica o nome da tabela que é excluída.|
||[tipo](/javascript/api/excel/excel.tabledeletedeventargs#type)|Especifica o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tabledeletedeventargs#worksheetid)|Obtém a id da planilha na qual a tabela é excluída.|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[autoFilter](/javascript/api/excel/excel.tableloadoptions#autofilter)|Representa o objeto AutoFilter da tabela.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getCount()](/javascript/api/excel/excel.tablescopedcollection#getcount--)|Obtém o número de tabelas na coleção.|
||[getFirst()](/javascript/api/excel/excel.tablescopedcollection#getfirst--)|Obtém a primeira tabela na coleção. As tabelas na coleção são classificadas de cima para baixo e da esquerda para a direita, de forma que a tabela superior esquerda seja a primeira tabela da coleção.|
||[getItem(key: string)](/javascript/api/excel/excel.tablescopedcollection#getitem-key-)|Obtém uma tabela pelo nome ou ID.|
||[items](/javascript/api/excel/excel.tablescopedcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[TableScopedCollectionLoadOptions](/javascript/api/excel/excel.tablescopedcollectionloadoptions)|[$all](/javascript/api/excel/excel.tablescopedcollectionloadoptions#$all)||
||[autoFilter](/javascript/api/excel/excel.tablescopedcollectionloadoptions#autofilter)|Para cada ITEM na coleção: representa o objeto AutoFilter da tabela.|
||[colunas](/javascript/api/excel/excel.tablescopedcollectionloadoptions#columns)|Para cada ITEM na coleção: representa uma coleção de todas as colunas da tabela.|
||[highlightFirstColumn](/javascript/api/excel/excel.tablescopedcollectionloadoptions#highlightfirstcolumn)|Para cada ITEM na coleção: indica se a primeira coluna contém formatação especial.|
||[highlightLastColumn](/javascript/api/excel/excel.tablescopedcollectionloadoptions#highlightlastcolumn)|Para cada ITEM na coleção: indica se a última coluna contém formatação especial.|
||[id](/javascript/api/excel/excel.tablescopedcollectionloadoptions#id)|Para cada ITEM na coleção: retorna um valor que identifica exclusivamente a tabela em uma determinada pasta de trabalho. O valor do identificador permanece o mesmo, ainda que a tabela seja renomeada. Somente leitura.|
||[legacyId](/javascript/api/excel/excel.tablescopedcollectionloadoptions#legacyid)|Para cada ITEM na coleção: retorna uma ID numérica.|
||[name](/javascript/api/excel/excel.tablescopedcollectionloadoptions#name)|Para cada ITEM na coleção: o nome da tabela.|
||[rows](/javascript/api/excel/excel.tablescopedcollectionloadoptions#rows)|Para cada ITEM na coleção: representa uma coleção de todas as linhas da tabela.|
||[showBandedColumns](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showbandedcolumns)|Para cada ITEM na coleção: indica se as colunas mostram a formatação em tiras nas quais as colunas ímpares são realçadas de forma diferente de mesmo para tornar a leitura da tabela mais fácil.|
||[showBandedRows](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showbandedrows)|Para cada ITEM na coleção: indica se as linhas mostram a formatação em tiras nas quais as linhas ímpares são realçadas de forma diferente de mesmo para tornar a leitura da tabela mais fácil.|
||[showFilterButton](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showfilterbutton)|Para cada ITEM na coleção: indica se os botões de filtro estão visíveis na parte superior de cada cabeçalho de coluna. Essa configuração só será permitida se a tabela tiver uma linha de cabeçalho.|
||[showHeaders](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showheaders)|Para cada ITEM na coleção: indica se a linha de cabeçalho está visível ou não. Esse valor pode ser definido para mostrar ou remover a linha do cabeçalho.|
||[showTotals](/javascript/api/excel/excel.tablescopedcollectionloadoptions#showtotals)|Para cada ITEM na coleção: indica se a linha de total está visível ou não. Esse valor pode ser definido para mostrar ou remover a linha do total.|
||[sort](/javascript/api/excel/excel.tablescopedcollectionloadoptions#sort)|Para cada ITEM na coleção: representa a classificação para a tabela.|
||[style](/javascript/api/excel/excel.tablescopedcollectionloadoptions#style)|Para cada ITEM da coleção: valor constante que representa o estilo de tabela. Os valores possíveis são: TableStyleLight1 a TableStyleLight21, TableStyleMedium1 a TableStyleMedium28, TableStyleStyleDark1 a TableStyleStyleDark11. Também é possível usar um estilo definido pelo usuário que esteja presente na planilha.|
||[worksheet](/javascript/api/excel/excel.tablescopedcollectionloadoptions#worksheet)|Para cada ITEM na coleção: a planilha que contém a tabela atual.|
|[TextFrame](/javascript/api/excel/excel.textframe)|[autoSizeSetting](/javascript/api/excel/excel.textframe#autosizesetting)|Obtém ou define as configurações de dimensionamento automático para o quadro de texto. Um quadro de texto pode ser configurado para ajustar automaticamente o texto ao quadro de texto, para ajustar automaticamente o quadro do texto ao texto ou não executar qualquer dimensionamento automático.|
||[bottomMargin](/javascript/api/excel/excel.textframe#bottommargin)|Representa margem inferior, em pontos, do quadro de texto.|
||[deleteText()](/javascript/api/excel/excel.textframe#deletetext--)|Exclui todo o texto no quadro de texto.|
||[horizontalAlignment](/javascript/api/excel/excel.textframe#horizontalalignment)|Representa o alinhamento horizontal do quadro de texto. Confira Excel.ShapeTextHorizontalAlignment para obter detalhes.|
||[horizontalOverflow](/javascript/api/excel/excel.textframe#horizontaloverflow)|Representa o comportamento de excedente horizontal do quadro de texto. Confira Excel.ShapeTextHorizontalOverflow para obter detalhes.|
||[leftMargin](/javascript/api/excel/excel.textframe#leftmargin)|Representa margem esquerda, em pontos, do quadro de texto.|
||[orientation](/javascript/api/excel/excel.textframe#orientation)|Representa a orientação do texto do quadro de texto. Confira Excel.ShapeTextOrientation para obter detalhes.|
||[readingOrder](/javascript/api/excel/excel.textframe#readingorder)|Representa a ordem de leitura do quadro de texto, da direita para a esquerda ou da direita para a esquerda. Confira Excel.ShapeTextReadingOrder para obter detalhes.|
||[hasText](/javascript/api/excel/excel.textframe#hastext)|Especifica se o quadro de texto contém texto.|
||[textRange](/javascript/api/excel/excel.textframe#textrange)|Representa o texto que está anexado a uma forma, bem como propriedades e métodos para manipular o texto. Confira Excel.TextRange para obter detalhes.|
||[rightMargin](/javascript/api/excel/excel.textframe#rightmargin)|Representa margem direita, em pontos, do quadro de texto.|
||[Set (Propriedades: Excel. TextFrame)](/javascript/api/excel/excel.textframe#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. TextFrameUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.textframe#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[topMargin](/javascript/api/excel/excel.textframe#topmargin)|Representa margem superior, em pontos, do quadro de texto.|
||[verticalAlignment](/javascript/api/excel/excel.textframe#verticalalignment)|Representa o alinhamento vertical do quadro de texto. Confira Excel.ShapeTextVerticalAlignment para obter detalhes.|
||[verticalOverflow](/javascript/api/excel/excel.textframe#verticaloverflow)|Representa o comportamento de excedente vertical do quadro de texto. Confira Excel.ShapeTextVerticalOverflow para obter detalhes.|
|[TextFrameData](/javascript/api/excel/excel.textframedata)|[autoSizeSetting](/javascript/api/excel/excel.textframedata#autosizesetting)|Obtém ou define as configurações de dimensionamento automático para o quadro de texto. Um quadro de texto pode ser configurado para ajustar automaticamente o texto ao quadro de texto, para ajustar automaticamente o quadro do texto ao texto ou não executar qualquer dimensionamento automático.|
||[bottomMargin](/javascript/api/excel/excel.textframedata#bottommargin)|Representa margem inferior, em pontos, do quadro de texto.|
||[hasText](/javascript/api/excel/excel.textframedata#hastext)|Especifica se o quadro de texto contém texto.|
||[horizontalAlignment](/javascript/api/excel/excel.textframedata#horizontalalignment)|Representa o alinhamento horizontal do quadro de texto. Confira Excel.ShapeTextHorizontalAlignment para obter detalhes.|
||[horizontalOverflow](/javascript/api/excel/excel.textframedata#horizontaloverflow)|Representa o comportamento de excedente horizontal do quadro de texto. Confira Excel.ShapeTextHorizontalOverflow para obter detalhes.|
||[leftMargin](/javascript/api/excel/excel.textframedata#leftmargin)|Representa margem esquerda, em pontos, do quadro de texto.|
||[orientation](/javascript/api/excel/excel.textframedata#orientation)|Representa a orientação do texto do quadro de texto. Confira Excel.ShapeTextOrientation para obter detalhes.|
||[readingOrder](/javascript/api/excel/excel.textframedata#readingorder)|Representa a ordem de leitura do quadro de texto, da direita para a esquerda ou da direita para a esquerda. Confira Excel.ShapeTextReadingOrder para obter detalhes.|
||[rightMargin](/javascript/api/excel/excel.textframedata#rightmargin)|Representa margem direita, em pontos, do quadro de texto.|
||[topMargin](/javascript/api/excel/excel.textframedata#topmargin)|Representa margem superior, em pontos, do quadro de texto.|
||[verticalAlignment](/javascript/api/excel/excel.textframedata#verticalalignment)|Representa o alinhamento vertical do quadro de texto. Confira Excel.ShapeTextVerticalAlignment para obter detalhes.|
||[verticalOverflow](/javascript/api/excel/excel.textframedata#verticaloverflow)|Representa o comportamento de excedente vertical do quadro de texto. Confira Excel.ShapeTextVerticalOverflow para obter detalhes.|
|[TextFrameLoadOptions](/javascript/api/excel/excel.textframeloadoptions)|[$all](/javascript/api/excel/excel.textframeloadoptions#$all)||
||[autoSizeSetting](/javascript/api/excel/excel.textframeloadoptions#autosizesetting)|Obtém ou define as configurações de dimensionamento automático para o quadro de texto. Um quadro de texto pode ser configurado para ajustar automaticamente o texto ao quadro de texto, para ajustar automaticamente o quadro do texto ao texto ou não executar qualquer dimensionamento automático.|
||[bottomMargin](/javascript/api/excel/excel.textframeloadoptions#bottommargin)|Representa margem inferior, em pontos, do quadro de texto.|
||[hasText](/javascript/api/excel/excel.textframeloadoptions#hastext)|Especifica se o quadro de texto contém texto.|
||[horizontalAlignment](/javascript/api/excel/excel.textframeloadoptions#horizontalalignment)|Representa o alinhamento horizontal do quadro de texto. Confira Excel.ShapeTextHorizontalAlignment para obter detalhes.|
||[horizontalOverflow](/javascript/api/excel/excel.textframeloadoptions#horizontaloverflow)|Representa o comportamento de excedente horizontal do quadro de texto. Confira Excel.ShapeTextHorizontalOverflow para obter detalhes.|
||[leftMargin](/javascript/api/excel/excel.textframeloadoptions#leftmargin)|Representa margem esquerda, em pontos, do quadro de texto.|
||[orientation](/javascript/api/excel/excel.textframeloadoptions#orientation)|Representa a orientação do texto do quadro de texto. Confira Excel.ShapeTextOrientation para obter detalhes.|
||[readingOrder](/javascript/api/excel/excel.textframeloadoptions#readingorder)|Representa a ordem de leitura do quadro de texto, da direita para a esquerda ou da direita para a esquerda. Confira Excel.ShapeTextReadingOrder para obter detalhes.|
||[rightMargin](/javascript/api/excel/excel.textframeloadoptions#rightmargin)|Representa margem direita, em pontos, do quadro de texto.|
||[textRange](/javascript/api/excel/excel.textframeloadoptions#textrange)|Representa o texto que está anexado a uma forma, bem como propriedades e métodos para manipular o texto. Confira Excel.TextRange para obter detalhes.|
||[topMargin](/javascript/api/excel/excel.textframeloadoptions#topmargin)|Representa margem superior, em pontos, do quadro de texto.|
||[verticalAlignment](/javascript/api/excel/excel.textframeloadoptions#verticalalignment)|Representa o alinhamento vertical do quadro de texto. Confira Excel.ShapeTextVerticalAlignment para obter detalhes.|
||[verticalOverflow](/javascript/api/excel/excel.textframeloadoptions#verticaloverflow)|Representa o comportamento de excedente vertical do quadro de texto. Confira Excel.ShapeTextVerticalOverflow para obter detalhes.|
|[TextFrameUpdateData](/javascript/api/excel/excel.textframeupdatedata)|[autoSizeSetting](/javascript/api/excel/excel.textframeupdatedata#autosizesetting)|Obtém ou define as configurações de dimensionamento automático para o quadro de texto. Um quadro de texto pode ser configurado para ajustar automaticamente o texto ao quadro de texto, para ajustar automaticamente o quadro do texto ao texto ou não executar qualquer dimensionamento automático.|
||[bottomMargin](/javascript/api/excel/excel.textframeupdatedata#bottommargin)|Representa margem inferior, em pontos, do quadro de texto.|
||[horizontalAlignment](/javascript/api/excel/excel.textframeupdatedata#horizontalalignment)|Representa o alinhamento horizontal do quadro de texto. Confira Excel.ShapeTextHorizontalAlignment para obter detalhes.|
||[horizontalOverflow](/javascript/api/excel/excel.textframeupdatedata#horizontaloverflow)|Representa o comportamento de excedente horizontal do quadro de texto. Confira Excel.ShapeTextHorizontalOverflow para obter detalhes.|
||[leftMargin](/javascript/api/excel/excel.textframeupdatedata#leftmargin)|Representa margem esquerda, em pontos, do quadro de texto.|
||[orientation](/javascript/api/excel/excel.textframeupdatedata#orientation)|Representa a orientação do texto do quadro de texto. Confira Excel.ShapeTextOrientation para obter detalhes.|
||[readingOrder](/javascript/api/excel/excel.textframeupdatedata#readingorder)|Representa a ordem de leitura do quadro de texto, da direita para a esquerda ou da direita para a esquerda. Confira Excel.ShapeTextReadingOrder para obter detalhes.|
||[rightMargin](/javascript/api/excel/excel.textframeupdatedata#rightmargin)|Representa margem direita, em pontos, do quadro de texto.|
||[topMargin](/javascript/api/excel/excel.textframeupdatedata#topmargin)|Representa margem superior, em pontos, do quadro de texto.|
||[verticalAlignment](/javascript/api/excel/excel.textframeupdatedata#verticalalignment)|Representa o alinhamento vertical do quadro de texto. Confira Excel.ShapeTextVerticalAlignment para obter detalhes.|
||[verticalOverflow](/javascript/api/excel/excel.textframeupdatedata#verticaloverflow)|Representa o comportamento de excedente vertical do quadro de texto. Confira Excel.ShapeTextVerticalOverflow para obter detalhes.|
|[TextRange](/javascript/api/excel/excel.textrange)|[getSubstring(start: number, length?: number)](/javascript/api/excel/excel.textrange#getsubstring-start--length-)|Retorna um objeto TextRange para a subcadeia de caracteres no intervalo especificado.|
||[font](/javascript/api/excel/excel.textrange#font)|Retorna um objeto ShapeFont que representa os atributos de fonte do intervalo de texto. Somente leitura.|
||[Set (Propriedades: Excel. TextRange)](/javascript/api/excel/excel.textrange#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. TextRangeUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.textrange#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[text](/javascript/api/excel/excel.textrange#text)|Representa o conteúdo de texto sem formatação do intervalo de texto.|
|[TextRangeData](/javascript/api/excel/excel.textrangedata)|[font](/javascript/api/excel/excel.textrangedata#font)|Retorna um objeto ShapeFont que representa os atributos de fonte do intervalo de texto. Somente leitura.|
||[text](/javascript/api/excel/excel.textrangedata#text)|Representa o conteúdo de texto sem formatação do intervalo de texto.|
|[TextRangeLoadOptions](/javascript/api/excel/excel.textrangeloadoptions)|[$all](/javascript/api/excel/excel.textrangeloadoptions#$all)||
||[font](/javascript/api/excel/excel.textrangeloadoptions#font)|Retorna um objeto ShapeFont que representa os atributos de fonte do intervalo de texto.|
||[text](/javascript/api/excel/excel.textrangeloadoptions#text)|Representa o conteúdo de texto sem formatação do intervalo de texto.|
|[TextRangeUpdateData](/javascript/api/excel/excel.textrangeupdatedata)|[font](/javascript/api/excel/excel.textrangeupdatedata#font)|Retorna um objeto ShapeFont que representa os atributos de fonte do intervalo de texto.|
||[text](/javascript/api/excel/excel.textrangeupdatedata#text)|Representa o conteúdo de texto sem formatação do intervalo de texto.|
|[Workbook](/javascript/api/excel/excel.workbook)|[chartDataPointTrack](/javascript/api/excel/excel.workbook#chartdatapointtrack)|True se todos os gráficos na pasta de trabalho estiverem rastreando os pontos de dados reais aos quais eles estão anexados.|
||[getActiveChart()](/javascript/api/excel/excel.workbook#getactivechart--)|Obtém o gráfico ativo no momento na pasta de trabalho. Se não houver um gráfico ativo, será lançada uma exceção quando essa instrução for invocada|
||[getActiveChartOrNullObject()](/javascript/api/excel/excel.workbook#getactivechartornullobject--)|Obtém o gráfico ativo no momento na pasta de trabalho. Se não houver um gráfico ativo, um objeto null será retornado|
||[getIsActiveCollabSession()](/javascript/api/excel/excel.workbook#getisactivecollabsession--)|True se a pasta de trabalho estiver sendo editada por vários usuários (coautoria).|
||[getSelectedRanges()](/javascript/api/excel/excel.workbook#getselectedranges--)|Obtém um ou mais intervalos atualmente selecionados da pasta de trabalho. Ao contrário de getSelectedRange(), esse método retorna um objeto RangeAreas que representa todos os intervalos selecionados.|
||[isDirty](/javascript/api/excel/excel.workbook#isdirty)|Especifica se as alterações foram feitas ou não desde que a pasta de trabalho foi salva pela última vez.|
||[autoSave](/javascript/api/excel/excel.workbook#autosave)|Especifica se a pasta de trabalho está ou não no modo de salvamento automático. Somente Leitura.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbook#calculationengineversion)|Retorna um número sobre a versão do Mecanismo de Cálculo do Excel. Somente Leitura.|
||[onAutoSaveSettingChanged](/javascript/api/excel/excel.workbook#onautosavesettingchanged)|Ocorre quando a configuração Salvamento automático é alterada na pasta de trabalho.|
||[previouslySaved](/javascript/api/excel/excel.workbook#previouslysaved)|Especifica se a pasta de trabalho já foi salva localmente ou online. Somente Leitura.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbook#useprecisionasdisplayed)|True se os cálculos dessa pasta de trabalho forem efetuados usando apenas a precisão dos números conforme forem exibidos.|
|[WorkbookAutoSaveSettingChangedEventArgs](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs)|[tipo](/javascript/api/excel/excel.workbookautosavesettingchangedeventargs#type)|Representa o tipo do evento. Para saber detalhes, confira Excel.EventType.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[autoSave](/javascript/api/excel/excel.workbookdata#autosave)|Especifica se a pasta de trabalho está ou não no modo de salvamento automático. Somente Leitura.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbookdata#calculationengineversion)|Retorna um número sobre a versão do Mecanismo de Cálculo do Excel. Somente Leitura.|
||[chartDataPointTrack](/javascript/api/excel/excel.workbookdata#chartdatapointtrack)|True se todos os gráficos na pasta de trabalho estiverem rastreando os pontos de dados reais aos quais eles estão anexados.|
||[isDirty](/javascript/api/excel/excel.workbookdata#isdirty)|Especifica se as alterações foram feitas ou não desde que a pasta de trabalho foi salva pela última vez.|
||[previouslySaved](/javascript/api/excel/excel.workbookdata#previouslysaved)|Especifica se a pasta de trabalho já foi salva localmente ou online. Somente Leitura.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookdata#useprecisionasdisplayed)|True se os cálculos dessa pasta de trabalho forem efetuados usando apenas a precisão dos números conforme forem exibidos.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[autoSave](/javascript/api/excel/excel.workbookloadoptions#autosave)|Especifica se a pasta de trabalho está ou não no modo de salvamento automático. Somente Leitura.|
||[calculationEngineVersion](/javascript/api/excel/excel.workbookloadoptions#calculationengineversion)|Retorna um número sobre a versão do Mecanismo de Cálculo do Excel. Somente Leitura.|
||[chartDataPointTrack](/javascript/api/excel/excel.workbookloadoptions#chartdatapointtrack)|True se todos os gráficos na pasta de trabalho estiverem rastreando os pontos de dados reais aos quais eles estão anexados.|
||[isDirty](/javascript/api/excel/excel.workbookloadoptions#isdirty)|Especifica se as alterações foram feitas ou não desde que a pasta de trabalho foi salva pela última vez.|
||[previouslySaved](/javascript/api/excel/excel.workbookloadoptions#previouslysaved)|Especifica se a pasta de trabalho já foi salva localmente ou online. Somente Leitura.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookloadoptions#useprecisionasdisplayed)|True se os cálculos dessa pasta de trabalho forem efetuados usando apenas a precisão dos números conforme forem exibidos.|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[chartDataPointTrack](/javascript/api/excel/excel.workbookupdatedata#chartdatapointtrack)|True se todos os gráficos na pasta de trabalho estiverem rastreando os pontos de dados reais aos quais eles estão anexados.|
||[isDirty](/javascript/api/excel/excel.workbookupdatedata#isdirty)|Especifica se as alterações foram feitas ou não desde que a pasta de trabalho foi salva pela última vez.|
||[usePrecisionAsDisplayed](/javascript/api/excel/excel.workbookupdatedata#useprecisionasdisplayed)|True se os cálculos dessa pasta de trabalho forem efetuados usando apenas a precisão dos números conforme forem exibidos.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[enableCalculation](/javascript/api/excel/excel.worksheet#enablecalculation)|Obtém ou define a propriedade enableCalculation da planilha.|
||[findAll(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findall-text--criteria-)|Encontra todas as ocorrências de determinada cadeia de caracteres com base nos critérios especificados e as retorna como um objeto RangeAreas, compreendendo um ou mais intervalos retangulares.|
||[findAllOrNullObject(text: string, criteria: Excel.WorksheetSearchCriteria)](/javascript/api/excel/excel.worksheet#findallornullobject-text--criteria-)|Encontra todas as ocorrências de determinada cadeia de caracteres com base nos critérios especificados e as retorna como um objeto RangeAreas, compreendendo um ou mais intervalos retangulares.|
||[getRanges(address?: string)](/javascript/api/excel/excel.worksheet#getranges-address-)|Obtém o objeto RangeAreas que representa um ou mais blocos de intervalos retangulares especificados pelo endereço ou nome.|
||[autoFilter](/javascript/api/excel/excel.worksheet#autofilter)|Representa o objeto AutoFilter da planilha. Somente Leitura.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheet#horizontalpagebreaks)|Obtém a coleção de quebra de página horizontal da planilha. Esta coleção contém apenas quebras de página manuais.|
||[onFormatChanged](/javascript/api/excel/excel.worksheet#onformatchanged)|Ocorre quando o formato é alterado em uma planilha específica.|
||[pageLayout](/javascript/api/excel/excel.worksheet#pagelayout)|Obtém o objeto PageLayout da planilha.|
||[shapes](/javascript/api/excel/excel.worksheet#shapes)|Retorna a coleção de todos os objetos Shape na planilha. Somente leitura.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheet#verticalpagebreaks)|Obtém a coleção de quebra de página vertical da planilha. Esta coleção contém apenas quebras de página manuais.|
||[replaceAll(text: string, replacement: string, criteria: Excel.ReplaceCriteria)](/javascript/api/excel/excel.worksheet#replaceall-text--replacement--criteria-)|Localiza e substitui a cadeia de caracteres fornecida com base nos critérios especificados na planilha atual.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[detalhes](/javascript/api/excel/excel.worksheetchangedeventargs#details)|Representa informações sobre os detalhes da alteração|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onChanged](/javascript/api/excel/excel.worksheetcollection#onchanged)|Ocorre quando uma planilha da pasta de trabalho é alterada.|
||[onFormatChanged](/javascript/api/excel/excel.worksheetcollection#onformatchanged)|Ocorre quando uma planilha na pasta de trabalho tem o formato alterado.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheetcollection#onselectionchanged)|Ocorre quando a seleção é alterada em uma planilha.|
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[autoFilter](/javascript/api/excel/excel.worksheetcollectionloadoptions#autofilter)|Para cada ITEM na coleção: representa o objeto AutoFilter da planilha.|
||[enableCalculation](/javascript/api/excel/excel.worksheetcollectionloadoptions#enablecalculation)|Para cada ITEM na coleção: Obtém ou define a propriedade enableCalculation da planilha.|
||[pageLayout](/javascript/api/excel/excel.worksheetcollectionloadoptions#pagelayout)|Para cada ITEM na coleção: Obtém o objeto PageLayout da planilha.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[autoFilter](/javascript/api/excel/excel.worksheetdata#autofilter)|Representa o objeto AutoFilter da planilha. Somente Leitura.|
||[enableCalculation](/javascript/api/excel/excel.worksheetdata#enablecalculation)|Obtém ou define a propriedade enableCalculation da planilha.|
||[horizontalPageBreaks](/javascript/api/excel/excel.worksheetdata#horizontalpagebreaks)|Obtém a coleção de quebra de página horizontal da planilha. Esta coleção contém apenas quebras de página manuais.|
||[pageLayout](/javascript/api/excel/excel.worksheetdata#pagelayout)|Obtém o objeto PageLayout da planilha.|
||[shapes](/javascript/api/excel/excel.worksheetdata#shapes)|Retorna a coleção de todos os objetos Shape na planilha. Somente leitura.|
||[verticalPageBreaks](/javascript/api/excel/excel.worksheetdata#verticalpagebreaks)|Obtém a coleção de quebra de página vertical da planilha. Esta coleção contém apenas quebras de página manuais.|
|[WorksheetFormatChangedEventArgs](/javascript/api/excel/excel.worksheetformatchangedeventargs)|[address](/javascript/api/excel/excel.worksheetformatchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrange-ctx-)|Obtém o intervalo que representa a área alterada de uma planilha específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetformatchangedeventargs#getrangeornullobject-ctx-)|Obtém o intervalo que representa a área alterada de uma planilha específica. Pode retornar o objeto null.|
||[source](/javascript/api/excel/excel.worksheetformatchangedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.worksheetformatchangedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetformatchangedeventargs#worksheetid)|Obtém o id da planilha na qual os dados são alterados.|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[autoFilter](/javascript/api/excel/excel.worksheetloadoptions#autofilter)|Representa o objeto AutoFilter da planilha.|
||[enableCalculation](/javascript/api/excel/excel.worksheetloadoptions#enablecalculation)|Obtém ou define a propriedade enableCalculation da planilha.|
||[pageLayout](/javascript/api/excel/excel.worksheetloadoptions#pagelayout)|Obtém o objeto PageLayout da planilha.|
|[WorksheetSearchCriteria](/javascript/api/excel/excel.worksheetsearchcriteria)|[completeMatch](/javascript/api/excel/excel.worksheetsearchcriteria#completematch)|Especifica se a correspondência deve ser completa ou parcial. Uma correspondência completa corresponde a todo o conteúdo da célula. O padrão é false (parcial).|
||[matchCase](/javascript/api/excel/excel.worksheetsearchcriteria#matchcase)|Especifica se a correspondência diferencia maiúsculas de minúsculas. O padrão é false (não diferencia maiúsculas de minúsculas).|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[enableCalculation](/javascript/api/excel/excel.worksheetupdatedata#enablecalculation)|Obtém ou define a propriedade enableCalculation da planilha.|
||[pageLayout](/javascript/api/excel/excel.worksheetupdatedata#pagelayout)|Obtém o objeto PageLayout da planilha.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do Excel](/javascript/api/excel)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
