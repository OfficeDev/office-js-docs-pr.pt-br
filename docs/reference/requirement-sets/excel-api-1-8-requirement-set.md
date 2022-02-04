---
title: Excel conjunto de requisitos da API JavaScript 1.8
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.8.
ms.date: 03/19/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-18"></a>Novidades na API JavaScript 1.8 Excel JavaScript

O conjunto de requisitos 1.8 da API JavaScript do Excel inclui APIs para tabelas dinâmicas, validação de dados, gráficos, eventos de gráficos, opções de desempenho e criação de pasta de trabalho.

## <a name="pivottable"></a>Tabela Dinâmica

Onda 2 das APIs de Tabela Dinâmica permite que os suplementos definam as hierarquias de uma Tabela Dinâmica. Agora você pode controlar os dados e como eles são agregados. Nosso [Artigo de Tabela Dinâmica](../../excel/excel-add-ins-pivottables.md) tem mais informações sobre a nova funcionalidade de tabela dinâmica.

## <a name="data-validation"></a>Validação de Dados

A validação de dados permite controlar o que um usuário digita em uma planilha. Você pode limitar as células a conjuntos de respostas predefinidos ou fornecer avisos pop-up sobre entradas indesejadas. Saiba mais sobre [adicionar a validação de dados para intervalos](../../excel/excel-add-ins-data-validation.md) hoje.

## <a name="charts"></a>Gráficos

Outra rodada de APIs de gráficos traz um controle programático ainda maior sobre os elementos do gráfico. Agora você tem maior acesso à legenda, eixos, linha de tendência e área de plotagem.

## <a name="events"></a>Eventos

Mais [eventos](../../excel/excel-add-ins-events.md) foram adicionados para os gráficos. Faça o seu suplemento reagir aos usuários interagindo com o gráfico. Você também pode [alternar eventos](../../excel/performance.md#enable-and-disable-events) disparados em toda a pasta de trabalho.

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.8. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.8 ou anterior, consulte Excel APIs no conjunto de requisitos [1.8](/javascript/api/excel?view=excel-js-1.8&preserve-view=true) ou anterior.

| Classe | Campos | Descrição |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[formula1](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-formula1-member)|Especifica o operador à direita quando a propriedade operator é definida como um operador binário como GreaterThan (o operand à esquerda é o valor que o usuário tenta inserir na célula).|
||[formula2](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-formula2-member)|Com os operadores ternários Between e NotBetween, especifica o operand superior ligado.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#excel-excel-basicdatavalidation-operator-member)|O operador a ser usado para validar os dados.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#excel-excel-chart-categorylabellevel-member)|Especifica uma constante de enumeração de nível de rótulo de categoria de gráfico, referindo-se ao nível dos rótulos de categoria de origem.|
||[displayBlanksAs](/javascript/api/excel/excel.chart#excel-excel-chart-displayblanksas-member)|Especifica a maneira como as células em branco são plotadas em um gráfico.|
||[onActivated](/javascript/api/excel/excel.chart#excel-excel-chart-onactivated-member)|Ocorre quando o gráfico é ativado.|
||[onDeactivated](/javascript/api/excel/excel.chart#excel-excel-chart-ondeactivated-member)|Ocorre quando o gráfico é desativado.|
||[plotArea](/javascript/api/excel/excel.chart#excel-excel-chart-plotarea-member)|Representa a área de plotagem do gráfico.|
||[plotBy](/javascript/api/excel/excel.chart#excel-excel-chart-plotby-member)|Especifica a forma como as colunas ou linhas são usadas como série de dados no gráfico.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#excel-excel-chart-plotvisibleonly-member)|Verdadeiro se apenas as células visíveis forem plotadas.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#excel-excel-chart-seriesnamelevel-member)|Especifica uma constante de enumeração de nível de nome de série de gráfico, referindo-se ao nível dos nomes da série de origem.|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#excel-excel-chart-showdatalabelsovermaximum-member)|Especifica se os rótulos de dados serão indicados quando o valor for maior do que o valor máximo no eixo do valor.|
||[style](/javascript/api/excel/excel.chart#excel-excel-chart-style-member)|Especifica o estilo do gráfico para o gráfico.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-chartid-member)|Obtém a ID do gráfico ativado.|
||[tipo](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#excel-excel-chartactivatedeventargs-worksheetid-member)|Obtém a ID da planilha na qual o gráfico é ativado.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartId](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-chartid-member)|Obtém a ID do gráfico adicionado à planilha.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-source-member)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#excel-excel-chartaddedeventargs-worksheetid-member)|Obtém a ID da planilha na qual o gráfico é adicionado.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[alignment](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-alignment-member)|Especifica o alinhamento do rótulo de escala do eixo especificado.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-isbetweencategories-member)|Especifica se o eixo do valor cruza o eixo de categoria entre categorias.|
||[multiLevel](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-multilevel-member)|Especifica se um eixo é multinível.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-numberformat-member)|Especifica o código de formato do rótulo de escala do eixo.|
||[offset](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-offset-member)|Especifica a distância entre os níveis de rótulos e a distância entre o primeiro nível e a linha do eixo.|
||[position](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-position-member)|Especifica a posição do eixo especificada onde o outro eixo cruza.|
||[positionAt](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-positionat-member)|Especifica a posição do eixo onde o outro eixo cruza.|
||[setPositionAt(value: number)](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-setpositionat-member(1))|Define a posição do eixo especificada onde o outro eixo cruza.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#excel-excel-chartaxis-textorientation-member)|Especifica o ângulo para o qual o texto é orientado para o rótulo de escala do eixo do gráfico.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#excel-excel-chartaxisformat-fill-member)|Especifica a formatação de preenchimento do gráfico.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-setformula-member(1))|Um valor de cadeia de caracteres que representa a fórmula do título do eixo do gráfico usando a notação no estilo A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[borda](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-border-member)|Especifica o formato de borda do título do eixo do gráfico, que inclui cor, estilo de linha e peso.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#excel-excel-chartaxistitleformat-fill-member)|Especifica a formatação de preenchimento do título do eixo do gráfico.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#excel-excel-chartborder-clear-member(1))|Limpa a formatação da borda de um elemento do gráfico.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onactivated-member)|Ocorre quando um gráfico é ativado.|
||[onAdded](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-onadded-member)|Ocorre quando um novo gráfico é adicionado à planilha.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeactivated-member)|Ocorre quando um gráfico é desativado.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-ondeleted-member)|Ocorre quando um gráfico é excluído.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[autoText](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-autotext-member)|Especifica se o rótulo de dados gera automaticamente o texto apropriado com base no contexto.|
||[format](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-format-member)|Representa o formato do rótulo de dados do gráfico.|
||[formula](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-formula-member)|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de dados usando a notação no estilo A1.|
||[height](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-height-member)|Retorna a altura, em pontos, do rótulo de dados do gráfico.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-horizontalalignment-member)|Representa o alinhamento horizontal de rótulo de dados do gráfico.|
||[left](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-left-member)|Representa a distância, em pontos, da borda esquerda do rótulo de dados do gráfico até a borda esquerda da área do gráfico.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-numberformat-member)|Valor de cadeia de caracteres que representa o código do formato do rótulo de dados.|
||[text](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-text-member)|Cadeia de caracteres que representa o texto do rótulo de dados em um gráfico.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-textorientation-member)|Representa o ângulo para o qual o texto é orientado para o rótulo de dados do gráfico.|
||[top](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-top-member)|Representa a distância, em pontos, da borda superior do rótulo de dados do gráfico até a borda superior da área do gráfico.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-verticalalignment-member)|Representa o alinhamento vertical do rótulo de dados do gráfico.|
||[width](/javascript/api/excel/excel.chartdatalabel#excel-excel-chartdatalabel-width-member)|Retorna a largura, em pontos, do rótulo de dados do gráfico.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[borda](/javascript/api/excel/excel.chartdatalabelformat#excel-excel-chartdatalabelformat-border-member)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[autoText](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-autotext-member)|Especifica se os rótulos de dados geram automaticamente o texto apropriado com base no contexto.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-horizontalalignment-member)|Especifica o alinhamento horizontal para o rótulo de dados do gráfico.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-numberformat-member)|Especifica o código de formato para rótulos de dados.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-textorientation-member)|Representa o ângulo para o qual o texto é orientado para rótulos de dados.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#excel-excel-chartdatalabels-verticalalignment-member)|Representa o alinhamento vertical do rótulo de dados do gráfico.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartId](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-chartid-member)|Obtém a ID do gráfico que é desativado.|
||[tipo](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#excel-excel-chartdeactivatedeventargs-worksheetid-member)|Obtém a ID da planilha na qual o gráfico é desativado.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartId](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-chartid-member)|Obtém a ID do gráfico excluído da planilha.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-source-member)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#excel-excel-chartdeletedeventargs-worksheetid-member)|Obtém a ID da planilha na qual o gráfico é excluído.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-height-member)|Especifica a altura da entrada da legenda na legenda do gráfico.|
||[índice](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-index-member)|Especifica o índice da entrada da legenda na legenda do gráfico.|
||[left](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-left-member)|Especifica o valor esquerdo de uma entrada de legenda de gráfico.|
||[top](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-top-member)|Especifica a parte superior de uma entrada de legenda de gráfico.|
||[width](/javascript/api/excel/excel.chartlegendentry#excel-excel-chartlegendentry-width-member)|Representa a largura da entrada da legenda no gráfico Legend.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[borda](/javascript/api/excel/excel.chartlegendformat#excel-excel-chartlegendformat-border-member)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[format](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-format-member)|Especifica a formatação de uma área de plotagem de gráfico.|
||[height](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-height-member)|Especifica o valor de altura de uma área de plotagem.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insideheight-member)|Especifica o valor de altura interna de uma área de plotagem.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insideleft-member)|Especifica o valor interno esquerdo de uma área de plotagem.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insidetop-member)|Especifica o valor superior interno de uma área de plotagem.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-insidewidth-member)|Especifica o valor de largura interna de uma área de plotagem.|
||[left](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-left-member)|Especifica o valor esquerdo de uma área de plotagem.|
||[position](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-position-member)|Especifica a posição de uma área de plotagem.|
||[top](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-top-member)|Especifica o valor superior de uma área de plotagem.|
||[width](/javascript/api/excel/excel.chartplotarea#excel-excel-chartplotarea-width-member)|Especifica o valor de largura de uma área de plotagem.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[borda](/javascript/api/excel/excel.chartplotareaformat#excel-excel-chartplotareaformat-border-member)|Especifica os atributos de borda de uma área de plotagem de gráfico.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#excel-excel-chartplotareaformat-fill-member)|Especifica o formato de preenchimento de um objeto, que inclui informações de formatação em segundo plano.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-axisgroup-member)|Especifica o grupo da série especificada.|
||[dataLabels](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-datalabels-member)|Representa uma coleção de todos os rótulos de dados na série.|
||[explosion](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-explosion-member)|Especifica o valor de explosão para uma fatia de gráfico de pizza ou gráfico de rosca.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-firstsliceangle-member)|Especifica o ângulo da primeira fatia do gráfico de pizza ou do gráfico de rosca, em graus (no sentido horário da vertical).|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-invertifnegative-member)|True se Excel inverte o padrão no item quando corresponde a um número negativo.|
||[sobreposição](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-overlap-member)|Especifica como barras e colunas são posicionadas.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-secondplotsize-member)|Especifica o tamanho da seção secundária de um gráfico de pizza de pizza ou um gráfico de barras de pizza, como uma porcentagem do tamanho da pizza primária.|
||[splitType](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-splittype-member)|Especifica a maneira como as duas seções de um gráfico de pizza de pizza ou um gráfico de barras de pizza são divididas.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-varybycategories-member)|True se Excel atribuir uma cor ou padrão diferente a cada marcador de dados.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-backwardperiod-member)|Representa o número de períodos que a linha de tendência se estende para trás.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-forwardperiod-member)|Representa o número de períodos que a linha de tendência se estende para frente.|
||[label](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-label-member)|Representa o rótulo de linha de tendência um gráfico.|
||[showEquation](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-showequation-member)|Verdadeiro se a equação da linha de tendência for exibida no gráfico.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#excel-excel-charttrendline-showrsquared-member)|True se o valor r-quadrado da linha de tendência for exibido no gráfico.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[autoText](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-autotext-member)|Especifica se o rótulo da linha de tendência gera automaticamente o texto apropriado com base no contexto.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-format-member)|O formato do rótulo de linha de tendência do gráfico.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-formula-member)|Valor de cadeia de caracteres que representa a fórmula do rótulo de linha de tendência do gráfico usando notação de estilo A1.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-height-member)|Retorna a altura, em pontos, do rótulo de linha de tendência do gráfico.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-horizontalalignment-member)|Representa o alinhamento horizontal do rótulo de linha de tendência do gráfico.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-left-member)|Representa a distância, em pontos, da borda esquerda do rótulo de linha de tendência do gráfico até a borda esquerda da área do gráfico.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-numberformat-member)|Valor de cadeia de caracteres que representa o código de formato do rótulo de linha de tendência.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-text-member)|Cadeia de caracteres que representa o texto do rótulo em um gráfico de linha de tendência.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-textorientation-member)|Representa o ângulo para o qual o texto é orientado para o rótulo de linha de tendência do gráfico.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-top-member)|Representa a distância, em pontos, da borda superior do rótulo de linha de tendência do gráfico até a parte superior da área do gráfico.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-verticalalignment-member)|Representa o alinhamento vertical do rótulo de linha de tendência do gráfico.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#excel-excel-charttrendlinelabel-width-member)|Retorna a largura, em pontos, do rótulo de linha de tendência do gráfico.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[borda](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-border-member)|Especifica o formato de borda, que inclui cor, estilo de linha e peso.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-fill-member)|Especifica o formato de preenchimento do rótulo de linha de tendência do gráfico atual.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#excel-excel-charttrendlinelabelformat-font-member)|Especifica os atributos de fonte (como nome da fonte, tamanho da fonte e cor) para um rótulo de linha de tendência de gráfico.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#excel-excel-customdatavalidation-formula-member)|Uma fórmula de validação de dados personalizados.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[campo](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-field-member)|Retorna PivotFields associados a DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-id-member)|ID do DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-name-member)|Nome da DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-numberformat-member)|Formato de número do DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-position-member)|Posição da DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-settodefault-member(1))|Redefina a DataPivotHierarchy para os valores padrão.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-showas-member)|Especifica se os dados devem ser mostrados como um cálculo de resumo específico.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#excel-excel-datapivothierarchy-summarizeby-member)|Especifica se todos os itens do DataPivotHierarchy são mostrados.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-add-member(1))|Adiciona o PivotHierarchy ao eixo atual.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getcount-member(1))|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getitem-member(1))|Obtém um DataPivotHierarchy pelo nome ou ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-getitemornullobject-member(1))|Obtém uma DataPivotHierarchy por nome.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[remove(DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#excel-excel-datapivothierarchycollection-remove-member(1))|Remove o PivotHierarchy do eixo atual.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-clear-member(1))|Desfazer a validação de dados do intervalo atual.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-erroralert-member)|Alerta de erro quando o usuário insere dados inválidos.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-ignoreblanks-member)|Especifica se a validação de dados será realizada em células em branco.|
||[prompt](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-prompt-member)|Prompt when users select a cell.|
||[rule](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-rule-member)|Regra de validação de dados que contém diferentes tipos de critérios de validação de dados.|
||[tipo](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-type-member)|Tipo de validação de dados, consulte `Excel.DataValidationType` para obter detalhes.|
||[valid](/javascript/api/excel/excel.datavalidation#excel-excel-datavalidation-valid-member)|Representa se todos os valores de célula são válidos de acordo com as regras de validação de dados.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-message-member)|Representa a mensagem de alerta de erro.|
||[showAlert](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-showalert-member)|Especifica se será exibida uma caixa de diálogo de alerta de erro quando um usuário inserir dados inválidos.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-style-member)|O tipo de alerta de validação de dados, consulte `Excel.DataValidationAlertStyle` para obter detalhes.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#excel-excel-datavalidationerroralert-title-member)|Representa o título da caixa de diálogo de alerta de erro.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-message-member)|Especifica a mensagem do prompt.|
||[showPrompt](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-showprompt-member)|Especifica se um prompt é mostrado quando um usuário seleciona uma célula com validação de dados.|
||[title](/javascript/api/excel/excel.datavalidationprompt#excel-excel-datavalidationprompt-title-member)|Especifica o título do prompt.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[custom](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-custom-member)|Critérios de validação de dados personalizados.|
||[data](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-date-member)|Critérios de validação de dados de data.|
||[decimal](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-decimal-member)|Critérios de validação de dados decimais.|
||[list](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-list-member)|Critérios de validação de dados da lista.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-textlength-member)|Critérios de validação de dados de comprimento de texto.|
||[time](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-time-member)|Critérios de validação de dados de tempo.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#excel-excel-datavalidationrule-wholenumber-member)|Critérios de validação de dados de número inteiro.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[formula1](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-formula1-member)|Especifica o operador à direita quando a propriedade operator é definida como um operador binário como GreaterThan (o operand à esquerda é o valor que o usuário tenta inserir na célula).|
||[formula2](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-formula2-member)|Com os operadores ternários Between e NotBetween, especifica o operand superior ligado.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#excel-excel-datetimedatavalidation-operator-member)|O operador a ser usado para validar os dados.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-enablemultiplefilteritems-member)|Determina se deseja permitir vários itens de filtro.|
||[campos](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-fields-member)|Retorna PivotFields associados a FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-id-member)|ID do FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-name-member)|Nome do FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-position-member)|Posição do FilterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#excel-excel-filterpivothierarchy-settodefault-member(1))|Redefina a FilterPivotHierarchy para os valores padrão.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-add-member(1))|Adiciona o PivotHierarchy ao eixo atual.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getcount-member(1))|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getitem-member(1))|Obtém um FilterPivotHierarchy pelo nome ou ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-getitemornullobject-member(1))|Obtém um FilterPivotHierarchy por nome.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[remove(filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#excel-excel-filterpivothierarchycollection-remove-member(1))|Remove o PivotHierarchy do eixo atual.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#excel-excel-listdatavalidation-incelldropdown-member)|Especifica se a lista deve ser exibida em um drop-down de célula.|
||[source](/javascript/api/excel/excel.listdatavalidation#excel-excel-listdatavalidation-source-member)|Fonte da lista de validação de dados|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[id](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-id-member)|ID do PivotField.|
||[items](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-items-member)|Retorna os PivotItems associados ao PivotField.|
||[name](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-name-member)|Nome do PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-showallitems-member)|Determina se deseja mostrar todos os itens de PivotField.|
||[sortByLabels(sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-sortbylabels-member(1))|Classifica o PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-subtotals-member)|Subtotais de PivotField.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getcount-member(1))|Obtém o número de campos de pivô na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getitem-member(1))|Obtém um PivotField pelo nome ou ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-getitemornullobject-member(1))|Obtém um PivotField pelo nome.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#excel-excel-pivotfieldcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[campos](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-fields-member)|Retorna PivotFields associados a PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-id-member)|ID do PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchy#excel-excel-pivothierarchy-name-member)|Nome do PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getcount-member(1))|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getitem-member(1))|Obtém um PivotHierarchy pelo nome ou ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-getitemornullobject-member(1))|Obtém o PivotHierarchy por nome.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#excel-excel-pivothierarchycollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[id](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-id-member)|ID do PivotItem.|
||[isExpanded](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-isexpanded-member)|Determina se o item está expandido para mostrar itens filho ou se ele está recolhido e os itens filho estão ocultos.|
||[name](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-name-member)|Nome do PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitem#excel-excel-pivotitem-visible-member)|Especifica se o PivotItem está visível.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getcount-member(1))|Obtém o número de PivotItems na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getitem-member(1))|Obtém um PivotItem pelo nome ou ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-getitemornullobject-member(1))|Obtém um PivotItem pelo nome.|
||[items](/javascript/api/excel/excel.pivotitemcollection#excel-excel-pivotitemcollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getcolumnlabelrange-member(1))|Retorna o intervalo onde residem os rótulos de coluna da Tabela Dinâmica.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getdatabodyrange-member(1))|Retorna o intervalo onde residem os valores de dados da tabela dinâmica.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getfilteraxisrange-member(1))|Retorna o intervalo de área de filtro da Tabela Dinâmica.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getrange-member(1))|Retorna o intervalo em que a Tabela Dinâmica existe, excluindo a área de filtro.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-getrowlabelrange-member(1))|Retorna o intervalo onde residem os rótulos de linha da Tabela Dinâmica.|
||[layoutType](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-layouttype-member)|Essa propriedade indica o PivotLayoutType de todos os campos da Tabela Dinâmica.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showcolumngrandtotals-member)|Especifica se o relatório de tabela dinâmica mostra totais grandes para colunas.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showrowgrandtotals-member)|Especifica se o relatório de tabela dinâmica mostra totais grandes para linhas.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-subtotallocation-member)|Essa propriedade indica o de `SubtotalLocationType` todos os campos na Tabela Dinâmica.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[columnHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-columnhierarchies-member)|As hierarquias de pivô da coluna da Tabela Dinâmica.|
||[dataHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-datahierarchies-member)|As hierarquias dinâmicas de dados da Tabela Dinâmica.|
||[delete()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-delete-member(1))|Exclui a Tabela Dinâmica.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-filterhierarchies-member)|As hierarquias de pivô do filtro da Tabela Dinâmica.|
||[hierarquias](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-hierarchies-member)|Hierarquias pivô da Tabela Dinâmica.|
||[layout](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-layout-member)|O PivotLayout descreve o layout e estrutura visual da Tabela Dinâmica.|
||[rowHierarchies](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-rowhierarchies-member)|As hierarquias de pivô de linha da Tabela Dinâmica.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[add(name: string, source: Range \| string \| Table, destination: Range \| string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-add-member(1))|Adicione uma Tabela Dinâmica com base nos dados de origem especificados e insira-a na célula superior esquerda do intervalo de destino.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#excel-excel-range-datavalidation-member)|Retorna um objeto de validação de dados.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[campos](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-fields-member)|Retorna PivotFields associados a RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-id-member)|ID do RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-name-member)|Nome da RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-position-member)|Posição da RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#excel-excel-rowcolumnpivothierarchy-settodefault-member(1))|Redefine o RowColumnPivotHierarchy para os valores padrão.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[add(pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-add-member(1))|Adiciona o PivotHierarchy ao eixo atual.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getcount-member(1))|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getitem-member(1))|Obtém um RowColumnPivotHierarchy pelo nome ou ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-getitemornullobject-member(1))|Obtém um RowColumnPivotHierarchy por nome.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-items-member)|Obtém os itens filhos carregados nesta coleção.|
||[remove(rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#excel-excel-rowcolumnpivothierarchycollection-remove-member(1))|Remove o PivotHierarchy do eixo atual.|
|[Tempo de execução](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#excel-excel-runtime-enableevents-member)|Alterne eventos JavaScript no painel de tarefas ou no complemento de conteúdo atual.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-basefield-member)|O PivotField para basear o `ShowAs` cálculo, se aplicável de acordo com o `ShowAsCalculation` tipo, senão `null`.|
||[baseItem](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-baseitem-member)|O item no qual basear o `ShowAs` cálculo, se aplicável de acordo com o `ShowAsCalculation` tipo, mais `null`.|
||[calculation](/javascript/api/excel/excel.showasrule#excel-excel-showasrule-calculation-member)|O `ShowAs` cálculo a ser usado para o PivotField.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#excel-excel-style-autoindent-member)|Especifica se o texto é recuado automaticamente quando o alinhamento de texto em uma célula é definido como distribuição igual.|
||[textOrientation](/javascript/api/excel/excel.style#excel-excel-style-textorientation-member)|A orientação de texto para o estilo.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-automatic-member)|Se `Automatic` estiver definido como `true`, todos os outros valores serão ignorados ao definir o `Subtotals`.|
||[average](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-average-member)||
||[Count](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-count-member)||
||[countNumbers](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-countnumbers-member)||
||[max](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-max-member)||
||[min](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-min-member)||
||[product](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-product-member)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-standarddeviation-member)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-standarddeviationp-member)||
||[sum](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-sum-member)||
||[variância](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-variance-member)||
||[varianceP](/javascript/api/excel/excel.subtotals#excel-excel-subtotals-variancep-member)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#excel-excel-table-legacyid-member)|Retorna uma ID numérica.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-getrange-member(1))|Obtém o intervalo que representa a área alterada de uma tabela em uma planilha específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#excel-excel-tablechangedeventargs-getrangeornullobject-member(1))|Obtém o intervalo que representa a área alterada de uma tabela em uma planilha específica.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#excel-excel-workbook-readonly-member)|Retorna `true` se a workbook estiver aberta no modo somente leitura.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)|||
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onCalculated](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncalculated-member)|Ocorre quando a planilha é calculada.|
||[showGridlines](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showgridlines-member)|Especifica se as linhas de grade estão visíveis para o usuário.|
||[showHeadings](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showheadings-member)|Especifica se os títulos estão visíveis para o usuário.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[tipo](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-type-member)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#excel-excel-worksheetcalculatedeventargs-worksheetid-member)|Obtém a ID da planilha na qual o cálculo ocorreu.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-getrange-member(1))|Obtém o intervalo que representa a área alterada de uma planilha específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-getrangeornullobject-member(1))|Obtém o intervalo que representa a área alterada de uma planilha específica.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onCalculated](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncalculated-member)|Ocorre quando qualquer planilha na pasta de trabalho é calculada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
