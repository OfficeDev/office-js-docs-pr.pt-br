---
title: Conjunto de requisitos de API JavaScript do Excel 1,8
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1,8.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6454a7429276148e36431bfaffdf929a19a36d76
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996204"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>O que há de novo na API JavaScript do Excel 1,8

O conjunto de requisitos 1.8 da API JavaScript do Excel inclui APIs para tabelas dinâmicas, validação de dados, gráficos, eventos de gráficos, opções de desempenho e criação de pasta de trabalho.

## <a name="pivottable"></a>Tabela Dinâmica

Onda 2 das APIs de Tabela Dinâmica permite que os suplementos definam as hierarquias de uma Tabela Dinâmica. Agora você pode controlar os dados e como eles são agregados. Nosso [Artigo de Tabela Dinâmica](../../excel/excel-add-ins-pivottables.md) tem mais informações sobre a nova funcionalidade de tabela dinâmica.

## <a name="data-validation"></a>Validação de Dados

A validação de dados permite controlar o que um usuário digita em uma planilha. Você pode limitar as células a conjuntos de respostas predefinidos ou fornecer avisos pop-up sobre entradas indesejadas. Saiba mais sobre [adicionar a validação de dados para intervalos](../../excel/excel-add-ins-data-validation.md) hoje.

## <a name="charts"></a>Gráficos

Outra rodada de APIs de gráficos traz um controle programático ainda maior sobre os elementos do gráfico. Agora você tem maior acesso à legenda, eixos, linha de tendência e área de plotagem.

## <a name="events"></a>Eventos

Mais [eventos](../../excel/excel-add-ins-events.md) foram adicionados para os gráficos. Faça o seu suplemento reagir aos usuários interagindo com o gráfico. Você também pode [alternar eventos](../../excel/performance.md#enable-and-disable-events) disparados em toda a pasta de trabalho.

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Excel 1,8. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Excel 1,8 ou anterior, confira [APIs do Excel no conjunto de requisitos 1,8 ou anterior](/javascript/api/excel?view=excel-js-1.8&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[Formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Especifica o operando direito quando a Propriedade Operator é definida como um operador binário como GreaterThan (o operando esquerdo é o valor que o usuário tenta inserir na célula).|
||[Formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|Com os operadores ternários between e não between, especifica o operando de limite superior.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#operator)|O operador a ser usado para validar os dados.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|Especifica uma constante de enumeração ChartCategoryLabelLevel que se refere a|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|Especifica a maneira como as células em branco são plotadas em um gráfico.|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|Especifica a forma como as colunas ou linhas são usadas como série de dados no gráfico.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotvisibleonly)|Verdadeiro se apenas as células visíveis forem plotadas. Falso se ambas as células visíveis e ocultas forem plotadas..|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|Ocorre quando o gráfico é ativado.|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|Ocorre quando o gráfico é desativado.|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|Representa a plotArea para o gráfico.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|Especifica uma constante de enumeração ChartSeriesNameLevel que se refere a|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|Especifica se os rótulos de dados devem ser mostrados quando o valor for maior do que o valor máximo no eixo dos valores.|
||[style](/javascript/api/excel/excel.chart#style)|Especifica o estilo de gráfico para o gráfico.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartid](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|Obtém o id do gráfico que está ativado.|
||[tipo](/javascript/api/excel/excel.chartactivatedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|Obtém o id da planilha na qual o gráfico é ativado.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartid](/javascript/api/excel/excel.chartaddedeventargs#chartid)|Obtém o id do gráfico que é adicionado à planilha.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.chartaddedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|Obtém o id da planilha na qual o gráfico é adicionado.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[Alignment](/javascript/api/excel/excel.chartaxis#alignment)|Especifica o alinhamento para o rótulo de escala do eixo especificado.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isbetweencategories)|Especifica se o eixo dos valores cruza o eixo das categorias entre categorias.|
||[Vários](/javascript/api/excel/excel.chartaxis#multilevel)|Especifica se um eixo é de vários níveis.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|Especifica o código de formatação para o rótulo de escala de eixo.|
||[partida](/javascript/api/excel/excel.chartaxis#offset)|Especifica a distância entre os níveis dos rótulos e a distância entre o primeiro nível e a linha do eixo.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Especifica a posição do eixo especificado onde o outro eixo cruza.|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionat)|Especifica a posição do eixo especificado onde o outro eixo cruza.|
||[setPositionAt (valor: número)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|Define a posição do eixo especificado onde o outro eixo cruza.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|Especifica o ângulo no qual o texto é orientado para o rótulo de escala do eixo do gráfico.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Especifica a formatação de preenchimento do gráfico.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setformula (fórmula: cadeia de caracteres)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|Um valor de cadeia de caracteres que representa a fórmula do título do eixo do gráfico usando a notação no estilo A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[borda](/javascript/api/excel/excel.chartaxistitleformat#border)|Especifica o formato de borda do título do eixo do gráfico, que inclui cores, LineStyle e Weight.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Especifica a formatação de preenchimento do título do eixo do gráfico.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Limpa a formatação da borda de um elemento do gráfico.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|Ocorre quando um gráfico é ativado.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|Ocorre quando um novo gráfico é adicionado à planilha.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|Ocorre quando um gráfico é desativado.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|Ocorre quando um gráfico é excluído.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[AutoTexto](/javascript/api/excel/excel.chartdatalabel#autotext)|Especifica se o rótulo de dados gerará automaticamente o texto apropriado com base no contexto.|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de dados usando a notação no estilo A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Representa o alinhamento horizontal de rótulo de dados do gráfico.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Representa a distância, em pontos, da borda esquerda do rótulo de dados do gráfico até a borda esquerda da área do gráfico.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|Valor de cadeia de caracteres que representa o código do formato do rótulo de dados.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Representa o formato do rótulo de dados do gráfico.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Retorna a altura, em pontos, do rótulo de dados do gráfico.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Retorna a largura, em pontos, do rótulo de dados do gráfico.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Cadeia de caracteres que representa o texto do rótulo de dados em um gráfico.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Representa o ângulo no qual o texto é orientado para o rótulo de dados do gráfico.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Representa a distância, em pontos, da borda superior do rótulo de dados do gráfico até a borda superior da área do gráfico.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Representa o alinhamento vertical do rótulo de dados do gráfico.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[borda](/javascript/api/excel/excel.chartdatalabelformat#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[AutoTexto](/javascript/api/excel/excel.chartdatalabels#autotext)|Especifica se os rótulos de dados geram automaticamente o texto apropriado com base no contexto.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|Especifica o alinhamento horizontal para o rótulo de dados do gráfico.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|Especifica o código de formatação para os rótulos de dados.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|Representa o ângulo no qual o texto é orientado para os rótulos de dados.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|Representa o alinhamento vertical do rótulo de dados do gráfico.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartid](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|Obtém o id do gráfico que está desativado.|
||[tipo](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|Obtém o id da planilha em que o gráfico está desativado.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartid](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|Obtém o id do gráfico que é excluído da planilha.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.chartdeletedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|Obtém o id da planilha na qual o gráfico foi deletado.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Especifica a altura do legendEntry na legenda do gráfico.|
||[índice](/javascript/api/excel/excel.chartlegendentry#index)|Especifica o índice do legendEntry na legenda do gráfico.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Especifica a esquerda de um gráfico legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Especifica a parte superior de um gráfico legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Representa a largura de legendEntry na legenda do gráfico.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[borda](/javascript/api/excel/excel.chartlegendformat#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Especifica o valor da altura de plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|Especifica o valor insideHeight de plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|Especifica o valor insideLeft de plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|Especifica o valor insideTop de plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|Especifica o valor de insideWidth de plotArea.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Especifica o valor esquerdo de plotArea.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Especifica a posição do plotArea.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Especifica a formatação de um plotArea de gráfico.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Especifica o valor principal de plotArea.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Especifica o valor da largura de plotArea.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[borda](/javascript/api/excel/excel.chartplotareaformat#border)|Especifica os atributos de borda de um plotArea de gráfico.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Especifica o formato de preenchimento de um objeto, que inclui informações de formatação de plano de fundo.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|Especifica o grupo para a série especificada.|
||[crescimento](/javascript/api/excel/excel.chartseries#explosion)|Especifica o valor de explosão para uma fatia de gráfico de pizza ou de rosca.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|Especifica o ângulo da primeira fatia do gráfico de pizza ou rosca, em graus (no sentido horário a partir da vertical).|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|True se o Excel inverte o padrão no item quando ele corresponde a um número negativo.|
||[ficar](/javascript/api/excel/excel.chartseries#overlap)|Especifica como barras e colunas são posicionadas.|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|Representa uma coleção de todos os dataLabels da série.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|Especifica o tamanho da seção secundária de um gráfico pizza de pizza ou um gráfico de barra de pizza, como uma porcentagem do tamanho da pizza principal.|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|Especifica a maneira como as duas seções de um gráfico de pizza de pizza ou de barra de pizza são divididas.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|True se o Excel atribui uma cor ou padrão diferente para cada marcador de dados.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|Representa o número de períodos que a linha de tendência se estende para trás.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|Representa o número de períodos que a linha de tendência se estende para frente.|
||[rótulo](/javascript/api/excel/excel.charttrendline#label)|Representa o rótulo de linha de tendência um gráfico.|
||[a equação](/javascript/api/excel/excel.charttrendline#showequation)|Verdadeiro se a equação da linha de tendência for exibida no gráfico.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|Verdadeiro se o R-quadrado da linha de tendência for exibido no gráfico.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[AutoTexto](/javascript/api/excel/excel.charttrendlinelabel#autotext)|Especifica se o rótulo de linhas de tendência gera automaticamente o texto apropriado com base no contexto.|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de linha de tendência usando a notação no estilo A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|Representa o alinhamento horizontal de rótulo de linha de tendência de gráfico.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Representa a distância, em pontos, da borda esquerda do rótulo da linha de tendência do gráfico até a borda esquerda da área do gráfico.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|Valor de cadeia de caracteres que representa o código do formato do rótulo de linha de tendência.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|O formato do rótulo de tendência do gráfico.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Retorna a altura, em pontos, do rótulo de linha de tendência do gráfico.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Retorna a largura, em pontos, do rótulo de linha de tendência do gráfico.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|Cadeia de caracteres que representa o texto do rótulo em um gráfico de linha de tendência.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|Representa o ângulo no qual o texto é orientado para o rótulo de tendência do gráfico.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Representa a distância, em pontos, da borda superior do rótulo de linha de tendência do gráfico até a borda superior da área do gráfico.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|Representa o alinhamento vertical do rótulo de linha de tendência de gráfico.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[borda](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Especifica o formato de borda, que inclui cor, LineStyle e Weight.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Especifica o formato de preenchimento do rótulo de tendência do gráfico atual.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Especifica os atributos de fonte (nome da fonte, tamanho da fonte, cor etc.) para um rótulo de tendência de gráfico.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|Uma fórmula de validação de dados personalizados.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Nome da DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|Formato de número do DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Posição da DataPivotHierarchy.|
||[campo](/javascript/api/excel/excel.datapivothierarchy#field)|Retorna PivotFields associados a DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|ID do DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Redefina a DataPivotHierarchy para os valores padrão.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|Especifica se os dados devem ser mostrados como um cálculo de resumo específico.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|Especifica se todos os itens do DataPivotHierarchy são exibidos.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[Adicionar (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|Adiciona o PivotHierarchy ao eixo atual.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|Obtém DataPivotHierarchy por nome ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|Obtém uma DataPivotHierarchy por nome.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[remover (DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|Remove o PivotHierarchy do eixo atual.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|Desfazer a validação de dados do intervalo atual.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|Alerta de erro quando o usuário insere dados inválidos.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|Especifica se a validação de dados será executada em células em branco, o padrão é true.|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|Avisar quando os usuários selecionarem uma célula.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Tipo de validação de dados, confira Excel.DataValidationType para obter detalhes.|
||[inválido](/javascript/api/excel/excel.datavalidation#valid)|Representa se todos os valores de célula são válidos de acordo com as regras de validação de dados.|
||[norma](/javascript/api/excel/excel.datavalidation#rule)|Regra de validação de dados que contém diferentes tipos de critérios de validação de dados.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Representa a mensagem de alerta de erro.|
||[Enviar alerta](/javascript/api/excel/excel.datavalidationerroralert#showalert)|Especifica se uma caixa de diálogo de alerta de erro deve ser exibida quando um usuário insere dados inválidos.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|O tipo de alerta de validação de dados, confira Excel. DataValidationAlertStyle para obter detalhes.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|Representa o título da caixa de diálogo de alerta de erro.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|Especifica a mensagem do prompt.|
||[Mostrar prompt](/javascript/api/excel/excel.datavalidationprompt#showprompt)|Especifica se um prompt é mostrado quando um usuário seleciona uma célula com validação de dados.|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|Especifica o título do prompt.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[cliente](/javascript/api/excel/excel.datavalidationrule#custom)|Critérios de validação de dados personalizados.|
||[data](/javascript/api/excel/excel.datavalidationrule#date)|Critérios de validação de dados de data.|
||[decimal](/javascript/api/excel/excel.datavalidationrule#decimal)|Critérios de validação de dados decimais.|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|Critérios de validação de dados da lista.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|Critérios de validação de dados TextLength.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Critérios de validação de dados de tempo.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|Critérios de validação de dados WholeNumber.|
|[DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|[Formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Especifica o operando direito quando a Propriedade Operator é definida como um operador binário como GreaterThan (o operando esquerdo é o valor que o usuário tenta inserir na célula).|
||[Formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|Com os operadores ternários between e não between, especifica o operando de limite superior.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|O operador a ser usado para validar os dados.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|Determina se deseja permitir vários itens de filtro.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Nome do FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Posição do FilterPivotHierarchy.|
||[campos](/javascript/api/excel/excel.filterpivothierarchy#fields)|Retorna PivotFields associados a FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|ID do FilterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|Redefina a FilterPivotHierarchy para os valores padrão.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[Adicionar (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|Adiciona o PivotHierarchy ao eixo atual.|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|Obtém FilterPivotHierarchy por nome ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|Obtém um FilterPivotHierarchy por nome.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[remover (filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|Remove o PivotHierarchy do eixo atual.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|Exibido na lista na célula suspensa ou não, ele será padronizado como verdadeiro.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Fonte da lista de validação de dados|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Nome do PivotField.|
||[id](/javascript/api/excel/excel.pivotfield#id)|ID do PivotField..|
||[items](/javascript/api/excel/excel.pivotfield#items)|Retorna PivotFields associados ao PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|Determina se deseja mostrar todos os itens de PivotField.|
||[sortByLabels (sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|Classifica o PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Subtotais de PivotField.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|Obtém o número de campos de tabela dinâmica na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|Obtém um PivotField por nome ou ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|Obtém um PivotField pelo nome.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Nome do PivotHierarchy.|
||[campos](/javascript/api/excel/excel.pivothierarchy#fields)|Retorna PivotFields associados a PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|ID do PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|Obtém PivotHierarchy por nome ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|Obtém o PivotHierarchy por nome.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|Determina se o item está expandido para mostrar itens filho ou se ele está recolhido e os itens filho estão ocultos.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Nome do PivotItem.|
||[id](/javascript/api/excel/excel.pivotitem#id)|ID do PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Especifica se o PivotItem é visível.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|Obtém o número de PivotItems na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|Obtém um PivotItem por seu nome ou ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|Obtém um PivotItem pelo nome.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|Retorna o intervalo onde residem os rótulos de coluna da Tabela Dinâmica.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|Retorna o intervalo onde residem os valores de dados da tabela dinâmica.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|Retorna o intervalo de área de filtro da Tabela Dinâmica.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|Retorna o intervalo em que a Tabela Dinâmica existe, excluindo a área de filtro.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|Retorna o intervalo onde residem os rótulos de linha da Tabela Dinâmica.|
||[LayoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|Essa propriedade indica o PivotLayoutType de todos os campos da Tabela Dinâmica.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|Especifica se o relatório de tabela dinâmica mostra os totais gerais para colunas.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|Especifica se o relatório de tabela dinâmica mostra os totais gerais para as linhas.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|Essa propriedade indica SubtotalLocationType de todos os campos da Tabela Dinâmica.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|Exclui a Tabela Dinâmica.|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|As hierarquias de pivô da coluna da Tabela Dinâmica.|
||[datahierarquias](/javascript/api/excel/excel.pivottable#datahierarchies)|As hierarquias dinâmicas de dados da Tabela Dinâmica.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|As hierarquias de pivô do filtro da Tabela Dinâmica.|
||[hierarquias](/javascript/api/excel/excel.pivottable#hierarchies)|Hierarquias pivô da Tabela Dinâmica.|
||[teclado](/javascript/api/excel/excel.pivottable#layout)|O PivotLayout descreve o layout e estrutura visual da Tabela Dinâmica.|
||[transhierarquias](/javascript/api/excel/excel.pivottable#rowhierarchies)|As hierarquias de pivô de linha da Tabela Dinâmica.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[Add (Name: String, Source: \| \| tabela de cadeia de caracteres de intervalo, destino: cadeia de caracteres de intervalo \| )](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|Adicione uma tabela dinâmica com base nos dados de origem especificados e insira-a na célula superior esquerda do intervalo de destino.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|Retorna um objeto de validação de dados.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Nome da RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Posição da RowColumnPivotHierarchy.|
||[campos](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Retorna PivotFields associados a RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|ID do RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|Redefine o RowColumnPivotHierarchy para os valores padrão.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[Adicionar (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|Adiciona o PivotHierarchy ao eixo atual.|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|Obtém RowColumnPivotHierarchy por nome ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|Obtém um RowColumnPivotHierarchy por nome.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[remover (rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|Remove o PivotHierarchy do eixo atual.|
|[Runtime](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|Alternar eventos JavaScript no painel de tarefas ou no suplemento de conteúdo atual.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|O PivotField base para basear o cálculo ShowAs, se aplicável com base no tipo ShowAsCalculation, caso contrário, null.|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|O Item base para basear o cálculo ShowAs, se aplicável com base no tipo ShowAsCalculation, caso contrário, null.|
||[cálculo](/javascript/api/excel/excel.showasrule#calculation)|O cálculo de ShowAs a ser usado para o Data PivotField.|
|[Estilo](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Especifica se o texto será recuado automaticamente quando o alinhamento do texto em uma célula for definido como distribuição igual.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|A orientação de texto para o estilo.|
|[Subtotals](/javascript/api/excel/excel.subtotals)|[automatic](/javascript/api/excel/excel.subtotals#automatic)|Se Automatic for definido como true, todos os outros valores serão ignorados ao definir os subtotais.|
||[normal](/javascript/api/excel/excel.subtotals#average)||
||[Count](/javascript/api/excel/excel.subtotals#count)||
||[countNumbers](/javascript/api/excel/excel.subtotals#countnumbers)||
||[número](/javascript/api/excel/excel.subtotals#max)||
||[comp](/javascript/api/excel/excel.subtotals#min)||
||[técnico](/javascript/api/excel/excel.subtotals#product)||
||[standardDeviation](/javascript/api/excel/excel.subtotals#standarddeviation)||
||[standardDeviationP](/javascript/api/excel/excel.subtotals#standarddeviationp)||
||[parcial](/javascript/api/excel/excel.subtotals#sum)||
||[matriz](/javascript/api/excel/excel.subtotals#variance)||
||[varianceP](/javascript/api/excel/excel.subtotals#variancep)||
|[Table](/javascript/api/excel/excel.table)|[legacyId](/javascript/api/excel/excel.table#legacyid)|Retorna uma ID numérica.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrange-ctx-)|Obtém o intervalo que representa a área alterada de uma tabela em uma planilha específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Obtém o intervalo que representa a área alterada de uma tabela em uma planilha específica.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|True se a pasta de trabalho estiver aberta no modo somente leitura.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[Worksheet](/javascript/api/excel/excel.worksheet)|[oncalculado](/javascript/api/excel/excel.worksheet#oncalculated)|Ocorre quando a planilha é calculada.|
||[Linhas de grade](/javascript/api/excel/excel.worksheet#showgridlines)|Especifica se as linhas de grade são visíveis para o usuário.|
||[meus títulos](/javascript/api/excel/excel.worksheet#showheadings)|Especifica se os títulos são visíveis para o usuário.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[tipo](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|Obtém a ID da planilha na qual o cálculo ocorreu.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Obtém o intervalo que representa a área alterada de uma planilha específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Obtém o intervalo que representa a área alterada de uma planilha específica.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[oncalculado](/javascript/api/excel/excel.worksheetcollection#oncalculated)|Ocorre quando qualquer planilha na pasta de trabalho é calculada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
