---
title: Conjunto de requisitos de API JavaScript do Excel 1,8
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,8
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad8c67fddffeec7937b66d43fb58a8608d662be1
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430832"
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
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[Formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Especifica o operando direito quando a Propriedade Operator é definida como um operador binário como GreaterThan (o operando esquerdo é o valor que o usuário tenta inserir na célula). Com os operadores ternários between e não between, especifica o operando de limite inferior.|
||[Formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|Com os operadores ternários between e não between, especifica o operando de limite superior. Não é usado com os operadores binários, como GreaterThan.|
||[operator](/javascript/api/excel/excel.basicdatavalidation#operator)|O operador a ser usado para validar os dados.|
|[Gráfico](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|Retorna ou define uma constante de enumeração ChartCategoryLabelLevel que se refere a|
||[displayBlanksAs](/javascript/api/excel/excel.chart#displayblanksas)|Retorna ou define a maneira como as células em branco são plotadas em um gráfico. Leitura/gravação.|
||[plotBy](/javascript/api/excel/excel.chart#plotby)|Retorna ou define como as colunas ou linhas são usadas como séries de dados no gráfico. Leitura/gravação.|
||[plotVisibleOnly](/javascript/api/excel/excel.chart#plotvisibleonly)|Verdadeiro se apenas as células visíveis forem plotadas.Falso se ambas as células visíveis e ocultas forem plotadas.. Leitura/gravação.|
||[onActivated](/javascript/api/excel/excel.chart#onactivated)|Ocorre quando o gráfico é ativado.|
||[onDeactivated](/javascript/api/excel/excel.chart#ondeactivated)|Ocorre quando o gráfico é desativado.|
||[plotArea](/javascript/api/excel/excel.chart#plotarea)|Representa a plotArea para o gráfico.|
||[seriesNameLevel](/javascript/api/excel/excel.chart#seriesnamelevel)|Retorna ou define uma constante de enumeração ChartSeriesNameLevel que se refere a|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chart#showdatalabelsovermaximum)|Representa se os rótulos de dados devem ser mostrados quando o valor for maior que o valor máximo no eixo de valor.|
||[style](/javascript/api/excel/excel.chart#style)|Retorna ou define o estilo do gráfico para o gráfico. Leitura/gravação.|
|[ChartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|[chartid](/javascript/api/excel/excel.chartactivatedeventargs#chartid)|Obtém o id do gráfico que está ativado.|
||[tipo](/javascript/api/excel/excel.chartactivatedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartactivatedeventargs#worksheetid)|Obtém o id da planilha na qual o gráfico é ativado.|
|[ChartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|[chartid](/javascript/api/excel/excel.chartaddedeventargs#chartid)|Obtém o id do gráfico que é adicionado à planilha.|
||[source](/javascript/api/excel/excel.chartaddedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.chartaddedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartaddedeventargs#worksheetid)|Obtém o id da planilha na qual o gráfico é adicionado.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[Alignment](/javascript/api/excel/excel.chartaxis#alignment)|Representa o alinhamento para o rótulo de escala do eixo especificado. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxis#isbetweencategories)|Representa se o eixo de valor cruza o eixo de categoria entre as categorias.|
||[Vários](/javascript/api/excel/excel.chartaxis#multilevel)|Representa se um eixo é multinível ou não.|
||[numberFormat](/javascript/api/excel/excel.chartaxis#numberformat)|Representa o código de formato para o rótulo de marcação do eixo.|
||[partida](/javascript/api/excel/excel.chartaxis#offset)|Representa a distância entre os níveis de rótulos e a distância entre o primeiro nível e a linha do eixo. O valor deve ser um inteiro de 0 a 1000.|
||[position](/javascript/api/excel/excel.chartaxis#position)|Representa a posição do eixo especificada onde o outro eixo cruza. Consulte Excel. ChartAxisPosition para obter detalhes.|
||[positionAt](/javascript/api/excel/excel.chartaxis#positionat)|Representa a posição do eixo especificada onde o outro eixo cruza. Você deve usar o método SetPositionAt (double) para definir essa propriedade.|
||[setPositionAt (valor: número)](/javascript/api/excel/excel.chartaxis#setpositionat-value-)|Define a posição do eixo especificada onde o outro eixo cruza.|
||[textOrientation](/javascript/api/excel/excel.chartaxis#textorientation)|Representa a orientação do texto do rótulo de seleção do eixo. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Representa a formatação de preenchimento de gráfico. Somente leitura.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setformula (fórmula: cadeia de caracteres)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|Um valor de cadeia de caracteres que representa a fórmula do título do eixo do gráfico usando a notação no estilo A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[Borderô](/javascript/api/excel/excel.chartaxistitleformat#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Representa a formatação de preenchimento de gráfico.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Limpa a formatação da borda de um elemento do gráfico.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|Ocorre quando um gráfico é ativado.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|Ocorre quando um novo gráfico é adicionado à planilha.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|Ocorre quando um gráfico é desativado.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|Ocorre quando um gráfico é excluído.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[AutoTexto](/javascript/api/excel/excel.chartdatalabel#autotext)|Valor booliano que representa se o rótulo de dados gerará automaticamente o texto apropriado com base no contexto..|
||[formula](/javascript/api/excel/excel.chartdatalabel#formula)|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de dados usando a notação no estilo A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabel#horizontalalignment)|Representa o alinhamento horizontal de rótulo de dados do gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[left](/javascript/api/excel/excel.chartdatalabel#left)|Representa a distância, em pontos, da borda esquerda do rótulo de dados do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabel#numberformat)|Valor de cadeia de caracteres que representa o código do formato do rótulo de dados.|
||[format](/javascript/api/excel/excel.chartdatalabel#format)|Representa o formato do rótulo de dados do gráfico.|
||[height](/javascript/api/excel/excel.chartdatalabel#height)|Retorna a altura, em pontos, do rótulo de dados do gráfico. Somente leitura. Nulo se o rótulo de dados do gráfico não estiver visível.|
||[width](/javascript/api/excel/excel.chartdatalabel#width)|Retorna a largura, em pontos, do rótulo de dados do gráfico. Somente leitura. Nulo se o rótulo de dados do gráfico não estiver visível.|
||[text](/javascript/api/excel/excel.chartdatalabel#text)|Cadeia de caracteres que representa o texto do rótulo de dados em um gráfico.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabel#textorientation)|Representa a orientação de texto de rótulo de dados do gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.chartdatalabel#top)|Representa a distância, em pontos, da borda superior do rótulo de dados do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabel#verticalalignment)|Representa o alinhamento vertical do rótulo de dados do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[Borderô](/javascript/api/excel/excel.chartdatalabelformat#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[AutoTexto](/javascript/api/excel/excel.chartdatalabels#autotext)|Indica se os rótulos de dados geram automaticamente texto apropriado com base no contexto.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|Representa o alinhamento horizontal de rótulo de dados do gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|Representa o código de formatação para rótulos de dados.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|Representa a orientação de texto dos rótulos de dados. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|Representa o alinhamento vertical do rótulo de dados do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartid](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|Obtém o id do gráfico que está desativado.|
||[tipo](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|Obtém o id da planilha em que o gráfico está desativado.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartid](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|Obtém o id do gráfico que é excluído da planilha.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.chartdeletedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|Obtém o id da planilha na qual o gráfico foi deletado.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Representa a altura de legendEntry na legenda do gráfico.|
||[índice](/javascript/api/excel/excel.chartlegendentry#index)|Representa o índice de legendEntry na legenda do gráfico.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Representa a esquerda de um gráfico legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Representa a parte superior de um gráfico legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Representa a largura de legendEntry na legenda do gráfico.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[Borderô](/javascript/api/excel/excel.chartlegendformat#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Representa o valor de altura de plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|Representa o valor insideHeight plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|Representa o valor insideLeft de plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|Representa o valor insideTop de plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|Representa o valor insideWidth de plotArea.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Representa o valor de plotArea à esquerda.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Represente a posição de plotArea.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Representa a formatação de um gráfico plotArea.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Representa o valor máximo de plotArea.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Representa o valor de largura de plotArea.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[Borderô](/javascript/api/excel/excel.chartplotareaformat#border)|Representa os atributos de borda de um gráfico plotArea.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|Retorna ou define o grupo da série especificada. Leitura/gravação|
||[crescimento](/javascript/api/excel/excel.chartseries#explosion)|Retorna ou define o valor de explosão para um gráfico de pizza ou fatia de gráfico de rosca. Retorna 0 (zero) se não houver explosão (a ponta da fatia está no centro da pizza). Leitura/gravação.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|Retorna ou define o ângulo do primeiro gráfico de pizza ou fatia de gráfico de rosca, em graus (no sentido horário a partir da vertical). Aplica-se apenas a pizza, torta 3-D e gráficos de rosca.. Pode ser um valor de 0 a 360. Leitura/gravação|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|Verdadeiro se o Microsoft Excel inverte o padrão no item quando ele corresponde a um número negativo. Leitura/gravação.|
||[ficar](/javascript/api/excel/excel.chartseries#overlap)|Especifica como barras e colunas são posicionadas. Pode ser um valor entre – 100 e 100. Se aplicam apenas às barras 2D e gráficos de colunas 2D. Leitura/gravação.|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|Representa uma coleção de todos os dataLabels da série.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|Retorna ou define o tamanho da seção secundária de uma pizza do gráfico de pizza ou de uma barra de gráfico de pizza, como uma porcentagem do tamanho da pizza primária. Pode ser um valor de 5 de 200. Leitura/gravação.|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|Retorna ou define a maneira como as duas seções de uma pizza do gráfico de pizza ou de uma barra do gráfico de pizza são divididas. Leitura/gravação.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|Verdadeiro se o Microsoft Excel atribuir uma cor ou padrão diferente a cada marcador de dados. O gráfico deve conter apenas uma série. Leitura/gravação.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|Representa o número de períodos que a linha de tendência se estende para trás.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|Representa o número de períodos que a linha de tendência se estende para frente.|
||[rótulo](/javascript/api/excel/excel.charttrendline#label)|Representa o rótulo de linha de tendência um gráfico.|
||[a equação](/javascript/api/excel/excel.charttrendline#showequation)|Verdadeiro se a equação da linha de tendência for exibida no gráfico.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|Verdadeiro se o R-quadrado da linha de tendência for exibido no gráfico.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[AutoTexto](/javascript/api/excel/excel.charttrendlinelabel#autotext)|Valor booliano que representa se o rótulo de linha de tendência gerará automaticamente o texto apropriado com base no contexto..|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de linha de tendência usando a notação no estilo A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|Representa o alinhamento horizontal de rótulo de linha de tendência de gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Representa a distância, em pontos, da borda esquerda do rótulo da linha de tendência do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|Valor de cadeia de caracteres que representa o código do formato do rótulo de linha de tendência.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|Representa o formato do rótulo de linha de tendência de gráfico.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Retorna a altura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Retorna a largura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|Cadeia de caracteres que representa o texto do rótulo em um gráfico de linha de tendência.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|Representa a orientação de texto de rótulo de linha de tendência de gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Representa a distância, em pontos, da borda superior do rótulo de linha de tendência do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|Representa o alinhamento vertical do rótulo de linha de tendência de gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[Borderô](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Representa o formato de preenchimento do rótulo de linha de tendência atual do gráfico.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Representa os atributos de fonte do rótulo de linha de tendência do gráfico, como nome, tamanho, cor, dentre outros.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|Uma fórmula de validação de dados personalizados. Isso cria regras de entrada especiais, como impedir duplicatas ou limitar o total em um intervalo de células.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Nome da DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|Formato de número do DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Posição da DataPivotHierarchy.|
||[campo](/javascript/api/excel/excel.datapivothierarchy#field)|Retorna PivotFields associados a DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|ID do DataPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Redefina a DataPivotHierarchy para os valores padrão.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|Determina se os dados devem ser mostrados como um cálculo de resumo específico ou não.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|Determina se deve mostrar todos os itens a DataPivotHierarchy.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[Adicionar (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|Adiciona o PivotHierarchy ao eixo atual.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|Obtém DataPivotHierarchy por nome ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|Obtém uma DataPivotHierarchy por nome. Se o DataPivotHierarchy não existir, retornará um objeto nulo.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[remover (DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|Remove o PivotHierarchy do eixo atual.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|Desfazer a validação de dados do intervalo atual.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|Alerta de erro quando o usuário insere dados inválidos.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|Ignora espaços em branco: nenhuma validação de dados será executada em células vazias, o padrão será definido como verdadeiro.|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|Avisar quando os usuários selecionarem uma célula.|
||[type](/javascript/api/excel/excel.datavalidation#type)|Tipo de validação de dados, confira Excel.DataValidationType para obter detalhes.|
||[inválido](/javascript/api/excel/excel.datavalidation#valid)|Representa se todos os valores de célula são válidos de acordo com as regras de validação de dados.|
||[norma](/javascript/api/excel/excel.datavalidation#rule)|Regra de validação de dados que contém diferentes tipos de critérios de validação de dados.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Representa a mensagem de alerta de erro.|
||[Enviar alerta](/javascript/api/excel/excel.datavalidationerroralert#showalert)|Determina se deseja mostrar uma caixa de diálogo de alerta de erro ou não quando um usuário insere dados inválidos. O padrão é verdadeiro.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|Representa o tipo de alerta de validação de dados, confira Excel.DataValidationAlertStyle para obter detalhes.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|Representa o título da caixa de diálogo de alerta de erro.|
|[DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|[message](/javascript/api/excel/excel.datavalidationprompt#message)|Representa a mensagem a solicitação.|
||[Mostrar prompt](/javascript/api/excel/excel.datavalidationprompt#showprompt)|Determina se deseja ou não mostrar o prompt quando o usuário seleciona uma célula com a validação de dados.|
||[title](/javascript/api/excel/excel.datavalidationprompt#title)|Representa o título para a solicitação.|
|[DataValidationRule](/javascript/api/excel/excel.datavalidationrule)|[cliente](/javascript/api/excel/excel.datavalidationrule#custom)|Critérios de validação de dados personalizados.|
||[data](/javascript/api/excel/excel.datavalidationrule#date)|Critérios de validação de dados de data.|
||[dígitos](/javascript/api/excel/excel.datavalidationrule#decimal)|Critérios de validação de dados decimais.|
||[list](/javascript/api/excel/excel.datavalidationrule#list)|Critérios de validação de dados da lista.|
||[textLength](/javascript/api/excel/excel.datavalidationrule#textlength)|Critérios de validação de dados TextLength.|
||[time](/javascript/api/excel/excel.datavalidationrule#time)|Critérios de validação de dados de tempo.|
||[wholeNumber](/javascript/api/excel/excel.datavalidationrule#wholenumber)|Critérios de validação de dados WholeNumber.|
|[Datetimedatavalidationcomo](/javascript/api/excel/excel.datetimedatavalidation)|[Formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Especifica o operando direito quando a Propriedade Operator é definida como um operador binário como GreaterThan (o operando esquerdo é o valor que o usuário tenta inserir na célula). Com os operadores ternários between e não between, especifica o operando de limite inferior.|
||[Formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|Com os operadores ternários between e não between, especifica o operando de limite superior. Não é usado com os operadores binários, como GreaterThan.|
||[operator](/javascript/api/excel/excel.datetimedatavalidation#operator)|O operador a ser usado para validar os dados.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|Determina se deseja permitir vários itens de filtro.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Nome do FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Posição do FilterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Retorna PivotFields associados a FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|ID do FilterPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|Redefina a FilterPivotHierarchy para os valores padrão.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[Adicionar (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|Adiciona o PivotHierarchy ao eixo atual. Se houver a hierarquia em outro lugar na linha, coluna,|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|Obtém FilterPivotHierarchy por nome ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|Obtém um FilterPivotHierarchy por nome. Se o FilterPivotHierarchy não existir, retornará um objeto nulo.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[remover (filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|Remove o PivotHierarchy do eixo atual.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|Exibido na lista na célula suspensa ou não, ele será padronizado como verdadeiro.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Fonte da lista de validação de dados|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Nome do PivotField.|
||[id](/javascript/api/excel/excel.pivotfield#id)|ID do PivotField..|
||[items](/javascript/api/excel/excel.pivotfield#items)|Retorna o PivotItems que é composto com o PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|Determina se deseja mostrar todos os itens de PivotField.|
||[sortByLabels (sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|Classifica o PivotField. Se um DataPivotHierarchy for especificado, a classificação será aplicada com base nele, se a classificação não for baseada no campo PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Subtotais de PivotField.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|Obtém o número de campos de tabela dinâmica na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|Obtém um PivotField por nome ou ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|Obtém um PivotField pelo nome. Se PivotField não existir, retornará um objeto NULL.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Nome do PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Retorna PivotFields associados a PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|ID do PivotHierarchy.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|Obtém PivotHierarchy por nome ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|Obtém o PivotHierarchy por nome. Se o PivotHierarchy não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|Determina se o item está expandido para mostrar itens filho ou se ele está recolhido e os itens filho estão ocultos.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Nome do PivotItem.|
||[id](/javascript/api/excel/excel.pivotitem#id)|ID do PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Determina se o PivotItem ficará visível ou não.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|Obtém o número de itens de tabela dinâmica na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|Obtém um PivotItem por seu nome ou ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|Obtém um PivotItem pelo nome. Se o PivotItem não existir, retornará um objeto NULL.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|Retorna o intervalo onde residem os rótulos de coluna da Tabela Dinâmica.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|Retorna o intervalo onde residem os valores de dados da tabela dinâmica.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|Retorna o intervalo de área de filtro da Tabela Dinâmica.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|Retorna o intervalo em que a Tabela Dinâmica existe, excluindo a área de filtro.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|Retorna o intervalo onde residem os rótulos de linha da Tabela Dinâmica.|
||[LayoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|Essa propriedade indica o PivotLayoutType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|Especifica se o relatório de tabela dinâmica mostra os totais gerais das colunas.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|Especifica se o relatório de tabela dinâmica mostra os totais gerais para as linhas.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|Essa propriedade indica SubtotalLocationType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|Exclui a Tabela Dinâmica.|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|As hierarquias de pivô da coluna da Tabela Dinâmica.|
||[datahierarquias](/javascript/api/excel/excel.pivottable#datahierarchies)|As hierarquias dinâmicas de dados da Tabela Dinâmica.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|As hierarquias de pivô do filtro da Tabela Dinâmica.|
||[hierarquias](/javascript/api/excel/excel.pivottable#hierarchies)|Hierarquias pivô da Tabela Dinâmica.|
||[teclado](/javascript/api/excel/excel.pivottable#layout)|O PivotLayout descreve o layout e estrutura visual da Tabela Dinâmica.|
||[transhierarquias](/javascript/api/excel/excel.pivottable#rowhierarchies)|As hierarquias de pivô de linha da Tabela Dinâmica.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[Add (Name: String, Source: \| \| tabela de cadeia de caracteres de intervalo, destino: cadeia de caracteres de intervalo \| )](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|Adiciona um Pivottable com base nos dados de origem especificados e insere-o na célula superior esquerda do intervalo de destino.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|Retorna um objeto de validação de dados.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Nome da RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Posição da RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Retorna PivotFields associados a RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|ID do RowColumnPivotHierarchy.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|Redefine o RowColumnPivotHierarchy para os valores padrão.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[Adicionar (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|Adiciona o PivotHierarchy ao eixo atual. Se houver a hierarquia em outro lugar na linha, coluna,|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|Obtém RowColumnPivotHierarchy por nome ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|Obtém um RowColumnPivotHierarchy por nome. Se o RowColumnPivotHierarchy não existir, retornará um objeto nulo.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[remover (rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|Remove o PivotHierarchy do eixo atual.|
|[Tempo de execução](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|Alternar eventos JavaScript no painel de tarefas ou no suplemento de conteúdo atual.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|O PivotField base para basear o cálculo ShowAs, se aplicável com base no tipo ShowAsCalculation, caso contrário, null.|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|O Item base para basear o cálculo ShowAs, se aplicável com base no tipo ShowAsCalculation, caso contrário, null.|
||[cálculo](/javascript/api/excel/excel.showasrule#calculation)|O cálculo de ShowAs a ser usado para o Data PivotField. Consulte Excel. ShowAsCalculation para obter detalhes.|
|[Style](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Indica se o texto é automaticamente indentado quando o alinhamento de texto em uma célula é definido como distribuição igual.|
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
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.tablechangedeventargs#getrangeornullobject-ctx-)|Obtém o intervalo que representa a área alterada de uma tabela em uma planilha específica. Pode retornar o objeto null.|
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|True se a pasta de trabalho estiver aberta no modo somente leitura. Somente leitura.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[Planilha](/javascript/api/excel/excel.worksheet)|[oncalculado](/javascript/api/excel/excel.worksheet#oncalculated)|Ocorre quando a planilha é calculada.|
||[Linhas de grade](/javascript/api/excel/excel.worksheet#showgridlines)|Obtém ou define um sinalizador de linhas de grade da planilha.|
||[meus títulos](/javascript/api/excel/excel.worksheet#showheadings)|É ou define um sinalizador de cabeçalhos da planilha.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[tipo](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|Obtém o id da planilha que é calculada.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Obtém o intervalo que representa a área alterada de uma planilha específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Obtém o intervalo que representa a área alterada de uma planilha específica. Pode retornar o objeto null.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[oncalculado](/javascript/api/excel/excel.worksheetcollection#oncalculated)|Ocorre quando qualquer planilha na pasta de trabalho é calculada.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.8&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
