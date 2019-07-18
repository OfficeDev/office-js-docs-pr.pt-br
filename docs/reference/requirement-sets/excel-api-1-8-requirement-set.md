---
title: Conjunto de requisitos de API JavaScript do Excel 1,8
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,8
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a5adcf56654070ca2a8336385f73062c34e90e1d
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35772006"
---
# <a name="whats-new-in-excel-javascript-api-18"></a>Quais são as novidades na API JavaScript do Excel 1.8

O conjunto de requisitos 1.8 da API JavaScript do Excel inclui APIs para tabelas dinâmicas, validação de dados, gráficos, eventos de gráficos, opções de desempenho e criação de pasta de trabalho.

## <a name="pivottable"></a>Tabela Dinâmica

Onda 2 das APIs de Tabela Dinâmica permite que os suplementos definam as hierarquias de uma Tabela Dinâmica. Agora você pode controlar os dados e como eles são agregados. Nosso [Artigo de Tabela Dinâmica](/office/dev/add-ins/excel/excel-add-ins-pivottables) tem mais informações sobre a nova funcionalidade de tabela dinâmica.

## <a name="data-validation"></a>Validação de Dados

A validação de dados permite controlar o que um usuário digita em uma planilha. Você pode limitar as células a conjuntos de respostas predefinidos ou fornecer avisos pop-up sobre entradas indesejadas. Saiba mais sobre [adicionar a validação de dados para intervalos](/office/dev/add-ins/excel/excel-add-ins-data-validation) hoje.

## <a name="charts"></a>Gráficos

Outra rodada de APIs de gráficos traz um controle programático ainda maior sobre os elementos do gráfico. Agora você tem maior acesso à legenda, eixos, linha de tendência e área de plotagem.

## <a name="events"></a>Eventos

Mais [eventos](/office/dev/add-ins/excel/excel-add-ins-events) foram adicionados para os gráficos. Faça o seu suplemento reagir aos usuários interagindo com o gráfico. Você também pode [alternar eventos](/office/dev/add-ins/excel/performance#enable-and-disable-events) disparados em toda a pasta de trabalho.

## <a name="api-list"></a>Lista de APIs

| Classe | Campos | Descrição |
|:---|:---|:---|
|[BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|[Formula1](/javascript/api/excel/excel.basicdatavalidation#formula1)|Especifica o operando direito quando a Propriedade Operator é definida como um operador binário como GreaterThan (o operando esquerdo é o valor que o usuário tenta inserir na célula). Com os operadores ternários between e não between, especifica o operando de limite inferior.|
||[Formula2](/javascript/api/excel/excel.basicdatavalidation#formula2)|Com os operadores ternários between e não between, especifica o operando de limite superior. Não é usado com os operadores binários, como GreaterThan.|
||[operador](/javascript/api/excel/excel.basicdatavalidation#operator)|O operador a ser usado para validar os dados.|
|[Chart](/javascript/api/excel/excel.chart)|[categoryLabelLevel](/javascript/api/excel/excel.chart#categorylabellevel)|Retorna ou define uma constante de enumeração ChartCategoryLabelLevel que se refere a|
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
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[Alignment](/javascript/api/excel/excel.chartaxisdata#alignment)|Representa o alinhamento para o rótulo de escala do eixo especificado. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisdata#isbetweencategories)|Representa se o eixo de valor cruza o eixo de categoria entre as categorias.|
||[Vários](/javascript/api/excel/excel.chartaxisdata#multilevel)|Representa se um eixo é multinível ou não.|
||[numberFormat](/javascript/api/excel/excel.chartaxisdata#numberformat)|Representa o código de formato para o rótulo de marcação do eixo.|
||[partida](/javascript/api/excel/excel.chartaxisdata#offset)|Representa a distância entre os níveis de rótulos e a distância entre o primeiro nível e a linha do eixo. O valor deve ser um inteiro de 0 a 1000.|
||[position](/javascript/api/excel/excel.chartaxisdata#position)|Representa a posição do eixo especificada onde o outro eixo cruza. Consulte Excel. ChartAxisPosition para obter detalhes.|
||[positionAt](/javascript/api/excel/excel.chartaxisdata#positionat)|Representa a posição do eixo especificada onde o outro eixo cruza. Você deve usar o método SetPositionAt (double) para definir essa propriedade.|
||[textOrientation](/javascript/api/excel/excel.chartaxisdata#textorientation)|Representa a orientação do texto do rótulo de seleção do eixo. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
|[ChartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|[fill](/javascript/api/excel/excel.chartaxisformat#fill)|Representa a formatação de preenchimento de gráfico. Somente leitura.|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[Alignment](/javascript/api/excel/excel.chartaxisloadoptions#alignment)|Representa o alinhamento para o rótulo de escala do eixo especificado. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisloadoptions#isbetweencategories)|Representa se o eixo de valor cruza o eixo de categoria entre as categorias.|
||[Vários](/javascript/api/excel/excel.chartaxisloadoptions#multilevel)|Representa se um eixo é multinível ou não.|
||[numberFormat](/javascript/api/excel/excel.chartaxisloadoptions#numberformat)|Representa o código de formato para o rótulo de marcação do eixo.|
||[partida](/javascript/api/excel/excel.chartaxisloadoptions#offset)|Representa a distância entre os níveis de rótulos e a distância entre o primeiro nível e a linha do eixo. O valor deve ser um inteiro de 0 a 1000.|
||[position](/javascript/api/excel/excel.chartaxisloadoptions#position)|Representa a posição do eixo especificada onde o outro eixo cruza. Consulte Excel. ChartAxisPosition para obter detalhes.|
||[positionAt](/javascript/api/excel/excel.chartaxisloadoptions#positionat)|Representa a posição do eixo especificada onde o outro eixo cruza. Você deve usar o método SetPositionAt (double) para definir essa propriedade.|
||[textOrientation](/javascript/api/excel/excel.chartaxisloadoptions#textorientation)|Representa a orientação do texto do rótulo de seleção do eixo. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[setformula (fórmula: cadeia de caracteres)](/javascript/api/excel/excel.chartaxistitle#setformula-formula-)|Um valor de cadeia de caracteres que representa a fórmula do título do eixo do gráfico usando a notação no estilo A1.|
|[ChartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|[Borderô](/javascript/api/excel/excel.chartaxistitleformat#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
||[fill](/javascript/api/excel/excel.chartaxistitleformat#fill)|Representa a formatação de preenchimento de gráfico.|
|[ChartAxisTitleFormatData](/javascript/api/excel/excel.chartaxistitleformatdata)|[Borderô](/javascript/api/excel/excel.chartaxistitleformatdata#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
|[ChartAxisTitleFormatLoadOptions](/javascript/api/excel/excel.chartaxistitleformatloadoptions)|[Borderô](/javascript/api/excel/excel.chartaxistitleformatloadoptions#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
|[ChartAxisTitleFormatUpdateData](/javascript/api/excel/excel.chartaxistitleformatupdatedata)|[Borderô](/javascript/api/excel/excel.chartaxistitleformatupdatedata#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[Alignment](/javascript/api/excel/excel.chartaxisupdatedata#alignment)|Representa o alinhamento para o rótulo de escala do eixo especificado. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[isBetweenCategories](/javascript/api/excel/excel.chartaxisupdatedata#isbetweencategories)|Representa se o eixo de valor cruza o eixo de categoria entre as categorias.|
||[Vários](/javascript/api/excel/excel.chartaxisupdatedata#multilevel)|Representa se um eixo é multinível ou não.|
||[numberFormat](/javascript/api/excel/excel.chartaxisupdatedata#numberformat)|Representa o código de formato para o rótulo de marcação do eixo.|
||[partida](/javascript/api/excel/excel.chartaxisupdatedata#offset)|Representa a distância entre os níveis de rótulos e a distância entre o primeiro nível e a linha do eixo. O valor deve ser um inteiro de 0 a 1000.|
||[position](/javascript/api/excel/excel.chartaxisupdatedata#position)|Representa a posição do eixo especificada onde o outro eixo cruza. Consulte Excel. ChartAxisPosition para obter detalhes.|
||[textOrientation](/javascript/api/excel/excel.chartaxisupdatedata#textorientation)|Representa a orientação do texto do rótulo de seleção do eixo. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[clear()](/javascript/api/excel/excel.chartborder#clear--)|Limpa a formatação da borda de um elemento do gráfico.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[onActivated](/javascript/api/excel/excel.chartcollection#onactivated)|Ocorre quando um gráfico é ativado.|
||[onAdded](/javascript/api/excel/excel.chartcollection#onadded)|Ocorre quando um novo gráfico é adicionado à planilha.|
||[onDeactivated](/javascript/api/excel/excel.chartcollection#ondeactivated)|Ocorre quando um gráfico é desativado.|
||[onDeleted](/javascript/api/excel/excel.chartcollection#ondeleted)|Ocorre quando um gráfico é excluído.|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[categoryLabelLevel](/javascript/api/excel/excel.chartcollectionloadoptions#categorylabellevel)|Para cada ITEM na coleção: Retorna ou define uma constante de enumeração ChartCategoryLabelLevel referindo-se a|
||[displayBlanksAs](/javascript/api/excel/excel.chartcollectionloadoptions#displayblanksas)|Para cada ITEM na coleção: Retorna ou define a maneira como as células em branco são plotadas em um gráfico. Leitura/gravação.|
||[plotArea](/javascript/api/excel/excel.chartcollectionloadoptions#plotarea)|Para cada ITEM na coleção: representa o plotArea do gráfico.|
||[plotBy](/javascript/api/excel/excel.chartcollectionloadoptions#plotby)|Para cada ITEM na coleção: Retorna ou define a maneira como colunas ou linhas são usadas como séries de dados no gráfico. Leitura/gravação.|
||[plotVisibleOnly](/javascript/api/excel/excel.chartcollectionloadoptions#plotvisibleonly)|Para cada ITEM na coleção: true se somente as células visíveis são plotadas.Falso se ambas as células visíveis e ocultas forem plotadas.. Leitura/gravação.|
||[seriesNameLevel](/javascript/api/excel/excel.chartcollectionloadoptions#seriesnamelevel)|Para cada ITEM na coleção: Retorna ou define uma constante de enumeração ChartSeriesNameLevel referindo-se a|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartcollectionloadoptions#showdatalabelsovermaximum)|Para cada ITEM na coleção: indica se os rótulos de dados devem ser mostrados quando o valor for maior do que o valor máximo no eixo dos valores.|
||[style](/javascript/api/excel/excel.chartcollectionloadoptions#style)|Para cada ITEM na coleção: Retorna ou define o estilo de gráfico do gráfico. Leitura/gravação.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[categoryLabelLevel](/javascript/api/excel/excel.chartdata#categorylabellevel)|Retorna ou define uma constante de enumeração ChartCategoryLabelLevel que se refere a|
||[displayBlanksAs](/javascript/api/excel/excel.chartdata#displayblanksas)|Retorna ou define a maneira como as células em branco são plotadas em um gráfico. Leitura/gravação.|
||[plotArea](/javascript/api/excel/excel.chartdata#plotarea)|Representa a plotArea para o gráfico.|
||[plotBy](/javascript/api/excel/excel.chartdata#plotby)|Retorna ou define como as colunas ou linhas são usadas como séries de dados no gráfico. Leitura/gravação.|
||[plotVisibleOnly](/javascript/api/excel/excel.chartdata#plotvisibleonly)|Verdadeiro se apenas as células visíveis forem plotadas.Falso se ambas as células visíveis e ocultas forem plotadas.. Leitura/gravação.|
||[seriesNameLevel](/javascript/api/excel/excel.chartdata#seriesnamelevel)|Retorna ou define uma constante de enumeração ChartSeriesNameLevel que se refere a|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartdata#showdatalabelsovermaximum)|Representa se os rótulos de dados devem ser mostrados quando o valor for maior que o valor máximo no eixo de valor.|
||[style](/javascript/api/excel/excel.chartdata#style)|Retorna ou define o estilo do gráfico para o gráfico. Leitura/gravação.|
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
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[AutoTexto](/javascript/api/excel/excel.chartdatalabeldata#autotext)|Valor booliano que representa se o rótulo de dados gerará automaticamente o texto apropriado com base no contexto..|
||[format](/javascript/api/excel/excel.chartdatalabeldata#format)|Representa o formato do rótulo de dados do gráfico.|
||[formula](/javascript/api/excel/excel.chartdatalabeldata#formula)|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de dados usando a notação no estilo A1.|
||[height](/javascript/api/excel/excel.chartdatalabeldata#height)|Retorna a altura, em pontos, do rótulo de dados do gráfico. Somente leitura. Nulo se o rótulo de dados do gráfico não estiver visível.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabeldata#horizontalalignment)|Representa o alinhamento horizontal de rótulo de dados do gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[left](/javascript/api/excel/excel.chartdatalabeldata#left)|Representa a distância, em pontos, da borda esquerda do rótulo de dados do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabeldata#numberformat)|Valor de cadeia de caracteres que representa o código do formato do rótulo de dados.|
||[text](/javascript/api/excel/excel.chartdatalabeldata#text)|Cadeia de caracteres que representa o texto do rótulo de dados em um gráfico.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabeldata#textorientation)|Representa a orientação de texto de rótulo de dados do gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.chartdatalabeldata#top)|Representa a distância, em pontos, da borda superior do rótulo de dados do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabeldata#verticalalignment)|Representa o alinhamento vertical do rótulo de dados do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
||[width](/javascript/api/excel/excel.chartdatalabeldata#width)|Retorna a largura, em pontos, do rótulo de dados do gráfico. Somente leitura. Nulo se o rótulo de dados do gráfico não estiver visível.|
|[ChartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|[Borderô](/javascript/api/excel/excel.chartdatalabelformat#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|
|[ChartDataLabelFormatData](/javascript/api/excel/excel.chartdatalabelformatdata)|[Borderô](/javascript/api/excel/excel.chartdatalabelformatdata#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|
|[ChartDataLabelFormatLoadOptions](/javascript/api/excel/excel.chartdatalabelformatloadoptions)|[Borderô](/javascript/api/excel/excel.chartdatalabelformatloadoptions#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
|[ChartDataLabelFormatUpdateData](/javascript/api/excel/excel.chartdatalabelformatupdatedata)|[Borderô](/javascript/api/excel/excel.chartdatalabelformatupdatedata#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[AutoTexto](/javascript/api/excel/excel.chartdatalabelloadoptions#autotext)|Valor booliano que representa se o rótulo de dados gerará automaticamente o texto apropriado com base no contexto..|
||[format](/javascript/api/excel/excel.chartdatalabelloadoptions#format)|Representa o formato do rótulo de dados do gráfico.|
||[formula](/javascript/api/excel/excel.chartdatalabelloadoptions#formula)|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de dados usando a notação no estilo A1.|
||[height](/javascript/api/excel/excel.chartdatalabelloadoptions#height)|Retorna a altura, em pontos, do rótulo de dados do gráfico. Somente leitura. Nulo se o rótulo de dados do gráfico não estiver visível.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelloadoptions#horizontalalignment)|Representa o alinhamento horizontal de rótulo de dados do gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[left](/javascript/api/excel/excel.chartdatalabelloadoptions#left)|Representa a distância, em pontos, da borda esquerda do rótulo de dados do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelloadoptions#numberformat)|Valor de cadeia de caracteres que representa o código do formato do rótulo de dados.|
||[text](/javascript/api/excel/excel.chartdatalabelloadoptions#text)|Cadeia de caracteres que representa o texto do rótulo de dados em um gráfico.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelloadoptions#textorientation)|Representa a orientação de texto de rótulo de dados do gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.chartdatalabelloadoptions#top)|Representa a distância, em pontos, da borda superior do rótulo de dados do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelloadoptions#verticalalignment)|Representa o alinhamento vertical do rótulo de dados do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
||[width](/javascript/api/excel/excel.chartdatalabelloadoptions#width)|Retorna a largura, em pontos, do rótulo de dados do gráfico. Somente leitura. Nulo se o rótulo de dados do gráfico não estiver visível.|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[AutoTexto](/javascript/api/excel/excel.chartdatalabelupdatedata#autotext)|Valor booliano que representa se o rótulo de dados gerará automaticamente o texto apropriado com base no contexto..|
||[format](/javascript/api/excel/excel.chartdatalabelupdatedata#format)|Representa o formato do rótulo de dados do gráfico.|
||[formula](/javascript/api/excel/excel.chartdatalabelupdatedata#formula)|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de dados usando a notação no estilo A1.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelupdatedata#horizontalalignment)|Representa o alinhamento horizontal de rótulo de dados do gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[left](/javascript/api/excel/excel.chartdatalabelupdatedata#left)|Representa a distância, em pontos, da borda esquerda do rótulo de dados do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelupdatedata#numberformat)|Valor de cadeia de caracteres que representa o código do formato do rótulo de dados.|
||[text](/javascript/api/excel/excel.chartdatalabelupdatedata#text)|Cadeia de caracteres que representa o texto do rótulo de dados em um gráfico.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelupdatedata#textorientation)|Representa a orientação de texto de rótulo de dados do gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.chartdatalabelupdatedata#top)|Representa a distância, em pontos, da borda superior do rótulo de dados do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelupdatedata#verticalalignment)|Representa o alinhamento vertical do rótulo de dados do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
|[ChartDataLabels](/javascript/api/excel/excel.chartdatalabels)|[AutoTexto](/javascript/api/excel/excel.chartdatalabels#autotext)|Indica se os rótulos de dados geram automaticamente texto apropriado com base no contexto.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabels#horizontalalignment)|Representa o alinhamento horizontal de rótulo de dados do gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabels#numberformat)|Representa o código de formatação para rótulos de dados.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabels#textorientation)|Representa a orientação de texto dos rótulos de dados. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabels#verticalalignment)|Representa o alinhamento vertical do rótulo de dados do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
|[ChartDataLabelsData](/javascript/api/excel/excel.chartdatalabelsdata)|[AutoTexto](/javascript/api/excel/excel.chartdatalabelsdata#autotext)|Indica se os rótulos de dados geram automaticamente texto apropriado com base no contexto.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsdata#horizontalalignment)|Representa o alinhamento horizontal de rótulo de dados do gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsdata#numberformat)|Representa o código de formatação para rótulos de dados.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsdata#textorientation)|Representa a orientação de texto dos rótulos de dados. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsdata#verticalalignment)|Representa o alinhamento vertical do rótulo de dados do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
|[ChartDataLabelsLoadOptions](/javascript/api/excel/excel.chartdatalabelsloadoptions)|[AutoTexto](/javascript/api/excel/excel.chartdatalabelsloadoptions#autotext)|Indica se os rótulos de dados geram automaticamente texto apropriado com base no contexto.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsloadoptions#horizontalalignment)|Representa o alinhamento horizontal de rótulo de dados do gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsloadoptions#numberformat)|Representa o código de formatação para rótulos de dados.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsloadoptions#textorientation)|Representa a orientação de texto dos rótulos de dados. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsloadoptions#verticalalignment)|Representa o alinhamento vertical do rótulo de dados do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
|[ChartDataLabelsUpdateData](/javascript/api/excel/excel.chartdatalabelsupdatedata)|[AutoTexto](/javascript/api/excel/excel.chartdatalabelsupdatedata#autotext)|Indica se os rótulos de dados geram automaticamente texto apropriado com base no contexto.|
||[horizontalAlignment](/javascript/api/excel/excel.chartdatalabelsupdatedata#horizontalalignment)|Representa o alinhamento horizontal de rótulo de dados do gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[numberFormat](/javascript/api/excel/excel.chartdatalabelsupdatedata#numberformat)|Representa o código de formatação para rótulos de dados.|
||[textOrientation](/javascript/api/excel/excel.chartdatalabelsupdatedata#textorientation)|Representa a orientação de texto dos rótulos de dados. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[verticalAlignment](/javascript/api/excel/excel.chartdatalabelsupdatedata#verticalalignment)|Representa o alinhamento vertical do rótulo de dados do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
|[ChartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|[chartid](/javascript/api/excel/excel.chartdeactivatedeventargs#chartid)|Obtém o id do gráfico que está desativado.|
||[tipo](/javascript/api/excel/excel.chartdeactivatedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartdeactivatedeventargs#worksheetid)|Obtém o id da planilha em que o gráfico está desativado.|
|[ChartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|[chartid](/javascript/api/excel/excel.chartdeletedeventargs#chartid)|Obtém o id do gráfico que é excluído da planilha.|
||[source](/javascript/api/excel/excel.chartdeletedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.chartdeletedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.chartdeletedeventargs#worksheetid)|Obtém o id da planilha na qual o gráfico foi deletado.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[height](/javascript/api/excel/excel.chartlegendentry#height)|Representa a altura de legendEntry na legenda do gráfico.|
||[index](/javascript/api/excel/excel.chartlegendentry#index)|Representa o índice de legendEntry na legenda do gráfico.|
||[left](/javascript/api/excel/excel.chartlegendentry#left)|Representa a esquerda de um gráfico legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentry#top)|Representa a parte superior de um gráfico legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentry#width)|Representa a largura de legendEntry na legenda do gráfico.|
|[ChartLegendEntryCollectionLoadOptions](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions)|[height](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#height)|Para cada ITEM na coleção: representa a altura do legendEntry na legenda do gráfico.|
||[index](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#index)|Para cada ITEM na coleção: representa o índice do legendEntry na legenda do gráfico.|
||[left](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#left)|Para cada ITEM na coleção: representa a esquerda de um gráfico legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#top)|Para cada ITEM na coleção: representa a parte superior de um gráfico legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#width)|Para cada ITEM na coleção: representa a largura do legendEntry na legenda do gráfico.|
|[ChartLegendEntryData](/javascript/api/excel/excel.chartlegendentrydata)|[height](/javascript/api/excel/excel.chartlegendentrydata#height)|Representa a altura de legendEntry na legenda do gráfico.|
||[index](/javascript/api/excel/excel.chartlegendentrydata#index)|Representa o índice de legendEntry na legenda do gráfico.|
||[left](/javascript/api/excel/excel.chartlegendentrydata#left)|Representa a esquerda de um gráfico legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentrydata#top)|Representa a parte superior de um gráfico legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentrydata#width)|Representa a largura de legendEntry na legenda do gráfico.|
|[ChartLegendEntryLoadOptions](/javascript/api/excel/excel.chartlegendentryloadoptions)|[height](/javascript/api/excel/excel.chartlegendentryloadoptions#height)|Representa a altura de legendEntry na legenda do gráfico.|
||[index](/javascript/api/excel/excel.chartlegendentryloadoptions#index)|Representa o índice de legendEntry na legenda do gráfico.|
||[left](/javascript/api/excel/excel.chartlegendentryloadoptions#left)|Representa a esquerda de um gráfico legendEntry.|
||[top](/javascript/api/excel/excel.chartlegendentryloadoptions#top)|Representa a parte superior de um gráfico legendEntry.|
||[width](/javascript/api/excel/excel.chartlegendentryloadoptions#width)|Representa a largura de legendEntry na legenda do gráfico.|
|[ChartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|[Borderô](/javascript/api/excel/excel.chartlegendformat#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|
|[ChartLegendFormatData](/javascript/api/excel/excel.chartlegendformatdata)|[Borderô](/javascript/api/excel/excel.chartlegendformatdata#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|
|[ChartLegendFormatLoadOptions](/javascript/api/excel/excel.chartlegendformatloadoptions)|[Borderô](/javascript/api/excel/excel.chartlegendformatloadoptions#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
|[ChartLegendFormatUpdateData](/javascript/api/excel/excel.chartlegendformatupdatedata)|[Borderô](/javascript/api/excel/excel.chartlegendformatupdatedata#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[categoryLabelLevel](/javascript/api/excel/excel.chartloadoptions#categorylabellevel)|Retorna ou define uma constante de enumeração ChartCategoryLabelLevel que se refere a|
||[displayBlanksAs](/javascript/api/excel/excel.chartloadoptions#displayblanksas)|Retorna ou define a maneira como as células em branco são plotadas em um gráfico. Leitura/gravação.|
||[plotArea](/javascript/api/excel/excel.chartloadoptions#plotarea)|Representa a plotArea para o gráfico.|
||[plotBy](/javascript/api/excel/excel.chartloadoptions#plotby)|Retorna ou define como as colunas ou linhas são usadas como séries de dados no gráfico. Leitura/gravação.|
||[plotVisibleOnly](/javascript/api/excel/excel.chartloadoptions#plotvisibleonly)|Verdadeiro se apenas as células visíveis forem plotadas.Falso se ambas as células visíveis e ocultas forem plotadas.. Leitura/gravação.|
||[seriesNameLevel](/javascript/api/excel/excel.chartloadoptions#seriesnamelevel)|Retorna ou define uma constante de enumeração ChartSeriesNameLevel que se refere a|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartloadoptions#showdatalabelsovermaximum)|Representa se os rótulos de dados devem ser mostrados quando o valor for maior que o valor máximo no eixo de valor.|
||[style](/javascript/api/excel/excel.chartloadoptions#style)|Retorna ou define o estilo do gráfico para o gráfico. Leitura/gravação.|
|[ChartPlotArea](/javascript/api/excel/excel.chartplotarea)|[height](/javascript/api/excel/excel.chartplotarea#height)|Representa o valor de altura de plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarea#insideheight)|Representa o valor insideHeight plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarea#insideleft)|Representa o valor insideLeft de plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarea#insidetop)|Representa o valor insideTop de plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarea#insidewidth)|Representa o valor insideWidth de plotArea.|
||[left](/javascript/api/excel/excel.chartplotarea#left)|Representa o valor de plotArea à esquerda.|
||[position](/javascript/api/excel/excel.chartplotarea#position)|Represente a posição de plotArea.|
||[format](/javascript/api/excel/excel.chartplotarea#format)|Representa a formatação de um gráfico plotArea.|
||[Set (Propriedades: Excel. ChartPlotArea)](/javascript/api/excel/excel.chartplotarea#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartPlotAreaUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.chartplotarea#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[top](/javascript/api/excel/excel.chartplotarea#top)|Representa o valor máximo de plotArea.|
||[width](/javascript/api/excel/excel.chartplotarea#width)|Representa o valor de largura de plotArea.|
|[ChartPlotAreaData](/javascript/api/excel/excel.chartplotareadata)|[format](/javascript/api/excel/excel.chartplotareadata#format)|Representa a formatação de um gráfico plotArea.|
||[height](/javascript/api/excel/excel.chartplotareadata#height)|Representa o valor de altura de plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotareadata#insideheight)|Representa o valor insideHeight plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotareadata#insideleft)|Representa o valor insideLeft de plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotareadata#insidetop)|Representa o valor insideTop de plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotareadata#insidewidth)|Representa o valor insideWidth de plotArea.|
||[left](/javascript/api/excel/excel.chartplotareadata#left)|Representa o valor de plotArea à esquerda.|
||[position](/javascript/api/excel/excel.chartplotareadata#position)|Represente a posição de plotArea.|
||[top](/javascript/api/excel/excel.chartplotareadata#top)|Representa o valor máximo de plotArea.|
||[width](/javascript/api/excel/excel.chartplotareadata#width)|Representa o valor de largura de plotArea.|
|[ChartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|[Borderô](/javascript/api/excel/excel.chartplotareaformat#border)|Representa os atributos de borda de um gráfico plotArea.|
||[fill](/javascript/api/excel/excel.chartplotareaformat#fill)|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo.|
||[Set (Propriedades: Excel. ChartPlotAreaFormat)](/javascript/api/excel/excel.chartplotareaformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartPlotAreaFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.chartplotareaformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[ChartPlotAreaFormatData](/javascript/api/excel/excel.chartplotareaformatdata)|[Borderô](/javascript/api/excel/excel.chartplotareaformatdata#border)|Representa os atributos de borda de um gráfico plotArea.|
|[ChartPlotAreaFormatLoadOptions](/javascript/api/excel/excel.chartplotareaformatloadoptions)|[$all](/javascript/api/excel/excel.chartplotareaformatloadoptions#$all)||
||[Borderô](/javascript/api/excel/excel.chartplotareaformatloadoptions#border)|Representa os atributos de borda de um gráfico plotArea.|
|[ChartPlotAreaFormatUpdateData](/javascript/api/excel/excel.chartplotareaformatupdatedata)|[Borderô](/javascript/api/excel/excel.chartplotareaformatupdatedata#border)|Representa os atributos de borda de um gráfico plotArea.|
|[ChartPlotAreaLoadOptions](/javascript/api/excel/excel.chartplotarealoadoptions)|[$all](/javascript/api/excel/excel.chartplotarealoadoptions#$all)||
||[format](/javascript/api/excel/excel.chartplotarealoadoptions#format)|Representa a formatação de um gráfico plotArea.|
||[height](/javascript/api/excel/excel.chartplotarealoadoptions#height)|Representa o valor de altura de plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotarealoadoptions#insideheight)|Representa o valor insideHeight plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotarealoadoptions#insideleft)|Representa o valor insideLeft de plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotarealoadoptions#insidetop)|Representa o valor insideTop de plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotarealoadoptions#insidewidth)|Representa o valor insideWidth de plotArea.|
||[left](/javascript/api/excel/excel.chartplotarealoadoptions#left)|Representa o valor de plotArea à esquerda.|
||[position](/javascript/api/excel/excel.chartplotarealoadoptions#position)|Represente a posição de plotArea.|
||[top](/javascript/api/excel/excel.chartplotarealoadoptions#top)|Representa o valor máximo de plotArea.|
||[width](/javascript/api/excel/excel.chartplotarealoadoptions#width)|Representa o valor de largura de plotArea.|
|[ChartPlotAreaUpdateData](/javascript/api/excel/excel.chartplotareaupdatedata)|[format](/javascript/api/excel/excel.chartplotareaupdatedata#format)|Representa a formatação de um gráfico plotArea.|
||[height](/javascript/api/excel/excel.chartplotareaupdatedata#height)|Representa o valor de altura de plotArea.|
||[insideHeight](/javascript/api/excel/excel.chartplotareaupdatedata#insideheight)|Representa o valor insideHeight plotArea.|
||[insideLeft](/javascript/api/excel/excel.chartplotareaupdatedata#insideleft)|Representa o valor insideLeft de plotArea.|
||[insideTop](/javascript/api/excel/excel.chartplotareaupdatedata#insidetop)|Representa o valor insideTop de plotArea.|
||[insideWidth](/javascript/api/excel/excel.chartplotareaupdatedata#insidewidth)|Representa o valor insideWidth de plotArea.|
||[left](/javascript/api/excel/excel.chartplotareaupdatedata#left)|Representa o valor de plotArea à esquerda.|
||[position](/javascript/api/excel/excel.chartplotareaupdatedata#position)|Represente a posição de plotArea.|
||[top](/javascript/api/excel/excel.chartplotareaupdatedata#top)|Representa o valor máximo de plotArea.|
||[width](/javascript/api/excel/excel.chartplotareaupdatedata#width)|Representa o valor de largura de plotArea.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[axisGroup](/javascript/api/excel/excel.chartseries#axisgroup)|Retorna ou define o grupo da série especificada. Leitura/gravação|
||[crescimento](/javascript/api/excel/excel.chartseries#explosion)|Retorna ou define o valor de explosão para um gráfico de pizza ou fatia de gráfico de rosca. Retorna 0 (zero) se não houver explosão (a ponta da fatia está no centro da pizza). Leitura/gravação.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseries#firstsliceangle)|Retorna ou define o ângulo do primeiro gráfico de pizza ou fatia de gráfico de rosca, em graus (no sentido horário a partir da vertical). Aplica-se apenas a pizza, torta 3-D e gráficos de rosca.. Pode ser um valor de 0 a 360. Leitura/gravação|
||[invertIfNegative](/javascript/api/excel/excel.chartseries#invertifnegative)|Verdadeiro se o Microsoft Excel inverte o padrão no item quando ele corresponde a um número negativo. Leitura/gravação.|
||[ficar](/javascript/api/excel/excel.chartseries#overlap)|Especifica como barras e colunas são posicionadas. Pode ser um valor entre – 100 e 100. Se aplicam apenas às barras 2D e gráficos de colunas 2D. Leitura/gravação.|
||[dataLabels](/javascript/api/excel/excel.chartseries#datalabels)|Representa uma coleção de todos os dataLabels da série.|
||[secondPlotSize](/javascript/api/excel/excel.chartseries#secondplotsize)|Retorna ou define o tamanho da seção secundária de uma pizza do gráfico de pizza ou de uma barra de gráfico de pizza, como uma porcentagem do tamanho da pizza primária. Pode ser um valor de 5 de 200. Leitura/gravação.|
||[splitType](/javascript/api/excel/excel.chartseries#splittype)|Retorna ou define a maneira como as duas seções de uma pizza do gráfico de pizza ou de uma barra do gráfico de pizza são divididas. Leitura/gravação.|
||[varyByCategories](/javascript/api/excel/excel.chartseries#varybycategories)|Verdadeiro se o Microsoft Excel atribuir uma cor ou padrão diferente a cada marcador de dados. O gráfico deve conter apenas uma série. Leitura/gravação.|
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[axisGroup](/javascript/api/excel/excel.chartseriescollectionloadoptions#axisgroup)|Para cada ITEM na coleção: Retorna ou define o grupo para a série especificada. Leitura/gravação|
||[dataLabels](/javascript/api/excel/excel.chartseriescollectionloadoptions#datalabels)|Para cada ITEM na coleção: representa uma coleção de todos os DataLabels da série.|
||[crescimento](/javascript/api/excel/excel.chartseriescollectionloadoptions#explosion)|Para cada ITEM na coleção: Retorna ou define o valor de explosão para uma fatia de gráfico de rosca ou de gráfico de pizza. Retorna 0 (zero) se não houver explosão (a ponta da fatia está no centro da pizza). Leitura/gravação.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriescollectionloadoptions#firstsliceangle)|Para cada ITEM na coleção: Retorna ou define o ângulo da primeira fatia do gráfico de pizza ou rosca, em graus (no sentido horário a partir da vertical). Aplica-se apenas a pizza, torta 3-D e gráficos de rosca.. Pode ser um valor de 0 a 360. Leitura/gravação|
||[invertIfNegative](/javascript/api/excel/excel.chartseriescollectionloadoptions#invertifnegative)|Para cada ITEM na coleção: true se o Microsoft Excel inverte o padrão no item quando ele corresponde a um número negativo. Leitura/gravação.|
||[ficar](/javascript/api/excel/excel.chartseriescollectionloadoptions#overlap)|Para cada ITEM na coleção: especifica como as barras e colunas são posicionadas. Pode ser um valor entre – 100 e 100. Se aplicam apenas às barras 2D e gráficos de colunas 2D. Leitura/gravação.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#secondplotsize)|Para cada ITEM na coleção: Retorna ou define o tamanho da seção secundária de uma pizza de gráfico de pizza ou uma barra de gráfico de pizza, como uma porcentagem do tamanho da pizza principal. Pode ser um valor de 5 de 200. Leitura/gravação.|
||[splitType](/javascript/api/excel/excel.chartseriescollectionloadoptions#splittype)|Para cada ITEM na coleção: Retorna ou define a maneira como as duas seções de uma pizza de gráfico de pizza ou uma barra de gráfico de pizza são divididas. Leitura/gravação.|
||[varyByCategories](/javascript/api/excel/excel.chartseriescollectionloadoptions#varybycategories)|Para cada ITEM na coleção: true se o Microsoft Excel atribuir uma cor ou padrão diferente para cada marcador de dados. O gráfico deve conter apenas uma série. Leitura/gravação.|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[axisGroup](/javascript/api/excel/excel.chartseriesdata#axisgroup)|Retorna ou define o grupo da série especificada. Leitura/gravação|
||[dataLabels](/javascript/api/excel/excel.chartseriesdata#datalabels)|Representa uma coleção de todos os dataLabels da série.|
||[crescimento](/javascript/api/excel/excel.chartseriesdata#explosion)|Retorna ou define o valor de explosão para um gráfico de pizza ou fatia de gráfico de rosca. Retorna 0 (zero) se não houver explosão (a ponta da fatia está no centro da pizza). Leitura/gravação.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesdata#firstsliceangle)|Retorna ou define o ângulo do primeiro gráfico de pizza ou fatia de gráfico de rosca, em graus (no sentido horário a partir da vertical). Aplica-se apenas a pizza, torta 3-D e gráficos de rosca.. Pode ser um valor de 0 a 360. Leitura/gravação|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesdata#invertifnegative)|Verdadeiro se o Microsoft Excel inverte o padrão no item quando ele corresponde a um número negativo. Leitura/gravação.|
||[ficar](/javascript/api/excel/excel.chartseriesdata#overlap)|Especifica como barras e colunas são posicionadas. Pode ser um valor entre – 100 e 100. Se aplicam apenas às barras 2D e gráficos de colunas 2D. Leitura/gravação.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesdata#secondplotsize)|Retorna ou define o tamanho da seção secundária de uma pizza do gráfico de pizza ou de uma barra de gráfico de pizza, como uma porcentagem do tamanho da pizza primária. Pode ser um valor de 5 de 200. Leitura/gravação.|
||[splitType](/javascript/api/excel/excel.chartseriesdata#splittype)|Retorna ou define a maneira como as duas seções de uma pizza do gráfico de pizza ou de uma barra do gráfico de pizza são divididas. Leitura/gravação.|
||[varyByCategories](/javascript/api/excel/excel.chartseriesdata#varybycategories)|Verdadeiro se o Microsoft Excel atribuir uma cor ou padrão diferente a cada marcador de dados. O gráfico deve conter apenas uma série. Leitura/gravação.|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[axisGroup](/javascript/api/excel/excel.chartseriesloadoptions#axisgroup)|Retorna ou define o grupo da série especificada. Leitura/gravação|
||[dataLabels](/javascript/api/excel/excel.chartseriesloadoptions#datalabels)|Representa uma coleção de todos os dataLabels da série.|
||[crescimento](/javascript/api/excel/excel.chartseriesloadoptions#explosion)|Retorna ou define o valor de explosão para um gráfico de pizza ou fatia de gráfico de rosca. Retorna 0 (zero) se não houver explosão (a ponta da fatia está no centro da pizza). Leitura/gravação.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesloadoptions#firstsliceangle)|Retorna ou define o ângulo do primeiro gráfico de pizza ou fatia de gráfico de rosca, em graus (no sentido horário a partir da vertical). Aplica-se apenas a pizza, torta 3-D e gráficos de rosca.. Pode ser um valor de 0 a 360. Leitura/gravação|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesloadoptions#invertifnegative)|Verdadeiro se o Microsoft Excel inverte o padrão no item quando ele corresponde a um número negativo. Leitura/gravação.|
||[ficar](/javascript/api/excel/excel.chartseriesloadoptions#overlap)|Especifica como barras e colunas são posicionadas. Pode ser um valor entre – 100 e 100. Se aplicam apenas às barras 2D e gráficos de colunas 2D. Leitura/gravação.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesloadoptions#secondplotsize)|Retorna ou define o tamanho da seção secundária de uma pizza do gráfico de pizza ou de uma barra de gráfico de pizza, como uma porcentagem do tamanho da pizza primária. Pode ser um valor de 5 de 200. Leitura/gravação.|
||[splitType](/javascript/api/excel/excel.chartseriesloadoptions#splittype)|Retorna ou define a maneira como as duas seções de uma pizza do gráfico de pizza ou de uma barra do gráfico de pizza são divididas. Leitura/gravação.|
||[varyByCategories](/javascript/api/excel/excel.chartseriesloadoptions#varybycategories)|Verdadeiro se o Microsoft Excel atribuir uma cor ou padrão diferente a cada marcador de dados. O gráfico deve conter apenas uma série. Leitura/gravação.|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[axisGroup](/javascript/api/excel/excel.chartseriesupdatedata#axisgroup)|Retorna ou define o grupo da série especificada. Leitura/gravação|
||[dataLabels](/javascript/api/excel/excel.chartseriesupdatedata#datalabels)|Representa uma coleção de todos os dataLabels da série.|
||[crescimento](/javascript/api/excel/excel.chartseriesupdatedata#explosion)|Retorna ou define o valor de explosão para um gráfico de pizza ou fatia de gráfico de rosca. Retorna 0 (zero) se não houver explosão (a ponta da fatia está no centro da pizza). Leitura/gravação.|
||[firstSliceAngle](/javascript/api/excel/excel.chartseriesupdatedata#firstsliceangle)|Retorna ou define o ângulo do primeiro gráfico de pizza ou fatia de gráfico de rosca, em graus (no sentido horário a partir da vertical). Aplica-se apenas a pizza, torta 3-D e gráficos de rosca.. Pode ser um valor de 0 a 360. Leitura/gravação|
||[invertIfNegative](/javascript/api/excel/excel.chartseriesupdatedata#invertifnegative)|Verdadeiro se o Microsoft Excel inverte o padrão no item quando ele corresponde a um número negativo. Leitura/gravação.|
||[ficar](/javascript/api/excel/excel.chartseriesupdatedata#overlap)|Especifica como barras e colunas são posicionadas. Pode ser um valor entre – 100 e 100. Se aplicam apenas às barras 2D e gráficos de colunas 2D. Leitura/gravação.|
||[secondPlotSize](/javascript/api/excel/excel.chartseriesupdatedata#secondplotsize)|Retorna ou define o tamanho da seção secundária de uma pizza do gráfico de pizza ou de uma barra de gráfico de pizza, como uma porcentagem do tamanho da pizza primária. Pode ser um valor de 5 de 200. Leitura/gravação.|
||[splitType](/javascript/api/excel/excel.chartseriesupdatedata#splittype)|Retorna ou define a maneira como as duas seções de uma pizza do gráfico de pizza ou de uma barra do gráfico de pizza são divididas. Leitura/gravação.|
||[varyByCategories](/javascript/api/excel/excel.chartseriesupdatedata#varybycategories)|Verdadeiro se o Microsoft Excel atribuir uma cor ou padrão diferente a cada marcador de dados. O gráfico deve conter apenas uma série. Leitura/gravação.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[backwardPeriod](/javascript/api/excel/excel.charttrendline#backwardperiod)|Representa o número de períodos que a linha de tendência se estende para trás.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendline#forwardperiod)|Representa o número de períodos que a linha de tendência se estende para frente.|
||[rótulo](/javascript/api/excel/excel.charttrendline#label)|Representa o rótulo de linha de tendência um gráfico.|
||[a equação](/javascript/api/excel/excel.charttrendline#showequation)|Verdadeiro se a equação da linha de tendência for exibida no gráfico.|
||[showRSquared](/javascript/api/excel/excel.charttrendline#showrsquared)|Verdadeiro se o R-quadrado da linha de tendência for exibido no gráfico.|
|[ChartTrendlineCollectionLoadOptions](/javascript/api/excel/excel.charttrendlinecollectionloadoptions)|[backwardPeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#backwardperiod)|Para cada ITEM na coleção: representa o número de períodos que a tendência se estende para trás.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#forwardperiod)|Para cada ITEM na coleção: representa o número de períodos que a tendência se estende para frente.|
||[rótulo](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#label)|Para cada ITEM na coleção: representa o rótulo de uma tendência de gráfico.|
||[a equação](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#showequation)|Para cada ITEM na coleção: true se a equação para a tendência é exibida no gráfico.|
||[showRSquared](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#showrsquared)|Para cada ITEM na coleção: true se o R-quadrado para a tendência é exibido no gráfico.|
|[ChartTrendlineData](/javascript/api/excel/excel.charttrendlinedata)|[backwardPeriod](/javascript/api/excel/excel.charttrendlinedata#backwardperiod)|Representa o número de períodos que a linha de tendência se estende para trás.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlinedata#forwardperiod)|Representa o número de períodos que a linha de tendência se estende para frente.|
||[rótulo](/javascript/api/excel/excel.charttrendlinedata#label)|Representa o rótulo de linha de tendência um gráfico.|
||[a equação](/javascript/api/excel/excel.charttrendlinedata#showequation)|Verdadeiro se a equação da linha de tendência for exibida no gráfico.|
||[showRSquared](/javascript/api/excel/excel.charttrendlinedata#showrsquared)|Verdadeiro se o R-quadrado da linha de tendência for exibido no gráfico.|
|[ChartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|[AutoTexto](/javascript/api/excel/excel.charttrendlinelabel#autotext)|Valor booliano que representa se o rótulo de linha de tendência gerará automaticamente o texto apropriado com base no contexto..|
||[formula](/javascript/api/excel/excel.charttrendlinelabel#formula)|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de linha de tendência usando a notação no estilo A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabel#horizontalalignment)|Representa o alinhamento horizontal de rótulo de linha de tendência de gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[left](/javascript/api/excel/excel.charttrendlinelabel#left)|Representa a distância, em pontos, da borda esquerda do rótulo da linha de tendência do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabel#numberformat)|Valor de cadeia de caracteres que representa o código do formato do rótulo de linha de tendência.|
||[format](/javascript/api/excel/excel.charttrendlinelabel#format)|Representa o formato do rótulo de linha de tendência de gráfico.|
||[height](/javascript/api/excel/excel.charttrendlinelabel#height)|Retorna a altura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[width](/javascript/api/excel/excel.charttrendlinelabel#width)|Retorna a largura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[Set (Propriedades: Excel. ChartTrendlineLabel)](/javascript/api/excel/excel.charttrendlinelabel#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartTrendlineLabelUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.charttrendlinelabel#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[text](/javascript/api/excel/excel.charttrendlinelabel#text)|Cadeia de caracteres que representa o texto do rótulo em um gráfico de linha de tendência.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabel#textorientation)|Representa a orientação de texto de rótulo de linha de tendência de gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.charttrendlinelabel#top)|Representa a distância, em pontos, da borda superior do rótulo de linha de tendência do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabel#verticalalignment)|Representa o alinhamento vertical do rótulo de linha de tendência de gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
|[ChartTrendlineLabelData](/javascript/api/excel/excel.charttrendlinelabeldata)|[AutoTexto](/javascript/api/excel/excel.charttrendlinelabeldata#autotext)|Valor booliano que representa se o rótulo de linha de tendência gerará automaticamente o texto apropriado com base no contexto..|
||[format](/javascript/api/excel/excel.charttrendlinelabeldata#format)|Representa o formato do rótulo de linha de tendência de gráfico.|
||[formula](/javascript/api/excel/excel.charttrendlinelabeldata#formula)|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de linha de tendência usando a notação no estilo A1.|
||[height](/javascript/api/excel/excel.charttrendlinelabeldata#height)|Retorna a altura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabeldata#horizontalalignment)|Representa o alinhamento horizontal de rótulo de linha de tendência de gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[left](/javascript/api/excel/excel.charttrendlinelabeldata#left)|Representa a distância, em pontos, da borda esquerda do rótulo da linha de tendência do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabeldata#numberformat)|Valor de cadeia de caracteres que representa o código do formato do rótulo de linha de tendência.|
||[text](/javascript/api/excel/excel.charttrendlinelabeldata#text)|Cadeia de caracteres que representa o texto do rótulo em um gráfico de linha de tendência.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabeldata#textorientation)|Representa a orientação de texto de rótulo de linha de tendência de gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.charttrendlinelabeldata#top)|Representa a distância, em pontos, da borda superior do rótulo de linha de tendência do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabeldata#verticalalignment)|Representa o alinhamento vertical do rótulo de linha de tendência de gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
||[width](/javascript/api/excel/excel.charttrendlinelabeldata#width)|Retorna a largura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
|[ChartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|[Borderô](/javascript/api/excel/excel.charttrendlinelabelformat#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
||[fill](/javascript/api/excel/excel.charttrendlinelabelformat#fill)|Representa o formato de preenchimento do rótulo de linha de tendência atual do gráfico.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformat#font)|Representa os atributos de fonte do rótulo de linha de tendência do gráfico, como nome, tamanho, cor, dentre outros.|
||[Set (Propriedades: Excel. ChartTrendlineLabelFormat)](/javascript/api/excel/excel.charttrendlinelabelformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartTrendlineLabelFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.charttrendlinelabelformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[ChartTrendlineLabelFormatData](/javascript/api/excel/excel.charttrendlinelabelformatdata)|[Borderô](/javascript/api/excel/excel.charttrendlinelabelformatdata#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatdata#font)|Representa os atributos de fonte do rótulo de linha de tendência do gráfico, como nome, tamanho, cor, dentre outros.|
|[ChartTrendlineLabelFormatLoadOptions](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#$all)||
||[Borderô](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatloadoptions#font)|Representa os atributos de fonte do rótulo de linha de tendência do gráfico, como nome, tamanho, cor, dentre outros.|
|[ChartTrendlineLabelFormatUpdateData](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata)|[Borderô](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata#border)|Representa o formato de borda, que inclui a espessura de cor e estilo de linha.|
||[font](/javascript/api/excel/excel.charttrendlinelabelformatupdatedata#font)|Representa os atributos de fonte do rótulo de linha de tendência do gráfico, como nome, tamanho, cor, dentre outros.|
|[ChartTrendlineLabelLoadOptions](/javascript/api/excel/excel.charttrendlinelabelloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinelabelloadoptions#$all)||
||[AutoTexto](/javascript/api/excel/excel.charttrendlinelabelloadoptions#autotext)|Valor booliano que representa se o rótulo de linha de tendência gerará automaticamente o texto apropriado com base no contexto..|
||[format](/javascript/api/excel/excel.charttrendlinelabelloadoptions#format)|Representa o formato do rótulo de linha de tendência de gráfico.|
||[formula](/javascript/api/excel/excel.charttrendlinelabelloadoptions#formula)|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de linha de tendência usando a notação no estilo A1.|
||[height](/javascript/api/excel/excel.charttrendlinelabelloadoptions#height)|Retorna a altura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabelloadoptions#horizontalalignment)|Representa o alinhamento horizontal de rótulo de linha de tendência de gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[left](/javascript/api/excel/excel.charttrendlinelabelloadoptions#left)|Representa a distância, em pontos, da borda esquerda do rótulo da linha de tendência do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabelloadoptions#numberformat)|Valor de cadeia de caracteres que representa o código do formato do rótulo de linha de tendência.|
||[text](/javascript/api/excel/excel.charttrendlinelabelloadoptions#text)|Cadeia de caracteres que representa o texto do rótulo em um gráfico de linha de tendência.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabelloadoptions#textorientation)|Representa a orientação de texto de rótulo de linha de tendência de gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.charttrendlinelabelloadoptions#top)|Representa a distância, em pontos, da borda superior do rótulo de linha de tendência do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabelloadoptions#verticalalignment)|Representa o alinhamento vertical do rótulo de linha de tendência de gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
||[width](/javascript/api/excel/excel.charttrendlinelabelloadoptions#width)|Retorna a largura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
|[ChartTrendlineLabelUpdateData](/javascript/api/excel/excel.charttrendlinelabelupdatedata)|[AutoTexto](/javascript/api/excel/excel.charttrendlinelabelupdatedata#autotext)|Valor booliano que representa se o rótulo de linha de tendência gerará automaticamente o texto apropriado com base no contexto..|
||[format](/javascript/api/excel/excel.charttrendlinelabelupdatedata#format)|Representa o formato do rótulo de linha de tendência de gráfico.|
||[formula](/javascript/api/excel/excel.charttrendlinelabelupdatedata#formula)|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de linha de tendência usando a notação no estilo A1.|
||[horizontalAlignment](/javascript/api/excel/excel.charttrendlinelabelupdatedata#horizontalalignment)|Representa o alinhamento horizontal de rótulo de linha de tendência de gráfico. Consulte Excel. ChartTextHorizontalAlignment para obter detalhes.|
||[left](/javascript/api/excel/excel.charttrendlinelabelupdatedata#left)|Representa a distância, em pontos, da borda esquerda do rótulo da linha de tendência do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[numberFormat](/javascript/api/excel/excel.charttrendlinelabelupdatedata#numberformat)|Valor de cadeia de caracteres que representa o código do formato do rótulo de linha de tendência.|
||[text](/javascript/api/excel/excel.charttrendlinelabelupdatedata#text)|Cadeia de caracteres que representa o texto do rótulo em um gráfico de linha de tendência.|
||[textOrientation](/javascript/api/excel/excel.charttrendlinelabelupdatedata#textorientation)|Representa a orientação de texto de rótulo de linha de tendência de gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.charttrendlinelabelupdatedata#top)|Representa a distância, em pontos, da borda superior do rótulo de linha de tendência do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.charttrendlinelabelupdatedata#verticalalignment)|Representa o alinhamento vertical do rótulo de linha de tendência de gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
|[ChartTrendlineLoadOptions](/javascript/api/excel/excel.charttrendlineloadoptions)|[backwardPeriod](/javascript/api/excel/excel.charttrendlineloadoptions#backwardperiod)|Representa o número de períodos que a linha de tendência se estende para trás.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlineloadoptions#forwardperiod)|Representa o número de períodos que a linha de tendência se estende para frente.|
||[rótulo](/javascript/api/excel/excel.charttrendlineloadoptions#label)|Representa o rótulo de linha de tendência um gráfico.|
||[a equação](/javascript/api/excel/excel.charttrendlineloadoptions#showequation)|Verdadeiro se a equação da linha de tendência for exibida no gráfico.|
||[showRSquared](/javascript/api/excel/excel.charttrendlineloadoptions#showrsquared)|Verdadeiro se o R-quadrado da linha de tendência for exibido no gráfico.|
|[ChartTrendlineUpdateData](/javascript/api/excel/excel.charttrendlineupdatedata)|[backwardPeriod](/javascript/api/excel/excel.charttrendlineupdatedata#backwardperiod)|Representa o número de períodos que a linha de tendência se estende para trás.|
||[forwardPeriod](/javascript/api/excel/excel.charttrendlineupdatedata#forwardperiod)|Representa o número de períodos que a linha de tendência se estende para frente.|
||[rótulo](/javascript/api/excel/excel.charttrendlineupdatedata#label)|Representa o rótulo de linha de tendência um gráfico.|
||[a equação](/javascript/api/excel/excel.charttrendlineupdatedata#showequation)|Verdadeiro se a equação da linha de tendência for exibida no gráfico.|
||[showRSquared](/javascript/api/excel/excel.charttrendlineupdatedata#showrsquared)|Verdadeiro se o R-quadrado da linha de tendência for exibido no gráfico.|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[categoryLabelLevel](/javascript/api/excel/excel.chartupdatedata#categorylabellevel)|Retorna ou define uma constante de enumeração ChartCategoryLabelLevel que se refere a|
||[displayBlanksAs](/javascript/api/excel/excel.chartupdatedata#displayblanksas)|Retorna ou define a maneira como as células em branco são plotadas em um gráfico. Leitura/gravação.|
||[plotArea](/javascript/api/excel/excel.chartupdatedata#plotarea)|Representa a plotArea para o gráfico.|
||[plotBy](/javascript/api/excel/excel.chartupdatedata#plotby)|Retorna ou define como as colunas ou linhas são usadas como séries de dados no gráfico. Leitura/gravação.|
||[plotVisibleOnly](/javascript/api/excel/excel.chartupdatedata#plotvisibleonly)|Verdadeiro se apenas as células visíveis forem plotadas.Falso se ambas as células visíveis e ocultas forem plotadas.. Leitura/gravação.|
||[seriesNameLevel](/javascript/api/excel/excel.chartupdatedata#seriesnamelevel)|Retorna ou define uma constante de enumeração ChartSeriesNameLevel que se refere a|
||[showDataLabelsOverMaximum](/javascript/api/excel/excel.chartupdatedata#showdatalabelsovermaximum)|Representa se os rótulos de dados devem ser mostrados quando o valor for maior que o valor máximo no eixo de valor.|
||[style](/javascript/api/excel/excel.chartupdatedata#style)|Retorna ou define o estilo do gráfico para o gráfico. Leitura/gravação.|
|[CustomDataValidation](/javascript/api/excel/excel.customdatavalidation)|[formula](/javascript/api/excel/excel.customdatavalidation#formula)|Uma fórmula de validação de dados personalizados. Isso cria regras de entrada especiais, como impedir duplicatas ou limitar o total em um intervalo de células.|
|[DataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|[name](/javascript/api/excel/excel.datapivothierarchy#name)|Nome da DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchy#numberformat)|Formato de número do DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchy#position)|Posição da DataPivotHierarchy.|
||[campo](/javascript/api/excel/excel.datapivothierarchy#field)|Retorna PivotFields associados a DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchy#id)|ID do DataPivotHierarchy.|
||[Set (Propriedades: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchy#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. DataPivotHierarchyUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.datapivothierarchy#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Redefina a DataPivotHierarchy para os valores padrão.|
||[showAs](/javascript/api/excel/excel.datapivothierarchy#showas)|Determina se os dados devem ser mostrados como um cálculo de resumo específico ou não.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchy#summarizeby)|Determina se deve mostrar todos os itens a DataPivotHierarchy.|
|[DataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|[Adicionar (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#add-pivothierarchy-)|Adiciona o PivotHierarchy ao eixo atual.|
||[getCount()](/javascript/api/excel/excel.datapivothierarchycollection#getcount--)|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitem-name-)|Obtém DataPivotHierarchy por nome ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection#getitemornullobject-name-)|Obtém uma DataPivotHierarchy por nome. Se o DataPivotHierarchy não existir, retornará um objeto nulo.|
||[items](/javascript/api/excel/excel.datapivothierarchycollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[remover (DataPivotHierarchy: Excel. DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection#remove-datapivothierarchy-)|Remove o PivotHierarchy do eixo atual.|
|[DataPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#$all)||
||[campo](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#field)|Para cada ITEM na coleção: retorna o PivotFields associado ao DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#id)|Para cada ITEM da coleção: ID do DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#name)|Para cada ITEM na coleção: nome da DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#numberformat)|Para cada ITEM na coleção: o formato de número do DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#position)|Para cada ITEM da coleção: posição do DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#showas)|Para cada ITEM na coleção: determina se os dados devem ser mostrados como um cálculo de resumo específico ou não.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchycollectionloadoptions#summarizeby)|Para cada ITEM na coleção: determina se todos os itens do DataPivotHierarchy devem ser exibidos.|
|[DataPivotHierarchyData](/javascript/api/excel/excel.datapivothierarchydata)|[campo](/javascript/api/excel/excel.datapivothierarchydata#field)|Retorna PivotFields associados a DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchydata#id)|ID do DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchydata#name)|Nome da DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchydata#numberformat)|Formato de número do DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchydata#position)|Posição da DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchydata#showas)|Determina se os dados devem ser mostrados como um cálculo de resumo específico ou não.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchydata#summarizeby)|Determina se deve mostrar todos os itens a DataPivotHierarchy.|
|[DataPivotHierarchyLoadOptions](/javascript/api/excel/excel.datapivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.datapivothierarchyloadoptions#$all)||
||[campo](/javascript/api/excel/excel.datapivothierarchyloadoptions#field)|Retorna PivotFields associados a DataPivotHierarchy.|
||[id](/javascript/api/excel/excel.datapivothierarchyloadoptions#id)|ID do DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchyloadoptions#name)|Nome da DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchyloadoptions#numberformat)|Formato de número do DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchyloadoptions#position)|Posição da DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchyloadoptions#showas)|Determina se os dados devem ser mostrados como um cálculo de resumo específico ou não.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchyloadoptions#summarizeby)|Determina se deve mostrar todos os itens a DataPivotHierarchy.|
|[DataPivotHierarchyUpdateData](/javascript/api/excel/excel.datapivothierarchyupdatedata)|[campo](/javascript/api/excel/excel.datapivothierarchyupdatedata#field)|Retorna PivotFields associados a DataPivotHierarchy.|
||[name](/javascript/api/excel/excel.datapivothierarchyupdatedata#name)|Nome da DataPivotHierarchy.|
||[numberFormat](/javascript/api/excel/excel.datapivothierarchyupdatedata#numberformat)|Formato de número do DataPivotHierarchy.|
||[position](/javascript/api/excel/excel.datapivothierarchyupdatedata#position)|Posição da DataPivotHierarchy.|
||[showAs](/javascript/api/excel/excel.datapivothierarchyupdatedata#showas)|Determina se os dados devem ser mostrados como um cálculo de resumo específico ou não.|
||[summarizeBy](/javascript/api/excel/excel.datapivothierarchyupdatedata#summarizeby)|Determina se deve mostrar todos os itens a DataPivotHierarchy.|
|[DataValidation](/javascript/api/excel/excel.datavalidation)|[clear()](/javascript/api/excel/excel.datavalidation#clear--)|Desfazer a validação de dados do intervalo atual.|
||[errorAlert](/javascript/api/excel/excel.datavalidation#erroralert)|Alerta de erro quando o usuário insere dados inválidos.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidation#ignoreblanks)|Ignora espaços em branco: nenhuma validação de dados será executada em células vazias, o padrão será definido como verdadeiro.|
||[prompt](/javascript/api/excel/excel.datavalidation#prompt)|Avisar quando os usuários selecionarem uma célula.|
||[tipo](/javascript/api/excel/excel.datavalidation#type)|Tipo de validação de dados, confira Excel.DataValidationType para obter detalhes.|
||[inválido](/javascript/api/excel/excel.datavalidation#valid)|Representa se todos os valores de célula são válidos de acordo com as regras de validação de dados.|
||[norma](/javascript/api/excel/excel.datavalidation#rule)|Regra de validação de dados que contém diferentes tipos de critérios de validação de dados.|
||[Set (Propriedades: Excel. DataValidation)](/javascript/api/excel/excel.datavalidation#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. DataValidationUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.datavalidation#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[DataValidationData](/javascript/api/excel/excel.datavalidationdata)|[errorAlert](/javascript/api/excel/excel.datavalidationdata#erroralert)|Alerta de erro quando o usuário insere dados inválidos.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationdata#ignoreblanks)|Ignora espaços em branco: nenhuma validação de dados será executada em células vazias, o padrão será definido como verdadeiro.|
||[prompt](/javascript/api/excel/excel.datavalidationdata#prompt)|Avisar quando os usuários selecionarem uma célula.|
||[norma](/javascript/api/excel/excel.datavalidationdata#rule)|Regra de validação de dados que contém diferentes tipos de critérios de validação de dados.|
||[tipo](/javascript/api/excel/excel.datavalidationdata#type)|Tipo de validação de dados, confira Excel.DataValidationType para obter detalhes.|
||[inválido](/javascript/api/excel/excel.datavalidationdata#valid)|Representa se todos os valores de célula são válidos de acordo com as regras de validação de dados.|
|[DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|[message](/javascript/api/excel/excel.datavalidationerroralert#message)|Representa a mensagem de alerta de erro.|
||[Enviar alerta](/javascript/api/excel/excel.datavalidationerroralert#showalert)|Determina se deseja mostrar uma caixa de diálogo de alerta de erro ou não quando um usuário insere dados inválidos. O padrão é verdadeiro.|
||[style](/javascript/api/excel/excel.datavalidationerroralert#style)|Representa o tipo de alerta de validação de dados, confira Excel.DataValidationAlertStyle para obter detalhes.|
||[title](/javascript/api/excel/excel.datavalidationerroralert#title)|Representa o título da caixa de diálogo de alerta de erro.|
|[DataValidationLoadOptions](/javascript/api/excel/excel.datavalidationloadoptions)|[$all](/javascript/api/excel/excel.datavalidationloadoptions#$all)||
||[errorAlert](/javascript/api/excel/excel.datavalidationloadoptions#erroralert)|Alerta de erro quando o usuário insere dados inválidos.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationloadoptions#ignoreblanks)|Ignora espaços em branco: nenhuma validação de dados será executada em células vazias, o padrão será definido como verdadeiro.|
||[prompt](/javascript/api/excel/excel.datavalidationloadoptions#prompt)|Avisar quando os usuários selecionarem uma célula.|
||[norma](/javascript/api/excel/excel.datavalidationloadoptions#rule)|Regra de validação de dados que contém diferentes tipos de critérios de validação de dados.|
||[tipo](/javascript/api/excel/excel.datavalidationloadoptions#type)|Tipo de validação de dados, confira Excel.DataValidationType para obter detalhes.|
||[inválido](/javascript/api/excel/excel.datavalidationloadoptions#valid)|Representa se todos os valores de célula são válidos de acordo com as regras de validação de dados.|
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
|[DataValidationUpdateData](/javascript/api/excel/excel.datavalidationupdatedata)|[errorAlert](/javascript/api/excel/excel.datavalidationupdatedata#erroralert)|Alerta de erro quando o usuário insere dados inválidos.|
||[ignoreBlanks](/javascript/api/excel/excel.datavalidationupdatedata#ignoreblanks)|Ignora espaços em branco: nenhuma validação de dados será executada em células vazias, o padrão será definido como verdadeiro.|
||[prompt](/javascript/api/excel/excel.datavalidationupdatedata#prompt)|Avisar quando os usuários selecionarem uma célula.|
||[norma](/javascript/api/excel/excel.datavalidationupdatedata#rule)|Regra de validação de dados que contém diferentes tipos de critérios de validação de dados.|
|[Datetimedatavalidationcomo](/javascript/api/excel/excel.datetimedatavalidation)|[Formula1](/javascript/api/excel/excel.datetimedatavalidation#formula1)|Especifica o operando direito quando a Propriedade Operator é definida como um operador binário como GreaterThan (o operando esquerdo é o valor que o usuário tenta inserir na célula). Com os operadores ternários between e não between, especifica o operando de limite inferior.|
||[Formula2](/javascript/api/excel/excel.datetimedatavalidation#formula2)|Com os operadores ternários between e não between, especifica o operando de limite superior. Não é usado com os operadores binários, como GreaterThan.|
||[operador](/javascript/api/excel/excel.datetimedatavalidation#operator)|O operador a ser usado para validar os dados.|
|[FilterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchy#enablemultiplefilteritems)|Determina se deseja permitir vários itens de filtro.|
||[name](/javascript/api/excel/excel.filterpivothierarchy#name)|Nome do FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchy#position)|Posição do FilterPivotHierarchy.|
||[fields](/javascript/api/excel/excel.filterpivothierarchy#fields)|Retorna PivotFields associados a FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchy#id)|ID do FilterPivotHierarchy.|
||[Set (Propriedades: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchy#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. FilterPivotHierarchyUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.filterpivothierarchy#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[setToDefault()](/javascript/api/excel/excel.filterpivothierarchy#settodefault--)|Redefina a FilterPivotHierarchy para os valores padrão.|
|[FilterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|[Adicionar (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#add-pivothierarchy-)|Adiciona o PivotHierarchy ao eixo atual. Se houver a hierarquia em outro lugar na linha, coluna,|
||[getCount()](/javascript/api/excel/excel.filterpivothierarchycollection#getcount--)|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitem-name-)|Obtém FilterPivotHierarchy por nome ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection#getitemornullobject-name-)|Obtém um FilterPivotHierarchy por nome. Se o FilterPivotHierarchy não existir, retornará um objeto nulo.|
||[items](/javascript/api/excel/excel.filterpivothierarchycollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[remover (filterPivotHierarchy: Excel. FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection#remove-filterpivothierarchy-)|Remove o PivotHierarchy do eixo atual.|
|[FilterPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#$all)||
||[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#enablemultiplefilteritems)|Para cada ITEM na coleção: determina se é para permitir vários itens de filtro.|
||[id](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#id)|Para cada ITEM da coleção: ID do FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#name)|Para cada ITEM na coleção: nome da FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchycollectionloadoptions#position)|Para cada ITEM da coleção: posição do FilterPivotHierarchy.|
|[FilterPivotHierarchyData](/javascript/api/excel/excel.filterpivothierarchydata)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchydata#enablemultiplefilteritems)|Determina se deseja permitir vários itens de filtro.|
||[fields](/javascript/api/excel/excel.filterpivothierarchydata#fields)|Retorna PivotFields associados a FilterPivotHierarchy.|
||[id](/javascript/api/excel/excel.filterpivothierarchydata#id)|ID do FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchydata#name)|Nome do FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchydata#position)|Posição do FilterPivotHierarchy.|
|[FilterPivotHierarchyLoadOptions](/javascript/api/excel/excel.filterpivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.filterpivothierarchyloadoptions#$all)||
||[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchyloadoptions#enablemultiplefilteritems)|Determina se deseja permitir vários itens de filtro.|
||[id](/javascript/api/excel/excel.filterpivothierarchyloadoptions#id)|ID do FilterPivotHierarchy.|
||[name](/javascript/api/excel/excel.filterpivothierarchyloadoptions#name)|Nome do FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchyloadoptions#position)|Posição do FilterPivotHierarchy.|
|[FilterPivotHierarchyUpdateData](/javascript/api/excel/excel.filterpivothierarchyupdatedata)|[enableMultipleFilterItems](/javascript/api/excel/excel.filterpivothierarchyupdatedata#enablemultiplefilteritems)|Determina se deseja permitir vários itens de filtro.|
||[name](/javascript/api/excel/excel.filterpivothierarchyupdatedata#name)|Nome do FilterPivotHierarchy.|
||[position](/javascript/api/excel/excel.filterpivothierarchyupdatedata#position)|Posição do FilterPivotHierarchy.|
|[ListDataValidation](/javascript/api/excel/excel.listdatavalidation)|[inCellDropDown](/javascript/api/excel/excel.listdatavalidation#incelldropdown)|Exibido na lista na célula suspensa ou não, ele será padronizado como verdadeiro.|
||[source](/javascript/api/excel/excel.listdatavalidation#source)|Fonte da lista de validação de dados|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[name](/javascript/api/excel/excel.pivotfield#name)|Nome do PivotField.|
||[id](/javascript/api/excel/excel.pivotfield#id)|ID do PivotField..|
||[items](/javascript/api/excel/excel.pivotfield#items)|Retorna o PivotItems que é composto com o PivotField.|
||[Set (Propriedades: Excel. PivotField)](/javascript/api/excel/excel.pivotfield#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. PivotFieldUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.pivotfield#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[showAllItems](/javascript/api/excel/excel.pivotfield#showallitems)|Determina se deseja mostrar todos os itens de PivotField.|
||[sortByLabels (sortBy: SortBy)](/javascript/api/excel/excel.pivotfield#sortbylabels-sortby-)|Classifica o PivotField. Se um DataPivotHierarchy for especificado, a classificação será aplicada com base nele, se a classificação não for baseada no campo PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfield#subtotals)|Subtotais de PivotField.|
|[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|[getCount()](/javascript/api/excel/excel.pivotfieldcollection#getcount--)|Obtém o número de campos de tabela dinâmica na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitem-name-)|Obtém um PivotField por nome ou ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection#getitemornullobject-name-)|Obtém um PivotField pelo nome. Se PivotField não existir, retornará um objeto NULL.|
||[items](/javascript/api/excel/excel.pivotfieldcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotFieldCollectionLoadOptions](/javascript/api/excel/excel.pivotfieldcollectionloadoptions)|[$all](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#id)|Para cada ITEM na coleção: ID do PivotField.|
||[name](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#name)|Para cada ITEM na coleção: nome do PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#showallitems)|Para cada ITEM na coleção: determina se todos os itens do PivotField serão mostrados.|
||[subtotals](/javascript/api/excel/excel.pivotfieldcollectionloadoptions#subtotals)|Para cada ITEM na coleção: subtotais do PivotField.|
|[PivotFieldData](/javascript/api/excel/excel.pivotfielddata)|[id](/javascript/api/excel/excel.pivotfielddata#id)|ID do PivotField..|
||[items](/javascript/api/excel/excel.pivotfielddata#items)|Retorna PivotFields associados ao PivotField.|
||[name](/javascript/api/excel/excel.pivotfielddata#name)|Nome do PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfielddata#showallitems)|Determina se deseja mostrar todos os itens de PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfielddata#subtotals)|Subtotais de PivotField.|
|[PivotFieldLoadOptions](/javascript/api/excel/excel.pivotfieldloadoptions)|[$all](/javascript/api/excel/excel.pivotfieldloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotfieldloadoptions#id)|ID do PivotField..|
||[name](/javascript/api/excel/excel.pivotfieldloadoptions#name)|Nome do PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfieldloadoptions#showallitems)|Determina se deseja mostrar todos os itens de PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfieldloadoptions#subtotals)|Subtotais de PivotField.|
|[PivotFieldUpdateData](/javascript/api/excel/excel.pivotfieldupdatedata)|[name](/javascript/api/excel/excel.pivotfieldupdatedata#name)|Nome do PivotField.|
||[showAllItems](/javascript/api/excel/excel.pivotfieldupdatedata#showallitems)|Determina se deseja mostrar todos os itens de PivotField.|
||[subtotals](/javascript/api/excel/excel.pivotfieldupdatedata#subtotals)|Subtotais de PivotField.|
|[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|[name](/javascript/api/excel/excel.pivothierarchy#name)|Nome do PivotHierarchy.|
||[fields](/javascript/api/excel/excel.pivothierarchy#fields)|Retorna PivotFields associados a PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchy#id)|ID do PivotHierarchy.|
||[Set (Propriedades: Excel. PivotHierarchy)](/javascript/api/excel/excel.pivothierarchy#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. PivotHierarchyUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.pivothierarchy#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|[getCount()](/javascript/api/excel/excel.pivothierarchycollection#getcount--)|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitem-name-)|Obtém PivotHierarchy por nome ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection#getitemornullobject-name-)|Obtém o PivotHierarchy por nome. Se o PivotHierarchy não existir, retornará um objeto null.|
||[items](/javascript/api/excel/excel.pivothierarchycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.pivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#id)|Para cada ITEM da coleção: ID do PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchycollectionloadoptions#name)|Para cada ITEM na coleção: nome da PivotHierarchy.|
|[PivotHierarchyData](/javascript/api/excel/excel.pivothierarchydata)|[fields](/javascript/api/excel/excel.pivothierarchydata#fields)|Retorna PivotFields associados a PivotHierarchy.|
||[id](/javascript/api/excel/excel.pivothierarchydata#id)|ID do PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchydata#name)|Nome do PivotHierarchy.|
|[PivotHierarchyLoadOptions](/javascript/api/excel/excel.pivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.pivothierarchyloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivothierarchyloadoptions#id)|ID do PivotHierarchy.|
||[name](/javascript/api/excel/excel.pivothierarchyloadoptions#name)|Nome do PivotHierarchy.|
|[PivotHierarchyUpdateData](/javascript/api/excel/excel.pivothierarchyupdatedata)|[name](/javascript/api/excel/excel.pivothierarchyupdatedata#name)|Nome do PivotHierarchy.|
|[PivotItem](/javascript/api/excel/excel.pivotitem)|[isExpanded](/javascript/api/excel/excel.pivotitem#isexpanded)|Determina se o item está expandido para mostrar itens filho ou se ele está recolhido e os itens filho estão ocultos.|
||[name](/javascript/api/excel/excel.pivotitem#name)|Nome do PivotItem.|
||[id](/javascript/api/excel/excel.pivotitem#id)|ID do PivotItem.|
||[Set (Propriedades: Excel. PivotItem)](/javascript/api/excel/excel.pivotitem#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. PivotItemUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.pivotitem#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[visible](/javascript/api/excel/excel.pivotitem#visible)|Determina se o PivotItem ficará visível ou não.|
|[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|[getCount()](/javascript/api/excel/excel.pivotitemcollection#getcount--)|Obtém o número de itens de tabela dinâmica na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitem-name-)|Obtém um PivotItem por seu nome ou ID.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection#getitemornullobject-name-)|Obtém um PivotItem pelo nome. Se o PivotItem não existir, retornará um objeto NULL.|
||[items](/javascript/api/excel/excel.pivotitemcollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[PivotItemCollectionLoadOptions](/javascript/api/excel/excel.pivotitemcollectionloadoptions)|[$all](/javascript/api/excel/excel.pivotitemcollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotitemcollectionloadoptions#id)|Para cada ITEM na coleção: ID do PivotItem.|
||[isExpanded](/javascript/api/excel/excel.pivotitemcollectionloadoptions#isexpanded)|Para cada ITEM na coleção: determina se o item é expandido para mostrar itens filhos ou se é recolhido e se os itens filhos estão ocultos.|
||[name](/javascript/api/excel/excel.pivotitemcollectionloadoptions#name)|Para cada ITEM na coleção: o nome do PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemcollectionloadoptions#visible)|Para cada ITEM na coleção: determina se o PivotItem está visível ou não.|
|[PivotItemData](/javascript/api/excel/excel.pivotitemdata)|[id](/javascript/api/excel/excel.pivotitemdata#id)|ID do PivotItem.|
||[isExpanded](/javascript/api/excel/excel.pivotitemdata#isexpanded)|Determina se o item está expandido para mostrar itens filho ou se ele está recolhido e os itens filho estão ocultos.|
||[name](/javascript/api/excel/excel.pivotitemdata#name)|Nome do PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemdata#visible)|Determina se o PivotItem ficará visível ou não.|
|[PivotItemLoadOptions](/javascript/api/excel/excel.pivotitemloadoptions)|[$all](/javascript/api/excel/excel.pivotitemloadoptions#$all)||
||[id](/javascript/api/excel/excel.pivotitemloadoptions#id)|ID do PivotItem.|
||[isExpanded](/javascript/api/excel/excel.pivotitemloadoptions#isexpanded)|Determina se o item está expandido para mostrar itens filho ou se ele está recolhido e os itens filho estão ocultos.|
||[name](/javascript/api/excel/excel.pivotitemloadoptions#name)|Nome do PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemloadoptions#visible)|Determina se o PivotItem ficará visível ou não.|
|[PivotItemUpdateData](/javascript/api/excel/excel.pivotitemupdatedata)|[isExpanded](/javascript/api/excel/excel.pivotitemupdatedata#isexpanded)|Determina se o item está expandido para mostrar itens filho ou se ele está recolhido e os itens filho estão ocultos.|
||[name](/javascript/api/excel/excel.pivotitemupdatedata#name)|Nome do PivotItem.|
||[visible](/javascript/api/excel/excel.pivotitemupdatedata#visible)|Determina se o PivotItem ficará visível ou não.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout#getcolumnlabelrange--)|Retorna o intervalo onde residem os rótulos de coluna da Tabela Dinâmica.|
||[getDataBodyRange()](/javascript/api/excel/excel.pivotlayout#getdatabodyrange--)|Retorna o intervalo onde residem os valores de dados da tabela dinâmica.|
||[getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout#getfilteraxisrange--)|Retorna o intervalo de área de filtro da Tabela Dinâmica.|
||[getRange()](/javascript/api/excel/excel.pivotlayout#getrange--)|Retorna o intervalo em que a Tabela Dinâmica existe, excluindo a área de filtro.|
||[getRowLabelRange()](/javascript/api/excel/excel.pivotlayout#getrowlabelrange--)|Retorna o intervalo onde residem os rótulos de linha da Tabela Dinâmica.|
||[LayoutType](/javascript/api/excel/excel.pivotlayout#layouttype)|Essa propriedade indica o PivotLayoutType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo.|
||[Set (Propriedades: Excel. PivotLayout)](/javascript/api/excel/excel.pivotlayout#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. PivotLayoutUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.pivotlayout#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayout#showcolumngrandtotals)|Especifica se o relatório de tabela dinâmica mostra os totais gerais das colunas.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayout#showrowgrandtotals)|Especifica se o relatório de tabela dinâmica mostra os totais gerais para as linhas.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayout#subtotallocation)|Essa propriedade indica SubtotalLocationType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo.|
|[PivotLayoutData](/javascript/api/excel/excel.pivotlayoutdata)|[LayoutType](/javascript/api/excel/excel.pivotlayoutdata#layouttype)|Essa propriedade indica o PivotLayoutType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutdata#showcolumngrandtotals)|Especifica se o relatório de tabela dinâmica mostra os totais gerais das colunas.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutdata#showrowgrandtotals)|Especifica se o relatório de tabela dinâmica mostra os totais gerais para as linhas.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutdata#subtotallocation)|Essa propriedade indica SubtotalLocationType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo.|
|[PivotLayoutLoadOptions](/javascript/api/excel/excel.pivotlayoutloadoptions)|[$all](/javascript/api/excel/excel.pivotlayoutloadoptions#$all)||
||[LayoutType](/javascript/api/excel/excel.pivotlayoutloadoptions#layouttype)|Essa propriedade indica o PivotLayoutType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutloadoptions#showcolumngrandtotals)|Especifica se o relatório de tabela dinâmica mostra os totais gerais das colunas.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutloadoptions#showrowgrandtotals)|Especifica se o relatório de tabela dinâmica mostra os totais gerais para as linhas.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutloadoptions#subtotallocation)|Essa propriedade indica SubtotalLocationType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo.|
|[PivotLayoutUpdateData](/javascript/api/excel/excel.pivotlayoutupdatedata)|[LayoutType](/javascript/api/excel/excel.pivotlayoutupdatedata#layouttype)|Essa propriedade indica o PivotLayoutType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo.|
||[showColumnGrandTotals](/javascript/api/excel/excel.pivotlayoutupdatedata#showcolumngrandtotals)|Especifica se o relatório de tabela dinâmica mostra os totais gerais das colunas.|
||[showRowGrandTotals](/javascript/api/excel/excel.pivotlayoutupdatedata#showrowgrandtotals)|Especifica se o relatório de tabela dinâmica mostra os totais gerais para as linhas.|
||[subtotalLocation](/javascript/api/excel/excel.pivotlayoutupdatedata#subtotallocation)|Essa propriedade indica SubtotalLocationType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[delete()](/javascript/api/excel/excel.pivottable#delete--)|Exclui a Tabela Dinâmica.|
||[columnHierarchies](/javascript/api/excel/excel.pivottable#columnhierarchies)|As hierarquias de pivô da coluna da Tabela Dinâmica.|
||[datahierarquias](/javascript/api/excel/excel.pivottable#datahierarchies)|As hierarquias dinâmicas de dados da Tabela Dinâmica.|
||[filterHierarchies](/javascript/api/excel/excel.pivottable#filterhierarchies)|As hierarquias de pivô do filtro da Tabela Dinâmica.|
||[hierarquias](/javascript/api/excel/excel.pivottable#hierarchies)|Hierarquias pivô da Tabela Dinâmica.|
||[teclado](/javascript/api/excel/excel.pivottable#layout)|O PivotLayout descreve o layout e estrutura visual da Tabela Dinâmica.|
||[transhierarquias](/javascript/api/excel/excel.pivottable#rowhierarchies)|As hierarquias de pivô de linha da Tabela Dinâmica.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[Add (Name: String, Source: tabela \| de \| cadeia de caracteres de intervalo \| , destino: cadeia de caracteres de intervalo)](/javascript/api/excel/excel.pivottablecollection#add-name--source--destination-)|Adiciona um Pivottable com base nos dados de origem especificados e insere-o na célula superior esquerda do intervalo de destino.|
|[PivotTableCollectionLoadOptions](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[teclado](/javascript/api/excel/excel.pivottablecollectionloadoptions#layout)|Para cada ITEM na coleção: o PivotLayout que descreve o layout e a estrutura visual da tabela dinâmica.|
|[PivotTableData](/javascript/api/excel/excel.pivottabledata)|[columnHierarchies](/javascript/api/excel/excel.pivottabledata#columnhierarchies)|As hierarquias de pivô da coluna da Tabela Dinâmica.|
||[datahierarquias](/javascript/api/excel/excel.pivottabledata#datahierarchies)|As hierarquias dinâmicas de dados da Tabela Dinâmica.|
||[filterHierarchies](/javascript/api/excel/excel.pivottabledata#filterhierarchies)|As hierarquias de pivô do filtro da Tabela Dinâmica.|
||[hierarquias](/javascript/api/excel/excel.pivottabledata#hierarchies)|Hierarquias pivô da Tabela Dinâmica.|
||[transhierarquias](/javascript/api/excel/excel.pivottabledata#rowhierarchies)|As hierarquias de pivô de linha da Tabela Dinâmica.|
|[PivotTableLoadOptions](/javascript/api/excel/excel.pivottableloadoptions)|[teclado](/javascript/api/excel/excel.pivottableloadoptions#layout)|O PivotLayout descreve o layout e estrutura visual da Tabela Dinâmica.|
|[Range](/javascript/api/excel/excel.range)|[dataValidation](/javascript/api/excel/excel.range#datavalidation)|Retorna um objeto de validação de dados.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[dataValidation](/javascript/api/excel/excel.rangedata#datavalidation)|Retorna um objeto de validação de dados.|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[dataValidation](/javascript/api/excel/excel.rangeloadoptions#datavalidation)|Retorna um objeto de validação de dados.|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[dataValidation](/javascript/api/excel/excel.rangeupdatedata#datavalidation)|Retorna um objeto de validação de dados.|
|[RowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchy#name)|Nome da RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchy#position)|Posição da RowColumnPivotHierarchy.|
||[fields](/javascript/api/excel/excel.rowcolumnpivothierarchy#fields)|Retorna PivotFields associados a RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchy#id)|ID do RowColumnPivotHierarchy.|
||[Set (Propriedades: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchy#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. RowColumnPivotHierarchyUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.rowcolumnpivothierarchy#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy#settodefault--)|Redefine o RowColumnPivotHierarchy para os valores padrão.|
|[RowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|[Adicionar (pivotHierarchy: Excel. PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#add-pivothierarchy-)|Adiciona o PivotHierarchy ao eixo atual. Se houver a hierarquia em outro lugar na linha, coluna,|
||[getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getcount--)|Obtém o número de hierarquias dinâmicas na coleção.|
||[getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitem-name-)|Obtém RowColumnPivotHierarchy por nome ou id.|
||[getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#getitemornullobject-name-)|Obtém um RowColumnPivotHierarchy por nome. Se o RowColumnPivotHierarchy não existir, retornará um objeto nulo.|
||[items](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#items)|Obtém os itens filhos carregados nesta coleção.|
||[remover (rowColumnPivotHierarchy: Excel. RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection#remove-rowcolumnpivothierarchy-)|Remove o PivotHierarchy do eixo atual.|
|[RowColumnPivotHierarchyCollectionLoadOptions](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions)|[$all](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#$all)||
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#id)|Para cada ITEM da coleção: ID do RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#name)|Para cada ITEM na coleção: nome da RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchycollectionloadoptions#position)|Para cada ITEM da coleção: posição do RowColumnPivotHierarchy.|
|[RowColumnPivotHierarchyData](/javascript/api/excel/excel.rowcolumnpivothierarchydata)|[fields](/javascript/api/excel/excel.rowcolumnpivothierarchydata#fields)|Retorna PivotFields associados a RowColumnPivotHierarchy.|
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchydata#id)|ID do RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchydata#name)|Nome da RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchydata#position)|Posição da RowColumnPivotHierarchy.|
|[RowColumnPivotHierarchyLoadOptions](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions)|[$all](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#$all)||
||[id](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#id)|ID do RowColumnPivotHierarchy.|
||[name](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#name)|Nome da RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchyloadoptions#position)|Posição da RowColumnPivotHierarchy.|
|[RowColumnPivotHierarchyUpdateData](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata)|[name](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata#name)|Nome da RowColumnPivotHierarchy.|
||[position](/javascript/api/excel/excel.rowcolumnpivothierarchyupdatedata#position)|Posição da RowColumnPivotHierarchy.|
|[Tempo](/javascript/api/excel/excel.runtime)|[enableEvents](/javascript/api/excel/excel.runtime#enableevents)|Alternar eventos JavaScript no painel de tarefas ou no suplemento de conteúdo atual.|
|[RuntimeData](/javascript/api/excel/excel.runtimedata)|[enableEvents](/javascript/api/excel/excel.runtimedata#enableevents)|Alternar eventos JavaScript no painel de tarefas ou no suplemento de conteúdo atual.|
|[RuntimeLoadOptions](/javascript/api/excel/excel.runtimeloadoptions)|[enableEvents](/javascript/api/excel/excel.runtimeloadoptions#enableevents)|Alternar eventos JavaScript no painel de tarefas ou no suplemento de conteúdo atual.|
|[RuntimeUpdateData](/javascript/api/excel/excel.runtimeupdatedata)|[enableEvents](/javascript/api/excel/excel.runtimeupdatedata#enableevents)|Alternar eventos JavaScript no painel de tarefas ou no suplemento de conteúdo atual.|
|[ShowAsRule](/javascript/api/excel/excel.showasrule)|[baseField](/javascript/api/excel/excel.showasrule#basefield)|O PivotField base para basear o cálculo ShowAs, se aplicável com base no tipo ShowAsCalculation, caso contrário, null.|
||[baseItem](/javascript/api/excel/excel.showasrule#baseitem)|O Item base para basear o cálculo ShowAs, se aplicável com base no tipo ShowAsCalculation, caso contrário, null.|
||[cálculo](/javascript/api/excel/excel.showasrule#calculation)|O cálculo de ShowAs a ser usado para o Data PivotField. Consulte Excel. ShowAsCalculation para obter detalhes.|
|[Estilo](/javascript/api/excel/excel.style)|[autoIndent](/javascript/api/excel/excel.style#autoindent)|Indica se o texto é automaticamente indentado quando o alinhamento de texto em uma célula é definido como distribuição igual.|
||[textOrientation](/javascript/api/excel/excel.style#textorientation)|A orientação de texto para o estilo.|
|[StyleCollectionLoadOptions](/javascript/api/excel/excel.stylecollectionloadoptions)|[autoIndent](/javascript/api/excel/excel.stylecollectionloadoptions#autoindent)|Para cada ITEM na coleção: indica se o texto será recuado automaticamente quando o alinhamento do texto em uma célula for definido como distribuição igual.|
||[textOrientation](/javascript/api/excel/excel.stylecollectionloadoptions#textorientation)|Para cada ITEM na coleção: a orientação do texto para o estilo.|
|[StyleData](/javascript/api/excel/excel.styledata)|[autoIndent](/javascript/api/excel/excel.styledata#autoindent)|Indica se o texto é automaticamente indentado quando o alinhamento de texto em uma célula é definido como distribuição igual.|
||[textOrientation](/javascript/api/excel/excel.styledata#textorientation)|A orientação de texto para o estilo.|
|[StyleLoadOptions](/javascript/api/excel/excel.styleloadoptions)|[autoIndent](/javascript/api/excel/excel.styleloadoptions#autoindent)|Indica se o texto é automaticamente indentado quando o alinhamento de texto em uma célula é definido como distribuição igual.|
||[textOrientation](/javascript/api/excel/excel.styleloadoptions#textorientation)|A orientação de texto para o estilo.|
|[StyleUpdateData](/javascript/api/excel/excel.styleupdatedata)|[autoIndent](/javascript/api/excel/excel.styleupdatedata#autoindent)|Indica se o texto é automaticamente indentado quando o alinhamento de texto em uma célula é definido como distribuição igual.|
||[textOrientation](/javascript/api/excel/excel.styleupdatedata#textorientation)|A orientação de texto para o estilo.|
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
|[TableCollectionLoadOptions](/javascript/api/excel/excel.tablecollectionloadoptions)|[legacyId](/javascript/api/excel/excel.tablecollectionloadoptions#legacyid)|Para cada ITEM na coleção: retorna uma ID numérica.|
|[TableData](/javascript/api/excel/excel.tabledata)|[legacyId](/javascript/api/excel/excel.tabledata#legacyid)|Retorna uma ID numérica.|
|[TableLoadOptions](/javascript/api/excel/excel.tableloadoptions)|[legacyId](/javascript/api/excel/excel.tableloadoptions#legacyid)|Retorna uma ID numérica.|
|[Workbook](/javascript/api/excel/excel.workbook)|[readOnly](/javascript/api/excel/excel.workbook#readonly)|True se a pasta de trabalho estiver aberta no modo somente leitura. Somente leitura.|
|[WorkbookCreated](/javascript/api/excel/excel.workbookcreated)||[WorkbookData](/javascript/api/excel/excel.workbookdata)|[readOnly](/javascript/api/excel/excel.workbookdata#readonly)|True se a pasta de trabalho estiver aberta no modo somente leitura. Somente leitura.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[readOnly](/javascript/api/excel/excel.workbookloadoptions#readonly)|True se a pasta de trabalho estiver aberta no modo somente leitura. Somente leitura.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[oncalculado](/javascript/api/excel/excel.worksheet#oncalculated)|Ocorre quando a planilha é calculada.|
||[Linhas de grade](/javascript/api/excel/excel.worksheet#showgridlines)|Obtém ou define um sinalizador de linhas de grade da planilha.|
||[meus títulos](/javascript/api/excel/excel.worksheet#showheadings)|É ou define um sinalizador de cabeçalhos da planilha.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[tipo](/javascript/api/excel/excel.worksheetcalculatedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetcalculatedeventargs#worksheetid)|Obtém o id da planilha que é calculada.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[getRange(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrange-ctx-)|Obtém o intervalo que representa a área alterada de uma planilha específica.|
||[getRangeOrNullObject(ctx: Excel.RequestContext)](/javascript/api/excel/excel.worksheetchangedeventargs#getrangeornullobject-ctx-)|Obtém o intervalo que representa a área alterada de uma planilha específica. Pode retornar o objeto null.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[oncalculado](/javascript/api/excel/excel.worksheetcollection#oncalculated)|Ocorre quando qualquer planilha na pasta de trabalho é calculada.|
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[Linhas de grade](/javascript/api/excel/excel.worksheetcollectionloadoptions#showgridlines)|Para cada ITEM na coleção: Obtém ou define o sinalizador de linhas de grade da planilha.|
||[meus títulos](/javascript/api/excel/excel.worksheetcollectionloadoptions#showheadings)|Para cada ITEM na coleção: Obtém ou define o sinalizador de títulos da planilha.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[Linhas de grade](/javascript/api/excel/excel.worksheetdata#showgridlines)|Obtém ou define um sinalizador de linhas de grade da planilha.|
||[meus títulos](/javascript/api/excel/excel.worksheetdata#showheadings)|É ou define um sinalizador de cabeçalhos da planilha.|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[Linhas de grade](/javascript/api/excel/excel.worksheetloadoptions#showgridlines)|Obtém ou define um sinalizador de linhas de grade da planilha.|
||[meus títulos](/javascript/api/excel/excel.worksheetloadoptions#showheadings)|É ou define um sinalizador de cabeçalhos da planilha.|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[Linhas de grade](/javascript/api/excel/excel.worksheetupdatedata#showgridlines)|Obtém ou define um sinalizador de linhas de grade da planilha.|
||[meus títulos](/javascript/api/excel/excel.worksheetupdatedata#showheadings)|É ou define um sinalizador de cabeçalhos da planilha.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do Excel](/javascript/api/excel)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
