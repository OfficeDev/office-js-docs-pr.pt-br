---
title: Conjuntos de requisitos de API JavaScript do Excel
description: ''
ms.date: 10/09/2018
localization_priority: Priority
ms.openlocfilehash: fdcbee0374851f0f88130ae8afe28eec3a0fe77c
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388721"
---
# <a name="excel-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Excel

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Excel são executados em várias versões do Office, incluindo Office 2016 ou posterior para Windows, Office para iPad, Office para Mac e Office Online. A tabela a seguir lista conjuntos de requisitos do Excel, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e as versões ou números de build desses aplicativos.

> [!NOTE]
> Qualquer API que esteja marcada como **Beta** não está pronta para produção do usuário final. Nós as disponibilizamos para que os desenvolvedores testarem em ambientes de teste e desenvolvimento. Porém, não devem ser usadas em documentos de produção/críticos para os negócios.
> 
> Para os conjuntos de requisitos que são marcados como **Beta**usar a versão especificada (ou posterior) do software do Office e usar a biblioteca Beta na CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Entradas não marcadas como **Beta** estão disponíveis e você pode usar biblioteca produção na CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.

|  Conjunto de requisitos  |  Office 365 para Windows\*  |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Beta  | Por favor [Visite nossa página de especificação para abrir API JavaScript do Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec) |
| ExcelApi1.8  | Versão 1808 (Build 10730.20102) ou posterior | 2.17 ou posterior | 16.17 ou posterior | Setembro de 2018 | Em breve |
| ExcelApi1.7  | Versão 1801 (Build 9001.2171) ou posterior   | 2.9 ou posterior | 16.9 ou posterior | Abril de 2018 | Em breve |
| ExcelApi1.6  | Versão 1704 (Compilação 8201.2001) ou posterior   | 2.2 ou posterior |15.36 ou posterior| Abril de 2017 | Em breve|
| ExcelApi1.5  | Versão 1703 (Compilação 8067.2070) ou posterior   | 2.2 ou posterior |15.36 ou posterior| Março de 2017 | Em breve|
| ExcelApi1.4  | Versão 1701 (build 7870.2024) ou posterior   | 2.2 ou posterior |15.36 ou posterior| Janeiro de 2017 | Em breve|
| ExcelApi1.3  | Versão 1608 (build 7369.2055) ou posterior | 1.27 ou posterior |  15.27 ou posterior| Setembro de 2016 | Versão 1608 (build 7601.6800) ou posterior|
| ExcelApi1.2  | Versão 1601 (build 6741.2088) ou posterior | 1.21 ou posterior | 15.22 ou posterior| janeiro de 2016 ||
| ExcelApi1.1  | Versão 1509 (build 4266.1001) ou posterior | 1.19 ou posterior | 15.20 ou posterior| janeiro de 2016 ||

> [!NOTE]
> O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém o conjunto de requisitos 1.1 de ExcelApi.

Para saber mais sobre as versões, números de build e sobre o Servidor do Office Online, confira:

- [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- 
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="whats-new-in-excel-javascript-api-18"></a>Quais são as novidades na API JavaScript do Excel 1.8

O conjunto de requisitos 1.8 da API JavaScript do Excel inclui APIs para tabelas dinâmicas, validação de dados, gráficos, eventos de gráficos, opções de desempenho e criação de pasta de trabalho.

### <a name="pivottable"></a>Tabela Dinâmica

Onda 2 das APIs de Tabela Dinâmica permite que os suplementos definam as hierarquias de uma Tabela Dinâmica. Agora você pode controlar os dados e como eles são agregados. Nosso [Artigo de Tabela Dinâmica](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables) tem mais informações sobre a nova funcionalidade de tabela dinâmica.

### <a name="data-validation"></a>Validação de Dados

A validação de dados permite controlar o que um usuário digita em uma planilha. Você pode limitar as células a conjuntos de respostas predefinidos ou fornecer avisos pop-up sobre entradas indesejadas. Saiba mais sobre [adicionar a validação de dados para intervalos](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-data-validation) hoje.

### <a name="charts"></a>Gráficos

Outra rodada de APIs de gráficos traz um controle programático ainda maior sobre os elementos do gráfico. Agora você tem maior acesso à legenda, eixos, linha de tendência e área de plotagem.

### <a name="events"></a>Eventos

Mais [eventos](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events) foram adicionados para os gráficos. Faça o seu suplemento reagir aos usuários interagindo com o gráfico. Você também pode [alternar eventos](https://docs.microsoft.com/office/dev/add-ins/excel/performance#enable-and-disable-events) disparados em toda a pasta de trabalho.


|Objeto| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Método_ > [createWorkbook(base64File: string)](/javascript/api/excel/excel.application)|Cria uma nova pasta de trabalho oculta usando um arquivo .xlsx com codificação base64 opcional.|1,8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Propriedade_ > formula1|Obtém ou define a Formula1, por exemplo, o valor mínimo ou valor, dependendo do operador.|1,8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Propriedade_ > formula2|Obtém ou define a Formula2, por exemplo, o valor máximo ou valor, dependendo do operador.|1,8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Relação_ > operator|O operador a ser usado para validar os dados.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > categoryLabelLevel|Retorna ou define uma constante de enumeração ChartCategoryLabelLevel referindo-se ao nível de onde os rótulos de categoria estão sendo originados. Leitura/gravação.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > plotVisibleOnly|Verdadeiro se apenas as células visíveis forem plotadas. Falso se ambas as células visíveis e ocultas forem plotadas.. ReadWrite.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > seriesNameLevel|Retorna ou define uma constante de enumeração ChartSeriesNameLevel referente ao nível de origem dos nomes das séries. Leitura/gravação.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > showDataLabelsOverMaximum|Representa se os rótulos de dados devem ser mostrados quando o valor for maior que o valor máximo no eixo de valor.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > style|Retorna ou define o estilo do gráfico para o gráfico. ReadWrite.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Relação_ > displayBlanksAs|Retorna ou define a maneira como as células em branco são plotadas em um gráfico. ReadWrite.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Relação_ > plotArea|Representa a plotArea para o gráfico. Somente leitura.|1,8|
|[chart](/javascript/api/excel/excel.chart)|_Relação_ > plotBy|Retorna ou define como as colunas ou linhas são usadas como séries de dados no gráfico. ReadWrite.|1,8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Propriedade_ > chartId|Obtém o id do gráfico que está ativado.|1,8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Propriedade_ > type|Obtém o tipo do evento.|1,8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual o gráfico é ativado.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Propriedade_ > chartId|Obtém o id do gráfico que é adicionado à planilha.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Propriedade_ > type|Obtém o tipo do evento.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual o gráfico é adicionado.|1,8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Relação_ > source|Obtém a origem do evento.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > isBetweenCategories|Representa se o eixo de valor cruza o eixo de categoria entre as categorias.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > multiLevel|Representa se um eixo é multinível ou não.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > numberFormat|Representa o código de formato para o rótulo de marcação do eixo.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > offset|Representa a distância entre os níveis de rótulos e a distância entre o primeiro nível e a linha do eixo. O valor deve ser um inteiro de 0 a 1000.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > positionAt|Representa a posição do eixo especificada onde o outro eixo cruza. Você deve usar o método SetPositionAt (double) para definir essa propriedade. Somente leitura.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > textOrientation|Representa a orientação do texto do rótulo de seleção do eixo. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > alignment|Representa o alinhamento para o rótulo de escala do eixo especificado.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > position|Representa a posição do eixo especificada onde o outro eixo cruza.|1,8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > [setPositionAt(value: double)](/javascript/api/excel/excel.chartaxis)|Define a posição do eixo especificada onde o outro eixo cruza.|1,8|
|[chartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|_Relação_ > fill|Representa a formatação de preenchimento de gráfico. Somente leitura.|1,8|
|[chartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|_Método_ > [setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle)|Um valor de cadeia de caracteres que representa a fórmula do título do eixo do gráfico usando a notação no estilo A1.|1,8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Relação_ > border|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|1,8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Relação_ > fill|Representa a formatação de preenchimento de gráfico. Somente leitura.|1,8|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Método_ > [clear()](/javascript/api/excel/excel.chartborder)|Limpa a formatação da borda de um elemento do gráfico.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > AutoText|Valor booliano que representa se o rótulo de dados gerará automaticamente o texto apropriado com base no contexto..|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > formula|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de dados usando a notação no estilo A1.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > height|Retorna a altura, em pontos, do rótulo de dados do gráfico. Somente leitura. Nulo se o rótulo de dados do gráfico não estiver visível. Somente leitura.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > left|Representa a distância, em pontos, da borda esquerda do rótulo de dados do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > numberFormat|Valor de cadeia de caracteres que representa o código do formato do rótulo de dados.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > text|Cadeia de caracteres que representa o texto do rótulo de dados em um gráfico.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > textOrientation|Representa a orientação de texto de rótulo de dados do gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > top|Representa a distância, em pontos, da borda superior do rótulo de dados do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > width|Retorna a largura, em pontos, do rótulo de dados do gráfico. Somente leitura. Nulo se o rótulo de dados do gráfico não estiver visível. Somente leitura.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relação_ > format|Representa o formato do rótulo de dados do gráfico. Somente leitura.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relação_ > horizontalAlignment|Representa o alinhamento horizontal de rótulo de dados do gráfico.|1,8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relação_ > verticalAlignment|Representa o alinhamento vertical do rótulo de dados do gráfico.|1,8|
|[chartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|_Relação_ > border|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Propriedade_ > AutoText|Indica se os rótulos de dados geram automaticamente texto apropriado com base no contexto.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Propriedade_ > numberFormat|Representa o código de formatação para rótulos de dados.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Propriedade_ > textOrientation|Representa a orientação de texto dos rótulos de dados. O valor deve ser um número inteiro de -90 a 90 ou 0 o 180 para texto orientado verticalmente.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Relação_ > horizontalAlignment|Representa o alinhamento horizontal de rótulo de dados do gráfico.|1,8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Relação_ > verticalAlignment|Representa o alinhamento vertical do rótulo de dados do gráfico.|1,8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Propriedade_ > chartId|Obtém o id do gráfico que está desativado.|1,8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Propriedade_ > type|Obtém o tipo do evento.|1,8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha em que o gráfico está desativado.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Propriedade_ > chartId|Obtém o id do gráfico que é excluído da planilha.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Propriedade_ > type|Obtém o tipo do evento.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual o gráfico foi deletado.|1,8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Relação_ > source|Obtém a origem do evento.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > height|Representa a altura de legendEntry na legenda do gráfico. Somente leitura.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > index|Representa o índice de legendEntry na legenda do gráfico. Somente leitura.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > left|Representa a esquerda de um gráfico legendEntry. Somente leitura.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > top|Representa a parte superior de um gráfico legendEntry. Somente leitura.|1,8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > width|Representa a largura de legendEntry na legenda do gráfico. Somente leitura.|1,8|
|[chartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|_Relação_ > border|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > height|Representa o valor de altura de plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > insideHeight|Representa o valor insideHeight plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > insideLeft|Representa o valor insideLeft de plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > insideTop|Representa o valor insideTop de plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > insideWidth|Representa o valor insideWidth de plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > left|Representa o valor de plotArea à esquerda.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > top|Representa o valor máximo de plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Propriedade_ > width|Representa o valor de largura de plotArea.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Relação_ > format|Representa a formatação de um gráfico plotArea. Somente leitura.|1,8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Relação_ > position|Represente a posição de plotArea.|1,8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Relação_ > border|Representa os atributos de borda de um gráfico plotArea. Somente leitura.|1,8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Relação_ > fill|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação da tela de fundo. Somente leitura.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > explosion|Retorna ou define o valor de explosão para um gráfico de pizza ou fatia de gráfico de rosca. Retorna 0 (zero) se não houver explosão (a ponta da fatia está no centro da pizza). ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > firstSliceAngle|Retorna ou define o ângulo do primeiro gráfico de pizza ou fatia de gráfico de rosca, em graus (no sentido horário a partir da vertical). Aplica-se apenas a pizza, torta 3-D e gráficos de rosca.. Pode ser um valor de 0 a 360. ReadWrite|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > invertIfNegative|Verdadeiro se o Microsoft Excel inverte o padrão no item quando ele corresponde a um número negativo. ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > overlap|Especifica como barras e colunas são posicionadas. Pode ser um valor entre -100 e 100. Se aplicam apenas às barras 2D e gráficos de colunas 2D. ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > secondPlotSize|Retorna ou define o tamanho da seção secundária de uma pizza do gráfico de pizza ou de uma barra de gráfico de pizza, como uma porcentagem do tamanho da pizza primária. Pode ser um valor de 5 de 200. ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > varyByCategories|Verdadeiro se o Microsoft Excel atribuir uma cor ou padrão diferente a cada marcador de dados. O gráfico deve conter apenas uma série. ReadWrite.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relação_ > axisGroup|Retorna ou define o grupo para a série especificada.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relação_ > dataLabels|Representa uma coleção de todos os dataLabels da série. Somente leitura.|1,8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relação_ > splitType|Retorna ou define a maneira como as duas seções de uma pizza do gráfico de pizza ou de uma barra do gráfico de pizza são divididas. ReadWrite.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > backwardPeriod|Representa o número de períodos que a linha de tendência se estende para trás.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > forwardPeriod|Representa o número de períodos que a linha de tendência se estende para frente.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > showEquation|Verdadeiro se a equação da linha de tendência for exibida no gráfico.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > showRSquared|Verdadeiro se o R-quadrado da linha de tendência for exibido no gráfico.|1,8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Relação_ > label|Representa o rótulo de linha de tendência um gráfico. Somente leitura.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > AutoText|Valor booliano que representa se o rótulo de linha de tendência gerará automaticamente o texto apropriado com base no contexto..|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > formula|Valor de cadeia de caracteres que representa a fórmula do título do rótulo de linha de tendência usando a notação no estilo A1.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > height|Retorna a altura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível. Somente leitura.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > left|Representa a distância, em pontos, da borda esquerda do rótulo da linha de tendência do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > numberFormat|Valor de cadeia de caracteres que representa o código do formato do rótulo de linha de tendência.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > text|Cadeia de caracteres que representa o texto do rótulo em um gráfico de linha de tendência.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > textOrientation|Representa a orientação de texto de rótulo de linha de tendência de gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > top|Representa a distância, em pontos, da borda superior do rótulo de linha de tendência do gráfico até a borda superior da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Propriedade_ > width|Retorna a largura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível. Somente leitura.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relação_ > format|Representa o formato do rótulo de linha de tendência de gráfico. Somente leitura.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relação_ > horizontalAlignment|Representa o alinhamento horizontal de rótulo de linha de tendência de gráfico.|1,8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relação_ > verticalAlignment|Representa o alinhamento vertical do rótulo de linha de tendência de gráfico.|1,8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relação_ > border|Representa o formato de borda, que inclui a espessura de cor e estilo de linha. Somente leitura.|1,8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relação_ > fill|Representa o formato de preenchimento do rótulo de linha de tendência atual do gráfico. Somente leitura.|1,8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relação_ > font|Representa os atributos de fonte do rótulo de linha de tendência do gráfico, como nome, tamanho, cor, dentre outros. Somente leitura.|1,8|
|[customDataValidation](/javascript/api/excel/excel.customdatavalidation)|_Propriedade_ > formula| Uma fórmula de validação de dados personalizados. Isso cria regras especiais de entrada, como impedir duplicatas ou limitar o total em um intervalo de células.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Propriedade_ > id|ID do DataPivotHierarchy. Somente leitura.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Propriedade_ > nome|Nome da DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Propriedade_ > numberFormat|Formato de número do DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Propriedade_ > posição|Posição da DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relação_ > field|Retorna PivotFields associados a DataPivotHierarchy. Somente leitura.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relação_ > showAs|Determina se os dados devem ser mostrados como um cálculo de resumo específico ou não.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relação_ > summarizeBy|Determina se deve mostrar todos os itens a DataPivotHierarchy.|1,8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Método_ > [setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault--)|Redefina a DataPivotHierarchy para os valores padrão.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Propriedade_ > itens|Um conjunto de objetos dataPivotHierarchy. Somente leitura.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Método_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Adiciona o PivotHierarchy ao eixo atual.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Método_ > [getCount()](/javascript/api/excel/excel.datapivothierarchycollection)|Obtém o número de hierarquias dinâmicas na coleção.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Método_ > [getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Obtém DataPivotHierarchy por nome ou id.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Método_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Obtém uma DataPivotHierarchy por nome. Se o DataPivotHierarchy não existir, retornará um objeto nulo.|1,8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Método_ > [remove(DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Remove o PivotHierarchy do eixo atual.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Propriedade_ > ignoreBlanks|Ignora espaços em branco: nenhuma validação de dados será executada em células vazias, o padrão será definido como verdadeiro.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Propriedade_ > valid|Representa se todos os valores de célula são válidos de acordo com as regras de validação de dados. Somente leitura.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relação_ > errorAlert|Alerta de erro quando o usuário insere dados inválidos.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relação_ > prompt|Avisa quando os usuários selecionam uma célula.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relação_ > rule|Regra de validação de dados que contém diferentes tipos de critérios de validação de dados.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relação_ > type|Tipo de validação de dados, confira [Excel.DataValidationType](/javascript/api/excel/excel.datavalidationtype) para obter detalhes. Somente leitura.|1,8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Método_ > [clear()](/javascript/api/excel/excel.datavalidation)|Desfazer a validação de dados do intervalo atual.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Propriedade_ > mensagem|Representa a mensagem de alerta de erro.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Propriedade_ > showAlert|Determina se deseja mostrar uma caixa de diálogo de alerta de erro ou não quando um usuário insere dados inválidos. O padrão é verdadeiro.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Propriedade_ > title|Representa o título da caixa de diálogo de alerta de erro.|1,8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Relação_ > style|Representa o tipo de alerta de validação de dados, confira [Excel.DataValidationAlertStyle](/javascript/api/excel/excel.datavalidationalertstyle) para obter detalhes.|1,8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).|_Propriedade_ > mensagem|Representa a mensagem a solicitação.|1,8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).|_Propriedade_ > showPrompt|Determina se deseja ou não mostrar o prompt quando o usuário seleciona uma célula com a validação de dados.|1,8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).|_Propriedade_ > title|Representa o título para a solicitação.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > custom|Critérios de validação de dados personalizados.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > date|Critérios de validação de dados de data.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > decimal|Critérios de validação de dados decimais.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > list|Critérios de validação de dados da lista.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > textLength|Critérios de validação de dados TextLength.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > time|Critérios de validação de dados de tempo.|1,8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule).|_Relação_ > wholeNumber|Critérios de validação de dados WholeNumber.|1,8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Propriedade_ > formula1|Obtém ou define a Formula1, por exemplo, o valor mínimo ou valor, dependendo do operador.|1,8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Propriedade_ > formula2|Obtém ou define a Formula2, por exemplo, o valor máximo ou valor, dependendo do operador.|1,8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Relação_ > operator|O operador a ser usado para validar os dados.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Propriedade_ > enableMultipleFilterItems|Determina se deseja permitir vários itens de filtro.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Propriedade_ > id|ID do FilterPivotHierarchy. Somente leitura.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Propriedade_ > nome|Nome do FilterPivotHierarchy.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Propriedade_ > posição|Posição do FilterPivotHierarchy.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Relação_ > fields|Retorna PivotFields associados a FilterPivotHierarchy. Somente leitura.|1,8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Método_ > [setToDefault()](/javascript/api/excel/excel.filterpivothierarchy)|Redefina a FilterPivotHierarchy para os valores padrão.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Propriedade_ > itens|Um conjunto de objetos filterPivotHierarchy. Somente leitura.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Método_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Adiciona o PivotHierarchy ao eixo atual. Se houver hierarquia em outro lugar na linha, coluna ou eixo de filtro, ele será removido desse local.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Método_ > [getCount()](/javascript/api/excel/excel.filterpivothierarchycollection)|Obtém o número de hierarquias dinâmicas na coleção.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Método_ > [getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Obtém FilterPivotHierarchy por nome ou id.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Método_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Obtém um FilterPivotHierarchy por nome. Se o FilterPivotHierarchy não existir, retornará um objeto nulo.|1,8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Método_ > [remove(filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Remove o PivotHierarchy do eixo atual.|1,8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Propriedade_ > inCellDropDown|Exibido na lista na célula suspensa ou não, ele será padronizado como verdadeiro.|1,8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Propriedade_ > source|Fonte da lista de validação de dados|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Propriedade_ > id|ID do PivotField.. Somente leitura.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Propriedade_ > nome|Nome do PivotField.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Propriedade_ > showAllItems|Determina se deseja mostrar todos os itens de PivotField.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Relação_ > items|Retorna PivotFields associados ao PivotField. Somente leitura.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Relação_ > subtotals|Subtotais de PivotField.|1,8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Método_ > [sortByLabels(sortby: SortBy)](/javascript/api/excel/excel.pivotfield)|Classifica o PivotField. Se um DataPivotHierarchy for especificado, a classificação será aplicada com base nele, se a classificação não for baseada no campo PivotField.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Propriedade_ > itens|Um conjunto de objetos pivotField. Somente leitura.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Método_ > [getCount()](/javascript/api/excel/excel.pivotfieldcollection)|Obtém o número de hierarquias dinâmicas na coleção.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Método_ > [getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Obtém PivotHierarchy por nome ou id.|1,8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Método_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Obtém o PivotHierarchy por nome. Se o PivotHierarchy não existir, retornará um objeto null.|1,8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Propriedade_ > id|ID do PivotHierarchy. Somente leitura.|1,8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Propriedade_ > nome|Nome do PivotHierarchy.|1,8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Relação_ > fields|Retorna PivotFields associados a PivotHierarchy. Somente leitura.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Propriedade_ > itens|Um conjunto de objetos pivotHierarchy. Somente leitura.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Método_ > [getCount()](/javascript/api/excel/excel.pivothierarchycollection)|Obtém o número de hierarquias dinâmicas na coleção.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Método_ > [getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Obtém PivotHierarchy por nome ou id.|1,8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Método_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Obtém o PivotHierarchy por nome. Se o PivotHierarchy não existir, retornará um objeto null.|1,8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Propriedade_ > id|ID do PivotItem. Somente leitura.|1,8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Propriedade_ > isExpanded|Determina se o item está expandido para mostrar itens filho ou se ele está recolhido e os itens filho estão ocultos.|1,8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Propriedade_ > nome|Nome do PivotItem.|1,8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Propriedade_ > visible|Determina se o PivotItem ficará visível ou não.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Propriedade_ > itens|Um conjunto de objetos pivotItem. Somente leitura.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Método_ > [getCount()](/javascript/api/excel/excel.pivotitemcollection)|Obtém o número de hierarquias dinâmicas na coleção.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Método_ > [getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Obtém PivotHierarchy por nome ou id.|1,8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Método_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Obtém o PivotHierarchy por nome. Se o PivotHierarchy não existir, retornará um objeto null.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Propriedade_ > showColumnGrandTotals|Verdadeiro, quando o relatório de Tabela Dinâmica mostra os totais de colunas.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Propriedade_ > showRowGrandTotals|Verdadeiro, quando o relatório de Tabela Dinâmica mostra os totais de linhas.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Propriedade_ > subtotalLocation|Essa propriedade indica SubtotalLocationType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo. Valores possíveis são: AtTop, AtBottom.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Relação_ > layoutType|Essa propriedade indica o PivotLayoutType de todos os campos da Tabela Dinâmica. Se os campos têm diferentes estados, ele será nulo.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Método_ > [getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo onde residem os rótulos de coluna da Tabela Dinâmica.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Método_ > [getDataBodyRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo onde residem os valores de dados da tabela dinâmica.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Método_ > [getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo de área de filtro da Tabela Dinâmica.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Método_ > [getRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo em que a Tabela Dinâmica existe, excluindo a área de filtro.|1,8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Método_ > [getRowLabelRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo onde residem os rótulos de linha da Tabela Dinâmica.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > columnHierarchies|As hierarquias de pivô da coluna da Tabela Dinâmica. Somente leitura.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > dataHierarchies|As hierarquias dinâmicas de dados da Tabela Dinâmica. Somente leitura.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > filterHierarchies|As hierarquias de pivô do filtro da Tabela Dinâmica. Somente leitura.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > hierarchies|Hierarquias pivô da Tabela Dinâmica. Somente leitura.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > layout|O PivotLayout descreve o layout e estrutura visual da Tabela Dinâmica. Somente leitura.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > rowHierarchies|As hierarquias de pivô de linha da Tabela Dinâmica. Somente leitura.|1,8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Método_ > [delete()](/javascript/api/excel/excel.pivottable)|Exclui a Tabela Dinâmica.|1,8|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > [adicionar (nome: cadeia de caracteres, fonte: objeto, destino: objeto)](/javascript/api/excel/excel.pivottablecollection)|Adiciona um Pivottable com base nos dados de origem especificados e insere-o na célula superior esquerda do intervalo de destino.|1,8|
|[range](/javascript/api/excel/excel.range)|_Relação_ > dataValidation|Retorna um objeto de validação de dados. Somente leitura.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Propriedade_ > id|ID do RowColumnPivotHierarchy. Somente leitura.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Propriedade_ > nome|Nome da RowColumnPivotHierarchy.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Propriedade_ > posição|Posição da RowColumnPivotHierarchy.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Relação_ > fields|Retorna PivotFields associados a RowColumnPivotHierarchy. Somente leitura.|1,8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Método_ > [setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy)|Redefine o RowColumnPivotHierarchy para os valores padrão.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Propriedade_ > itens|Um conjunto de objetos rowColumnPivotHierarchy. Somente leitura.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Método_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Adiciona o PivotHierarchy ao eixo atual. Se houver a hierarquia em outro lugar na linha, coluna,|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Método_ > [getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Obtém o número de hierarquias dinâmicas na coleção.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Método_ > [getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Obtém RowColumnPivotHierarchy por nome ou id.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Método_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Obtém um RowColumnPivotHierarchy por nome. Se o RowColumnPivotHierarchy não existir, retornará um objeto nulo.|1,8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Método_ > [remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Remove o PivotHierarchy do eixo atual.|1,8|
|[runtime](/javascript/api/excel/excel.runtime)|_Propriedade_ > enableEvents|Alterna os eventos JavaScript no painel de tarefas atual ou no suplemento de conteúdo.|1,8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relação_ > baseField|O PivotField base para basear o cálculo ShowAs, se aplicável com base no tipo ShowAsCalculation, caso contrário, null.|1,8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relação_ > baseItem|O Item base para basear o cálculo ShowAs, se aplicável com base no tipo ShowAsCalculation, caso contrário, null.|1,8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relação_ > calculation|O cálculo de ShowAs a ser usado para o Data PivotField.|1,8|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > autoIndent|Indica se o texto é automaticamente indentado quando o alinhamento de texto em uma célula é definido como distribuição igual.|1,8|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > textOrientation|A orientação de texto para o estilo.|1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > automatic|Se Automatic for definido como true, todos os outros valores serão ignorados ao definir os subtotais.|1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > average| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > count| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > countNumbers| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > max| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > min| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > product| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > standardDeviation| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > standardDeviationP| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > sum| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > variance| |1,8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > varianceP| |1,8|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > legacyId|Retorna uma identificação numérica. Somente leitura.|1,8|
|[workbook](/javascript/api/excel/excel.workbook)|_Propriedade_ > readOnly|True se a pasta de trabalho estiver aberta no modo somente leitura. Somente leitura.|1,8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Propriedade_ > id|Retorna um valor que identifica de forma exclusiva o objeto WorkbookCreated. Somente leitura.|1,8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Método_ > [Open()](/javascript/api/excel/excel.workbookcreated)|Abra a pasta de trabalho.|1,8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > showGridlines|Obtém ou define um sinalizador de linhas de grade da planilha.|1,8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > showHeadings|É ou define um sinalizador de cabeçalhos da planilha.|1,8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Propriedade_ > type|Obtém o tipo do evento.|1,8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha que é calculada.|1,8|

## <a name="whats-new-in-excel-javascript-api-17"></a>Quais são as novidades na API JavaScript do Excel 1.7

O conjunto de requisitos 1.7 da API JavaScript do Excel incluei APIs para gráficos, eventos, planilhas, intervalos, propriedades do documento, itens nomeados, opções de proteção e estilos.

### <a name="customize-charts"></a>Personalize gráficos

Com as novas APIs de gráficos, você pode criar tipos degráficos adicionais, adicionar uma série de dados a um gráfico, definir o título do gráfico, adicionar um título de eixo, adicionar unidade de exibição, adicionar uma linha de tendência com média móvel, alterar uma linha de tendência para linear e muito mais. Estes são alguns exemplos:

* Eixo gráfico - obtenha, defina, formate e remova unidade de eixo, etiqueta e título em um gráfico.
* Série de gráficos - adicione, defina e exclua uma série em um gráfico.  Alterar marcadores da série, pedidos de plotagem e dimensionamento.
* Gráfico de linhas de tendências: adicione, receba e formate linhas de tendências em um gráfico.
* Legenda do gráfico - formate a fonte de legenda de um gráfico.
* Ponto do gráfico - defina a cor do ponto do gráfico.
* Subtítulo do título do gráfico - obtenha e defina a subseqüência do título para um gráfico.
* Tipo de gráfico - opção para criar mais tipos de gráfico.

### <a name="events"></a>Eventos

As APIs de eventos JavaScript do Excel fornecem diversos,  manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Você pode criar essa função para executar as ações que seu cenário exige. Para obter uma lista de eventos que estão disponíveis, confira [trabalhar com eventos usando as API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events).

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Personalizar a aparência de planilhas e intervalos

Nas novas APIs você pode personalizar a aparência das planilhas de várias maneiras:

* Congele painéis para manter linhas ou colunas específicas visíveis durante a rolagem na planilha. Por exemplo, se a primeira linha da planilha inclui cabeçalhos, você pode congelá-la para que os cabeçalhos das colunas permaneçam visíveis enquanto rola para baixo na planilha.
* Modificar a cor da guia de planilha.
* Adicione títulos de planilha.


Você pode personalizar a aparência de intervalos de várias maneiras:

* Defina o estilo de célula para um intervalo para garantir que todas as células no intervalo tenham formatação consistente. Um estilo de célula é um conjunto definido de características de formatação, como fontes e tamanhos de fonte, formatos numéricos, bordas de célula e sombreamento de célula. Use qualquer um dos estilos de célula internas do Excel ou crie seu próprio estilo de célula personalizado.
* Defina a orientação de texto para um intervalo.
* Adicione ou modifique um hiperlink em um intervalo vinculado a outro local na pasta de trabalho ou a um local externo.

### <a name="manage-document-properties"></a>Gerenciar propriedades dos documentos

Usando as APIs de propriedades do documento, você pode acessar as propriedades do documento interno e também criar e gerenciar propriedades personalizadas do documento para armazenar o estado da pasta de trabalho e direcionar o fluxo de trabalho e a lógica comercial.

### <a name="copy-worksheets"></a>Copiar planilhas

Usando a cópia da planilha APIs, você pode copiar os dados e o formato de uma planilha para uma nova planilha na mesma pasta de trabalho e reduzir a quantidade de transferência de dados necessária.

### <a name="handle-ranges-with-ease"></a>Lidar com intervalos com facilidade

Usando várias APIs de intervalo, você pode fazer coisas como obter região ao redor, obter um intervalo redimensionado e muito mais. Essas APIs devem tornar as tarefas, como manipulação de intervalo e endereçamento, muito mais eficientes.

Além disso:

* Opções de proteção de pasta de trabalho e planilha - use estas APIs para proteger dados em uma planilha e a estrutura da pasta de trabalho.
* Atualizar um item nomeado - usar esta API para atualizar um item nomeado.
* Obter a célula ativa - usar esta API para acessar a célula ativa da pasta de trabalho.

|Objeto| Quais são as novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > chartType|Representa o tipo de gráfico. Valores ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > id|Id exclusiva do gráfico. Somente leitura.|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > showAllFieldButtons|Representa se deseja exibir todos os botões de campo em um Gráfico Dinâmico.|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_Relação_ > border|Representa o formato da borda da área de gráfico, incluindo a cor, estilo de linha e espessura. Somente leitura.|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_Método_ > getItem (tipo: cadeia de caracteres, grupo: cadeia de caracteres)|Retorna o eixo específico identificado por tipo e grupo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > axisBetweenCategories|Representa se o eixo de valor cruza o eixo de categoria entre as categorias.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > axisGroup|Representa o grupo para o eixo especificado. Somente leitura. Os valores possíveis são: Primary, Secondary.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > categoryType|Retorna ou define o tipo de eixo de categoria. Os valores possíveis são: TextAxis, DateAxis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > crosses.|Representa eixo especificado onde o outro eixo cruza. Os valores possíveis são: Automatic, Maximum, Minimum, Custom.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > crossesAt|Representa eixo especificado onde o outro eixo cruza. Somente leitura. A definição para essa propriedade deve usar o método SetCrossesAt (duplo). Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > customDisplayUnit|Representa o valor da unidade de exibição do eixo personalizado. Somente leitura. Para definir essa propriedade, use o método de SetCustomDisplayUnit(duplo). Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > displayUnit|Representa a unidade de exibição de eixo. Os valores possíveis são: None, Hundreds, Thousands, TenThousands, HundredThousands, Millions, TenMillions, HundredMillions, Billions, Trillions, Custom.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > height|Representa a altura, em pontos, do eixo do gráfico. Nulo se o eixo não for visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > left|Representa a distância, em pontos, da borda esquerda do eixo à esquerda da área do gráfico. Nulo se o eixo não for visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > logBase|Representa a base do logaritmo ao usar escalas logarítmicas.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > reversePlotOrder|Representa se o Microsoft Excel plota os pontos de dados do último para o primeiro.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > scaleType|Representa o tipo de escala do eixo dos valores. Valores possíveis são: Linear, Logarithmic.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > showDisplayUnitLabel|Indica se a etiqueta de unidade de exibição de eixo está visível.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > tickLabelSpacing|Representa o número série ou categorias entre os rótulos de marcas de escala. Pode ser um valor de 1 a 31999 ou uma cadeia de caracteres vazia para configuração automática. O valor retornado sempre é um número.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > tickMarkSpacing|Representa o número de série ou categorias entre as marcas de escala.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > top|Representa a distância, em pontos, da borda superior do eixo a parte superior da área do gráfico. Nulo se o eixo não for visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > type|Representa o tipo de eixo. Somente leitura. Os valores possíveis são: Invalid, Category, Value, Series.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > visible|Um valor booliano representa a visibilidade do eixo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > width|Representa a largura, em pontos, do eixo do gráfico. Nulo se o eixo não for visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > baseTimeUnit|Retorna ou define a unidade base para o eixo da categoria especificada.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > majorTickMark|Representa o tipo de marca de escala principal para o eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > majorTimeUnitScale|Retorna ou define o valor de escala de unidades principais para o eixo das categorias quando a propriedade CategoryType estiver definida como escala de tempo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > minorTickMark|Representa o tipo de marca de escala secundária para o eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > minorTimeUnitScale|Retorna ou define o valor da escala unitária secundária para o eixo da categoria quando a propriedade CategoryType estiver definida como TimeScale.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relação_ > tickLabelPosition|Representa a posição dos rótulos de marcas de escala no eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > setCategoryNames(sourceData: Range)|Define todos os nomes de categoria para o eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > setCrossesAt(valor: duplo)|Define o eixo especificado onde o outro eixo cruza.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > setCustomDisplayUnit(valor: duplo)|Definirá a unidade de exibição de eixo a um valor personalizado.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Propriedade_ > color|Código de cor HTML que representa a cor das bordas no gráfico.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Propriedade_ > espessura|Representa a espessura da borda, em pontos.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Relação_ > lineStyle|Representa o estilo de linha da borda.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > posição|Valor de DataLabelPosition que representa a posição do rótulo de dados. Os valores possíveis são: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > separator|Cadeia de caracteres que representa o separador usado para o rótulo de dados em um gráfico.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showBubbleSize|Valor booliano que determina se o tamanho da bolha do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showCategoryName|Valor booliano que determina se o nome da categoria do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showLegendKey|Valor booliano que determina se o código de legenda do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showPercentage|Valor booliano que determina se o percentual do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showSeriesName|Valor booliano que determina se o nome da série do rótulo de dados fica visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Propriedade_ > showValue|Valor booliano que determina se o valor do rótulo de dados fica visível ou não.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > height|Representa a altura da legenda no gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > left|Representa a esquerda de uma legenda do gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > showShadow|Representa se a legenda tem sombra no gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > top|Representa o início de uma legenda do gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Propriedade_ > width|Representa a largura da legenda no gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Relação_ > legendEntries|Representa uma coleção de legendEntries na legenda. Somente leitura.|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Propriedade_ > visible|Representa o visível de uma entrada de legenda do gráfico.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Propriedade_ > itens|Um conjunto de objetos chartLegendEntry. Somente leitura.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Método_ > getCount()|Retorna o número de legendEntry da coleção.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Método_ > getItemAt(index: número)|Retorna legendEntry no índice fornecido.|1.7|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > hasDataLabel|Representa se um ponto de dados possui um datalabel. Não aplicável para gráficos de superfície.|1.7|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > markerBackgroundColor|Representação do código de cor HTML da cor de fundo do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|1.7|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > markerForegroundColor|Representação do código de cor HTML da cor de primeiro plano do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|1.7|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > markerSize|Representa o tamanho do marcador do ponto de dados.|1.7|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|_Propriedade_ > markerStyle|Representa estilo do marcador de um ponto de dados do gráfico. Os valores possíveis são: Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|1.7|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|_Relação_ > dataLabel|Retorna o rótulo de dados de um ponto de gráfico. Somente leitura.|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_Relação_ > border|Representa o formato da borda de um ponto de dados do gráfico, que inclui informações de cor, estilo e a espessura. Somente leitura.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > chartType|Representa o tipo de gráfico de uma série. Valores ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > doughnutHoleSize|Representa o tamanho do furo de rosca de uma série de gráficos.  Válida apenas em gráficos de rosca e doughnutExploded.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > filtered|Valor booliano representando se a série é filtrada ou não. Não aplicável para gráficos de superfície.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > gapWidth|Representa a largura do espaçamento de uma série de gráfico.  Válida apenas sobre gráficos de barras e colunas, bem como|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > hasDataLabels|Valor booliano representando se a série tem rótulos de dados ou não.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > markerBackgroundColor|Representa a cor de fundo dos marcadores de uma série de gráficos.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > markerForegroundColor|Representa cor de primeiro plano dos marcadores de uma série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > markerSize|Representa o tamanho do marcador de uma série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > markerStyle|Representa o estilo do marcador de uma série de gráfico. Os valores possíveis são: Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > plotOrder|Representa a ordem de plotagem de uma série de gráficos dentro do grupo de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > showShadow|Valor booliano representando se a série tem sombra ou não.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Propriedade_ > smooth|Valor booliano representando se a série é suave ou não. Apenas para gráficos de linha e de dispersão.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relação_ > dataLabels|Representa uma coleção de todos os dataLabels da série. Somente leitura.|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relação_ > linhas de tendência|Representa uma coleção de todas as linha de tendência da série. Somente leitura.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > Delete()|Exclui a série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setBubbleSizes(sourceData: Range)|Definir tamanhos das bolhas para uma série de gráfico. Funciona apenas para gráficos de bolhas.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setValues(sourceData: Range)|Definir valores de uma série de gráficos. Para gráfico de dispersão, isso significa valores do eixo Y.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setXAxisValues(sourceData: Range)|Definir valores do eixo X para uma série de gráficos. Funciona apenas para gráficos de dispersão.|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Método_ > add (nome: cadeia de caracteres, indexar: número)|Adiciona uma nova série para o conjunto.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > height|Representa a altura, em pontos, do título do gráfico. Somente leitura. Nulo se o título do gráfico não estiver visível. Somente leitura.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > horizontalAlignment|Representa o alinhamento horizontal para título do gráfico. Os valores possíveis são: Center, Left, Justify, Distributed, Right.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > left|Representa a distância, em pontos, da borda esquerda do título do gráfico até a borda esquerda da área do gráfico. Nulo se o título do gráfico não estiver visível.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > posição|Representa a posição de título do gráfico. Os valores possíveis são: Top, Automatic, Bottom, Right, Left.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > showShadow|Representa um valor booliano que determina se o título do gráfico tiver uma sombra.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > textOrientation|Representa a orientação de texto do título do gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > top|Representa a distância em pontos, da borda superior do título do gráfico a parte superior da área do gráfico. Nulo se o título do gráfico não estiver visível.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > verticalAlignment|Representa o alinhamento vertical do título do gráfico. Os valores possíveis são: Center, Bottom, Top, Justify, Distributed.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Propriedade_ > width|Retorna a largura em pontos do título do gráfico. Somente leitura. Nulo se o título do gráfico não estiver visível. Somente leitura.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Método_ > setFormula(fórmula: cadeia de caracteres)|Define um valor de cadeia de caracteres que representa a fórmula do título do eixo do gráfico usando a notação no estilo A1.|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_Relação_ > border|Representa o formato da borda do título do gráfico, incluindo a cor, estilo de linha e espessura. Somente leitura.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > backward|Representa o número de períodos que a linha de tendência se estende para trás.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > displayEquation|Verdadeiro se a equação da linha de tendência for exibida no gráfico.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > displayRSquared|Verdadeiro se o R-quadrado da linha de tendência for exibido no gráfico.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > forward|Representa o número de períodos que a linha de tendência se estende para frente.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > intercept|Representa o valor de intercepção da linha de tendência. Pode ser definido como um valor numérico ou uma cadeia de caracteres vazia (para valores automáticos de eixo). O valor retornado sempre é um número.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > movingAveragePeriod|Representa o período de uma linha de tendência do gráfico, apenas para a linha de tendência com o tipo MovingAverage.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > nome|Representa o nome da linha de tendência. Pode ser definido como um valor de sequência ou pode ser definido como valor nulo para representar valores automáticos. O valor retornado sempre é uma cadeia de caracteres.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > polynomialOrder|Representa a ordem de uma linha de tendência do gráfico, apenas para a linha de tendência com o tipo Polynomial.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Propriedade_ > type|Representa o tipo da linha de tendência de um gráfico. Valores possíveis são: Linear, Exponential, Logarithmic, MovingAverage, Polynomial, Power.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Relação_ > format|Representa a formatação de uma linha de tendência do gráfico. Somente leitura.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Método_ > Delete()|Deleta o objeto Trendline.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Propriedade_ > itens|Um conjunto de objetos chartTrendline. Somente leitura.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Método_ > add(type: string)|Adiciona uma nova linha de tendência ao conjunto de linha de tendência.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Método_ > getCount()|Retorna o número de linha de tendência na coleção.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Método_ > getItem(index: number)|Obtém o objeto da linha de tendência por índice, que é a ordem de inserção na matriz de itens.|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_Relação_ > line|Representa a formatação de linha do gráfico. Somente leitura.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propriedade_ > key|Obtém a chave da propriedade personalizada. Somente leitura. Somente leitura.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propriedade_ > type|Obtém o tipo de valor da propriedade personalizada. Somente leitura. Somente leitura. Os valores possíveis são: Number, Boolean, Date, String, Float.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Propriedade_ > value|Obtém ou define o valor da propriedade personalizada.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Método_ > Delete()|Exclui a propriedade personalizada.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Propriedade_ > itens|Uma coleção de objetos customProperty. Somente leitura.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > add (key: string, value: object)|Cria uma nova propriedade personalizada ou define uma existente.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > deleteAll()|Exclui todas as propriedades personalizadas nesta coleção.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > getCount()|Obtém a contagem das propriedades personalizadas.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > getItem(key: string)|Obtém um objeto de propriedade personalizado por sua chave, que não faz distinção entre maiúsculas e minúsculas. Lança se a propriedade customizada não existir.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Método_ > getItemOrNullObject(key: string)|Obtém um objeto de propriedade personalizado por sua chave, que não faz distinção entre maiúsculas e minúsculas. Retorna um objeto nulo se a propriedade customizada não existir..|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Propriedade_ > itens|Um conjunto de objetos de conexão de dados. Somente leitura.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Método_ > refreshAll()|Atualiza todas as conexões de dados da coleção.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > author|Obtém ou define o autor da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > category|Obtém ou define a categoria da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > comments|Obtém ou define os comentários da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > company|Obtém ou define a empresa do documento.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > keywords|Obtém ou define as palavras-chave da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > lastAuthor|Obtém o último autor da pasta de trabalho. Somente leitura. Somente leitura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > manager|Obtém ou define o gerenciador da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > revisionNumber|Obtém o número de revisão da pasta de trabalho. Somente leitura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > subject|Obtém ou define o assunto da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Propriedade_ > title|Obtém ou define o título da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relação_ > creationDate|Obtém a data de criação da pasta de trabalho. Somente leitura. Somente leitura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relação_ > custom|Obtém a coleção de propriedades personalizadas da pasta de trabalho. Somente leitura. Somente leitura.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propriedade_ > formula|Obtém ou define a fórmula do item nomeado.  A fórmula sempre começa com um sinal de "=".|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relação_ > arrayValues|Retorna um objeto que contém valores e tipos do item nomeado. Somente leitura.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Propriedade_ > types|Representa os tipos de cada item na matriz de itens nomeados como somente leitura. Os valores possíveis são: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Propriedade_ > values|Representa os valores de cada item na matriz de itens nomeados. Somente leitura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > isEntireColumn|Representa se o intervalo atual está em uma coluna inteira. Somente leitura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > isEntireRow|Representa se o intervalo atual está em uma linha inteira. Somente leitura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > numberFormatLocal|Representa o código de formato numérico do Excel para o intervalo fornecido como uma cadeia de caracteres no idioma do usuário.|1.7|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > style|Representa o estilo de intervalo atual. Isso retornará nulo ou uma cadeia de caracteres.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > getAbsoluteResizedRange (numRows: número numColumns: número)|Obtém um objeto Range com a mesma célula superior esquerda do objeto Range atual, mas com os números especificados de linhas e colunas.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > getImage()|O intervalo é renderizado como uma imagem em base 64.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > getSurroundingRegion()|Retorna um objeto Range que representa a região circundante da célula superior esquerda nesse intervalo. Uma região ao redor é um intervalo limitado por qualquer combinação de linhas e colunas em branco em relação a esse intervalo.|1.7|
|[range](/javascript/api/excel/excel.range)|_Método_ > showCard()|Exibe o cartão para uma célula ativa se ele tiver um conteúdo valioso.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > textOrientation|Obtém ou define a orientação de texto de todas as células no intervalo.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > useStandardHeight|Determina se a altura da linha do objeto Range é igual a altura padrão da planilha.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > useStandardWidth|Determina se a largura da coluna do objeto Range é igual a largura padrão da planilha.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriedade_ > address|Representa o destino da url do hiperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriedade_ > document.|Representa o documento. meta do hiperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriedade_ > screenTip|Representa a cadeia exibida ao passar o mouse sobre o hiperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Propriedade_ > textToDisplay|Representa a cadeia de caracteres exibida na parte superior esquerda da maioria das células no intervalo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > addIndent|Indica se o texto é automaticamente indentado quando o alinhamento de texto em uma célula é definido como distribuição igual.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > autoIndent|Indica se o texto é automaticamente indentado quando o alinhamento de texto em uma célula é definido como distribuição igual.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > builtIn|Indica se o estilo é um estilo interno. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > formulaHidden|Indica se a fórmula ficará oculta quando a planilha estiver protegida.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > horizontalAlignment|Representa o alinhamento horizontal para o estilo. Os valores possíveis são: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeAlignment|Indica se o estilo incluem as propriedades AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, e TextOrientation.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeBorder|Indica se o estilo inclui as propriedades de borda Color, ColorIndex, LineStyle e Weight.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeFont|Indica se o estilo inclui as propriedades de fonte Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript e Underline.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeNumber|Indica se o estilo inclui a propriedade NumberFormat.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includePatterns|Indica se o estilo inclui as propriedades internas Color, ColorIndex, InvertIfNegative, Pattern, PatternColor e PatternColorIndex.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > includeProtection|Indica se o estilo incluirá as propriedades de proteção FormulaHidden e Locked.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > indentLevel|Um número inteiro entre 0 e 250 que indica o nível de recuo do estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > locked|Indica se o objeto é bloqueado quando a planilha está protegida.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > nome|O nome do estilo. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > numberFormat|O código de formatação de formato de número para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > numberFormatLocal|O código de formato localizado do formato numérico para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > orientation|A orientação de texto para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > readingOrder|A ordem de leitura para o estilo. Os valores possíveis são: Context, LeftToRight, RightToLeft.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > shrinkToFit|Indica se o texto é automaticamente reduzido para caber na largura da coluna disponível.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > textOrientation|A orientação de texto para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > verticalAlignment|Representa o alinhamento vertical do estilo. Os valores possíveis são: Top, Center, Bottom, Justify, Distributed.|1.7|
|[style](/javascript/api/excel/excel.style)|_Propriedade_ > wrapText|Indica se o Microsoft Excel quebra automaticamente a linha de texto no objeto.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relação_ > borders|Uma coleção Border de quatro objetos Border que representam o estilo das quatro bordas. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relação_ > fill|O preenchimento do estilo. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relação_ > font|Objeto de fonte que representa a fonte do estilo. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Método_ > Delete()|Exclui este estilo.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Propriedade_ > itens|Uma coleção de objetos de estilo. Somente leitura.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Método_ > add(name: string)]|Adiciona um novo estilo para o conjunto.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Método_ > getItem(name: string)|Obtém um estilo por nome.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > address|Obtém o endereço que representa a área alterada de uma tabela em uma planilha específica.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > changeType|Obtém o tipo de mudança que representa como o evento Changed é acionado. Os valores possíveis são: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > source|Obtém a origem do evento. Os valores possíveis são: Local, Remote.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > tableId|Obtém o id da tabela na qual os dados foram alterados.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual os dados são alterados.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > address|Obtém o endereço do intervalo que representa a área selecionada da tabela em uma planilha específica.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > isInsideTable|Indica se a seleção está dentro de uma tabela, o endereço será inútil se IsInsideTable for falso.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > tableId|Obtém o id da tabela na qual a seleção foi alterada.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual a seleção foi alterada.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Propriedade_ > nome|Obtém o nome da pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > dataConnections|Atualiza todas as conexões de dados na pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > properties|Obtém as propriedades da pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > protection|Retorna o objeto de proteção de pasta de trabalho para uma pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > styles|Representa uma coleção de estilos associados à pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Método_ > getActiveCell()|Obtém a célula ativa no momento da pasta de trabalho.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Propriedade_ > protected|Indica se a pasta de trabalho está protegida. Somente Leitura. Somente leitura.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Método_ > protect(password: string)|Protege uma pasta de trabalho. Falhará se a pasta de trabalho estiver protegida.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Método_ > unprotect(password: string)|Desprotege uma pasta de trabalho.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > gridlines|Obtém ou define um sinalizador de linhas de grade da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > headings|É ou define um sinalizador de cabeçalhos da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > showHeadings|É ou define um sinalizador de cabeçalhos da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > standardHeight|Retorna a altura padrão de todas as linhas na planilha, em pontos. Somente leitura.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > standardWidth|Retorna ou define a largura padrão de todas as colunas na planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Propriedade_ > tabColor|Obtém ou define a cor da guia de planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relação_ > freezePanes|Obtém um objeto que pode ser usado para manipular painéis congelados na planilha somente leitura.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > copy(positionType: WorksheetPositionType, relativeTo: Worksheet)|Copia uma planilha e a coloca na posição especificada. Retorna à planilha copiada.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)|Obtém o objeto Range que começa em um determinado índice de linha e índice de coluna e que abrange um determinado número de linhas e colunas.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha que está ativada.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propriedade_ > source|Obtém a origem do evento. Os valores possíveis são: Local, Remote.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha que é adicionada à pasta de trabalho.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > address|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > changeType|Obtém o tipo de mudança que representa como o evento Changed é acionado. Os valores possíveis são: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > source|Obtém a origem do evento. Os valores possíveis são: Local, Remote.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual os dados são alterados.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha que está desativada.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propriedade_ > source|Obtém a origem do evento. Os valores possíveis são: Local, Remote.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Propriedade_ > worksheetId|Obtém o id do gráfico que é excluído da pasta de trabalho.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > freezeAt(frozenRange: Range or string)|Define as células congeladas no modo de exibição da planilha ativa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > freezeColumns(count: number)|Congela a primeira colunas da planilha no local.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > freezeRows(count: number)|Congela as linhas superiores da planilha no local.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > getLocation()|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > getLocationOrNullObject()|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Método_ > unfreeze()|Remove todos os painéis congelados na planilha.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowEditObjects|Indica a opção de proteção de planilha para permitir a edição de objetos.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowEditScenarios|Indica a opção de proteção de planilha para permitir a edição de cenários.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Relação_ > selectionMode|Representa a opção de proteção da planilha do modo de seleção.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propriedade_ > address|Obtém o endereço do intervalo que representa a área selecionada de uma planilha específica.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propriedade_ > type|Obtém o tipo do evento. Os valores possíveis são: WorksheetDataChanged WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Propriedade_ > worksheetId|Obtém o id da planilha na qual a seleção foi alterada.|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>Quais são as novidades na API JavaScript do Excel 1.6 

### <a name="conditional-formatting"></a>Formatação condicional

Introduz a formatação condicional de um intervalo. Permite os seguintes tipos de formatação condicional:

* Escala de cores
* Barra de dados
* Conjunto de ícones
* Personalizado

Além disso:

* Retorna o intervalo ao qual o formatato condicional é aplicada. 
* Remoção da formatação condicional. 
* Fornece a capacidade de priority e stopifTrue. 
* Obtém a coleção de toda a formatação condicional em um determinado intervalo. 
* Limpa todos os formatos condicionais ativos no intervalo atual especificado. 

|Objeto| Quais são as novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Método_ > suspendApiCalculationUntilNextSync()|Suspende o cálculo até que o próximo "context.sync()" seja chamado. Uma vez definido, é responsabilidade do desenvolvedor recalcular a pasta de trabalho, para garantir que todas as dependências sejam propagadas.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relação_ > format|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relação_ > rule|Representa o objeto Regra neste formato condicional.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Propriedade_ > threeColorScale|Caso verdadeiro, a escala de cores terá três pontos (mínimo, médio, máximo). Caso contrário, terá dois (mínimo, máximo). Somente leitura.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Relação_ > criteria|Os critérios da escala de cores. O ponto médio é opcional ao se usar uma escala de cores de dois pontos.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propriedade_ > formula1|A fórmula, se necessário, para avaliar a regra de formatação condicional.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propriedade_ > formula2|A fórmula, se necessário, para avaliar a regra de formatação condicional.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Propriedade_ > operator|O operador do formato condicional de texto. Os valores possíveis são: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relação_ > maximum|O critério de escala de cores de ponto máximo.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relação_ > midpoint|O critério de escala de cores de ponto médio, se a escala de cores for uma escala de três cores.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relação_ > minimum|O critério de escala de cores de ponto mínimo.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propriedade_ > color|Representação de código de cor HTML da cor de escala de cores. Por exemplo, #FF0000 representa vermelho.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propriedade_ > formula|Um número, uma fórmula ou nulo (se Type for LowestValue).|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Propriedade_ > type|No que a fórmula condicional de ícone deve se basear. Os valores possíveis são: Invalid, LowestValue, HighestValue, Number, Percent, Formula, Percentile.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriedade_ > borderColor|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriedade_ > fillColor|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriedade_ > matchPositiveBorderColor|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de borda que o DataBar positivo.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Propriedade_ > matchPositiveFillColor|Representação booliana para indicar se o DataBar negativo tem ou não a mesma cor de preenchimento que o DataBar positivo.|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propriedade_ > borderColor|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propriedade_ > fillColor|Código de cor HTML que representa a cor #RRGGBB do formulário (por exemplo, "FFA500") ou um nome de cor HTML (por exemplo, "laranja").|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Propriedade_ > gradientFill|Representação booliana para indicar se a DataBar tem um gradiente ou não.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Propriedade_ > formula|A fórmula, se necessário, para avaliar a regra databar.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Propriedade_ > type|O tipo de regra para databar. Os valores possíveis são: LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriedade_ > id|A prioridade do formato condicional na atual ConditionalFormatCollection. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriedade_ > priority|A prioridade (ou índice) dentro da coleção de formatos condicionais na qual se encontra atualmente esse formato condicional. Alterando isso também.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriedade_ > stopIfTrue|Se as condições desse formato condicional forem atendidas, nenhum formato de prioridade mais baixa terá efeito nessa célula.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Propriedade_ > type|Um tipo de formatação condicional. É possível definir somente um por vez. Somente leitura. Os valores possíveis são: Custom, DataBar, ColorScale, IconSet.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > cellValue|Retornará as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo de CellValue. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > cellValueOrNullObject|Retornará as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo de CellValue. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > colorScale|Retornará as propriedades de formato condicional de ColorScale se o formato condicional atual for um tipo de ColorScale. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > colorScaleOrNullObject|Retornará as propriedades de formato condicional de ColorScale se o formato condicional atual for um tipo de ColorScale. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > custom|Retornará as propriedades personalizadas do formato condicional se o formato condicional atual for um tipo personalizado. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > customOrNullObject|Retornará as propriedades personalizadas do formato condicional se o formato condicional atual for um tipo personalizado. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > dataBar|Retornará as propriedades da barra de dados se o formato condicional atual for uma barra de dados. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > dataBarOrNullObject|Retornará as propriedades da barra de dados se o formato condicional atual for uma barra de dados. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > iconSet|Retornará as propriedades do formato condicional de IconSet se o formato condicional atual for um tipo de IconSet. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > iconSetOrNullObject|Retornará as propriedades do formato condicional de IconSet se o formato condicional atual for um tipo de IconSet. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > preset|Retornará o formato condicional de critérios predefinidos, como as propriedades above averagebelow averageunique valuescontains blanknonblankerrornoerror.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > presetOrNullObject|Retornará o formato condicional de critérios predefinidos, como as propriedades above averagebelow averageunique valuescontains blanknonblankerrornoerror.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > textComparison|Retornará as propriedades específicas do formato condicional de texto se o formato condicional atual for um tipo de texto.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > textComparisonOrNullObject|Retornará as propriedades específicas do formato condicional de texto se o formato condicional atual for um tipo de texto.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > topBottom|Retornará as propriedades do formato condicional de TopBottom se o formato condicional atual for um tipo de TopBottom. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relação_ > topBottomOrNullObject|Retornará as propriedades do formato condicional de TopBottom se o formato condicional atual for um tipo de TopBottom. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Método_ > Delete()|Exclui esse formato condicional.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Método_ > getRange()|Retornará o intervalo ao qual o formato condicional está aplicado ou um objeto nulo se o intervalo for descontínuo. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Método_ > getRangeOrNullObject()|Retornará o intervalo ao qual o formato condicional está aplicado ou um objeto nulo se o intervalo for descontínuo. Somente leitura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Propriedade_ > itens|Uma coleção de objetos conditionalFormat. Somente leitura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > add(type: string)|Adiciona um novo formato condicional à coleção na prioridade firsttop.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > clearAll()|Limpa todos os formatos condicionais ativos no intervalo atual especificado.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > getCount()|Retorna o número de formatos condicionais na pasta de trabalho. Somente leitura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > getItem(id: string)|Retorna um formato condicional para o ID fornecido.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Método_ > getItemAt(index: número)|Retorna um formato condicional no índice fornecido.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propriedade_ > formula|A fórmula, se necessário, para avaliar a regra de formatação condicional.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propriedade_ > formulaLocal|A fórmula, caso necessário, para avaliar a regra de formatação condicional no idioma do usuário.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Propriedade_ > formulaR1C1|A fórmula, caso necessário, para avaliar a regra de formatação condicional em notação de estilo R1C1.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Propriedade_ > formula|Um número ou uma fórmula, dependendo do tipo.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Propriedade_ > operator|GreaterThan ou GreaterThanOrEqual para cada tipo de regra para o formato de ícone condicional. Os valores possíveis são Invalid, GreaterThan, GreaterThanOrEqual.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relação_ > customIcon|O ícone personalizado para o critério atual, se diferente do IconSet padrão; caso contrário, será retornado nulo.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relação_ > type|No que a fórmula condicional de ícone deve se basear.|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_Propriedade_ > criterion|O critério do formato condicional. Os valores possíveis são: Invalid, Blanks, NonBlanks, Errors, NonErrors, Yesterday, Today, Tomorrow, LastSevenDays, LastWeek, ThisWeek, NextWeek, LastMonth, ThisMonth, NextMonth, AboveAverage, BelowAverage, EqualOrAboveAverage, EqualOrBelowAverage, OneStdDevAboveAverage, OneStdDevBelowAverage, TwoStdDevAboveAverage, TwoStdDevBelowAverage, ThreeStdDevAboveAverage, ThreeStdDevBelowAverage, UniqueValues, DuplicateValues.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriedade_ > color|Código de cor HTML que representa a cor #RRGGBB da linha de borda do formulário (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriedade_ > id|Representa o identificador da borda. Somente leitura. Os valores possíveis são: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriedade_ > sideIndex|Valor constante que indica o lado específico da borda. Somente leitura. Os valores possíveis são: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Propriedade_ > style|Uma das constantes de estilo de linha especificando o estilo de linha da borda. Os valores possíveis são: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Propriedade_ > count|Número de objetos de borda da coleção. Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Propriedade_ > itens|Uma coleção de objetos conditionalRangeBorder. Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relação_ > bottom|Torna a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relação_ > left|Torna a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relação_ > right|Torna a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relação_ > top|Torna a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Método_ > getItem(index: string)|Obtém um objeto de borda usando seu nome|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Método_ > getItemAt(index: número)|Obtém um objeto de borda usando seu índice.|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Propriedade_ > color|Código de cor HTML que representa a cor do preenchimento do formulário #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Método_ > clear()|Redefine o preenchimento.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > bold|Representa o status da fonte em negrito.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > color|Representação de código de cor HTML para a cor do texto. Por exemplo, #FF0000 representa vermelho.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > italic|Representa o status da fonte em itálico.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > strikethrough|Representa o status de tachado da fonte.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Propriedade_ > underline|Tipo de sublinhado aplicado à fonte. Os valores possíveis são: None, Single, Double.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Método_ > clear()|Redefine os formatos de fonte.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Propriedade_ > numberFormat|Representa o código de formato numérico do Excel para determinado intervalo. Desmarcado se o nulo for passado.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relação_ > borders|Coleção de objetos de borda que se aplicam ao intervalo de formatos condicionais geral. Somente leitura.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relação_ > fill|Retorna o objeto de preenchimento definido no intervalo de formatos condicionais gerais. Somente leitura.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relação_ > font|Retorna o objeto de fonte definido no intervalo de formatos condicionais gerais. Somente leitura.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Propriedade_ > operator|O operador do formato condicional de texto. Os valores possíveis são: Invalid, Contains, NotContains, BeginsWith, EndsWith.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Propriedade_ > text|O valor de texto do formato condicional.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Propriedade_ > rank|A classificação entre 1 e 1000 para classificações numéricas ou 1 e 100 para classificações percentuais.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Propriedade_ > type|Formatar valores com base na classificação superior ou inferior. Os valores possíveis são: Invalid, TopItems, TopPercent, BottomItems, BottomPercent.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relação_ > format|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relação_ > rule|Representa o objeto Regra neste formato condicional. Somente leitura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriedade_ > axisColor|Código de cor HTML que representa a cor da linha de Eixo, no formato #RRGGBB (por exemplo, "FFA500") ou uma cor HTML nomeada (por exemplo, "laranja").|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriedade_ > axisFormat|Representação de como o eixo é determinado para uma barra de dados do Excel. Os valores possíveis são: Automatic, None, CellMidPoint.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriedade_ > barDirection|Representa a direção em que o gráfico de barras de dados deve ser baseado. Os valores possíveis são: Context, LeftToRight, RightToLeft.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Propriedade_ > showDataBarOnly|Caso verdadeiro, oculta os valores das células às quais a barra de dados é aplicada.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relação_ > lowerBoundRule|A regra para o que constitui o limite inferior (e como calculá-lo, se aplicável) para uma barra de dados.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relação_ > negativeFormat|Representação de todos os valores à esquerda do eixo em uma barra de dados do Excel. Somente leitura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relação_ > positiveFormat|Representação de todos os valores à direita do eixo em uma barra de dados do Excel. Somente leitura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relação_ > upperBoundRule|A regra para o que constitui o limite superior (e como calculá-lo, se aplicável) para uma barra de dados.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propriedade_ > reverseIconOrder|Caso verdadeiro, inverte as ordens de ícones para IconSet. Observe que não será possível definir isso se ícones personalizados forem usados.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propriedade_ > showIconOnly|Caso verdadeiro, oculta os valores e mostra somente ícones.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Propriedade_ > style|Caso definido, exibe a opção IconSet do formato condicional. Os valores possíveis são: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Relação_ > criteria|Uma matriz de IconSets e critérios para as regras e os ícones personalizados potenciais para ícones condicionais. Observe que, para o primeiro critério, apenas o ícone personalizado pode ser modificado, enquanto tipo, fórmula e operador serão ignorados quando definidos.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relação_ > format|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relação_ > rule|A regra da formatação condicional.|1.6|
|[range](/javascript/api/excel/excel.range)|_Relação_ > conditionalFormats|Coleção de ConditionalFormats que formam uma interseção do intervalo. Somente leitura.|1.6|
|[range](/javascript/api/excel/excel.range)|_Método_ > calculate()|Calcula um intervalo de células em uma planilha.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relação_ > format|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relação_ > rule|A regra da formatação condicional.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relação_ > format|Retorna um objeto de formato, encapsulando a fonte, o preenchimento, as bordas e outras propriedades de formatos condicionais. Somente leitura.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relação_ > rule|Os critérios da formatação condicional TopBottom.|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > internalTest|Somente para uso interno. Somente leitura.|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > calculate(markAllDirty: bool)|Calcula todas as células em uma planilha.|1.6|

##  <a name="whats-new-in-excel-javascript-api-15"></a>Quais são as novidades na API JavaScript do Excel 1.5

### <a name="custom-xml-part"></a>Parte XML personalizada

* Adição de uma coleção de partes XML personalizadas ao objeto workbook.
* Obter parte XML personalizada usando ID
* Obtenção de um novo conjunto com escopo de partes XML personalizadas cujos namespaces correspondam ao namespace especificado.
* Obtenha uma cadeia XML associada a uma parte.
* Forneça id e namespace de uma parte.
* Adiciona uma nova parte XML personalizada à pasta de trabalho.
* Defina a parte XML inteira.
* Exclua uma parte XML personalizada.
* Exclua um atributo com o nome especificado do elemento identificado por xpath.
* Consulte o conteúdo XML por xpath.
* Insira, atualize e exclua o atributo.

**Implementação de referência:** Consulte [aqui](https://github.com/mandren/Excel-CustomXMLPart-Demo) para conhecer uma implementação de referência que mostra como partes XML personalizadas podem ser usadas em um suplemento.

### <a name="others"></a>Outros
* `range.getSurroundingRegion()` Retorna um objeto Range que representa a região ao redor desse intervalo. Uma região ao redor é um intervalo limitado por qualquer combinação de linhas e colunas em branco em relação a esse intervalo.
* `getNextColumn()` e `getPreviousColumn()`, `getLast() na coluna da tabela.
* `getActiveWorksheet()` na pasta de trabalho.
* `getRange(address: string)` fora da pasta de trabalho.
* `getBoundingRange(ranges: )` Obtém o menor objeto Range que abrange os intervalos fornecidos. Por exemplo, o intervalo delimitador entre "B2:C5" e "D10:E15" é "B2:E15".
* `getCount()` em várias coleções, como itens nomeados, planilhas, tabelas etc. para obter o número de itens em uma coleção. `workbook.worksheets.getCount()`
* `getFirst()` e `getLast()` e obter o último em várias coleções, como coleções de planilhas, colunas de tabela, pontos de gráfico e exibições de intervalo.
* `getNext()` e `getPrevious()` na coleção de planilhas e colunas de tabela.
* `getRangeR1C1()` Obtém o objeto Range que começa em um determinado índice de linha e índice de coluna e que abrange um determinado número de linhas e colunas.

|Objeto| Quais são as novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[customXMLPart](/javascript/api/excel/excel.customxmlpart)|_Propriedade_ > id|ID da parte XML personalizada. Somente leitura.|1,5|
|[customXMLPart](/javascript/api/excel/excel.customxmlpart)|_Propriedade_ > namespaceUri|URI do namespace da parte XML personalizada. Somente leitura.|1,5|
|[customXMLPart](/javascript/api/excel/excel.customxmlpart)|_Método_ > Delete()|Exclui a parte XML personalizada.|1,5|
|[customXMLPart](/javascript/api/excel/excel.customxmlpart)|_Método_ > getXml()|Obtém o conteúdo XML completo da parte XML personalizada.|1,5|
|[customXMLPart](/javascript/api/excel/excel.customxmlpart)|_Método_ > setXml(xml: string)|Define o conteúdo XML completo da parte XML personalizada.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Propriedade_ > itens|Uma coleção de objetos customXmlPart. Somente leitura.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > add(xml: string)|Adiciona uma nova parte XML personalizada à pasta de trabalho.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getByNamespace(namespaceUri: string)|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getCount()|Obtém o número de partes CustomXml na coleção.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getItem(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1,5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Método_ > getItemOrNullObject(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Propriedade_ > itens|Uma coleção de objetos customXmlPartScoped. Somente leitura.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getCount()|Obtém o número de partes CustomXML nesta coleção.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getItem(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getItemOrNullObject(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getOnlyItem()|Se o conjunto contiver exatamente um item, esse método o retornará.|1,5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Método_ > getOnlyItemOrNullObject()|Se o conjunto contiver exatamente um item, esse método o retornará.|1,5|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > customXmlParts|Representa a coleção de partes XML contidas nesta pasta de trabalho. Somente leitura.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getNext(visibleOnly: bool)|Obtém a planilha posterior a esta. Se não houver nenhuma planilha após esta, este método gerará um erro.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getNextOrNullObject(visibleOnly: bool)|Obtém a planilha posterior a esta. Se não houver nenhuma planilha após esta, este método retornará um objeto nulo.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getPrevious(visibleOnly: bool)|Obtém a planilha anterior a esta. Se não houver nenhuma planilha anterior, esse método lançará um erro.|1,5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getPreviousOrNullObject(visibleOnly: bool)|Obtém a planilha anterior a esta. Se não houver nenhuma planilha anterior, este método retornará um objeto nulo.|1,5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getFirst(visibleOnly: bool)|Obtém a primeira planilha na coleção.|1,5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getLast(visibleOnly: bool)|Obtém a última planilha na coleção.|1,5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Quais são as novidades na API JavaScript do Excel 1.4
A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.4.

### <a name="named-item-add-and-new-properties"></a>Adicionar item nomeado e novas propriedades

Novas propriedades:

* `comment`
* `scope` itens com escopo de planilha ou pasta de trabalho
* `worksheet` retorna a planilha que o item nomeado tem como escopo.

Novos métodos:

* `add(name: string, reference: Range or string, comment: string)`Adiciona um novo nome à coleção do escopo fornecido.
* `addFormulaLocal(name: string, formula: string, comment: string)` Adiciona um novo nome à coleção do escopo fornecido usando a localidade do usuário para a fórmula.

### <a name="settings-api-in-the-excel-namespace"></a>Configurações de API no namespace do Excel

O objeto [Configuração](/javascript/api/excel/excel.setting) representa um par chave-valor de uma configuração persistente ao documento. O recurso `Excel.Setting` é equivalente a `Office.Settings`, mas usa a sintaxe da API em lote, em vez de modelo de retorno de chamada de API comuns.

As APIs incluem `getItem()` para acessar configuração de entrada por meio da chave, `add()` para adicionar o par de configuração de chave:valor especificado na pasta de trabalho.

### <a name="others"></a>Outros

* Definir nome de coluna de tabela (a versão anterior permite somente leitura).
* Adicionar coluna de tabela ao fim da tabela (a versão anterior permite apenas em qualquer lugar, exceto o último).
* Adicione várias linhas a uma tabela de cada vez (a versão anterior só permite uma linha por vez).
* `range.getColumnsAfter(count: number)` e `range.getColumnsBefore(count: number)` para obter determinado número de colunas à direita/esquerda do objeto Range atual.
* Obter item ou função de objeto null: Esta funcionalidade permite obter o objeto utilizando a chave. Se o objeto não existir, a propriedade isNullObject do objeto retornado será true. Isso permite que os desenvolvedores verifiquem se existe um objeto ou não sem ter de lidar com ele por meio do tratamento de exceção. Disponível na planilha, item nomeado, associação, série de gráficos etc.

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|Objeto| Quais são as novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > getCount()|Obtém o número de associações da coleção.|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > getItemOrNullObject(id: string)|Obtém um objeto binding pela ID. Se o objeto binding não existir, retornará um objeto null.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Método_ > getCount()|Retorna o número de gráficos da planilha.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Método_ > getItemOrNullObject(name: string)|Obtém um gráfico usando o respectivo nome. Quando houver vários gráficos com o mesmo nome, o sistema retornará o primeiro deles.|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_Método_ > getCount()|Retorna o número de pontos do gráfico da série.|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Método_ > getCount()|Retorna o número de série da coleção.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propriedade_ > comment|Representa o comentário associado a esse nome.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Propriedade_ > escopo|Indica se o nome tem escopo para a pasta de trabalho ou uma planilha específica. Somente leitura. Os valores possíveis são: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relação_ > planilha|Retorna a planilha em que o item nomeado tem escopo. Gerará um erro se os itens tiverem escopo para a pasta de trabalho em vez disso. Somente leitura.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relação_ > worksheetOrNullObject|Retorna a planilha em que o item nomeado tem escopo. Retornará um objeto null se o item tiver escopo para a pasta de trabalho em vez disso. Somente leitura.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Método_ > Delete()|Exclui o nome fornecido.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Método_ > getRangeOrNullObject()|Retorna o objeto Range associado ao nome. Retornará um objeto null se o tipo do item nomeado não for um intervalo.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > add(name: string, reference: Range or string, comment: string)|Adiciona um novo nome à coleção do escopo fornecido.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > addFormulaLocal (name: string, formula: string, comment: string)|Adiciona um novo nome à coleção de escopo fornecido usando a localidade do usuário para a fórmula.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > getCount()|Obtém o número de itens nomeados na coleção.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > getItemOrNullObject(name: string)|Obtém um objeto NamedItem usando o respectivo nome. Se o objeto getNamedItem não existir, retornará um objeto null.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getCount()|Obtém o número de tabelas dinâmicas na coleção.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getItemOrNullObject(name: string)|Obtém uma Tabela Dinâmica por nome. Se a tabela dinâmica não existir, retornará um objeto null.|1.4|
|[range](/javascript/api/excel/excel.range)|_Método_ > getIntersectionOrNullObject(anotherRange: Intervalo ou cadeia de caracteres)|Obtém o objeto de intervalo que representa a interseção retangular dos intervalos determinados. Se nenhuma interseção for encontrada, retornará um objeto null.|1.4|
|[range](/javascript/api/excel/excel.range)|_Método_ > getUsedRangeOrNullObject(valuesOnly: bool)|Retorna o intervalo usado do objeto range determinado. Se não houver nenhuma célula usada no intervalo, esta função retornará um objeto null.|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Método_ > getCount()|Obtém o número de objetos RangeView na coleção.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Propriedade_ > key|Retorna a chave que representa a id da configuração. Somente leitura.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Propriedade_ > value|Representa o valor armazenado para esta configuração.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Método_ > Delete()|Exclui a configuração.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Propriedade_ > itens|Uma coleção de objetos de configuração. Somente leitura.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > add(key: string, value: (any))|Define na pasta de trabalho ou adiciona a ela a configuração especificada.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getCount()|Obtém o número de Configurações na coleção.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItem(key: string)|Obtém uma entrada de configuração por meio da tecla.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItemOrNullObject(key: string)|Obtém uma entrada de configuração por meio da tecla. Se a Configuração não existir, retornará um objeto null.|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relação_ > settings|Obtém o objeto Setting, que representa as associações que geraram o evento settingsChanged.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Método_ > getCount()]|Obtém o número de tabelas na coleção.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Método_ > getItemOrNullObject(key: number or string)|Obtém uma tabela pelo nome ou ID. Se a tabela não existir, retornará um objeto null.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Método_ > getCount()|Obtém a quantidade de colunas na tabela.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Método_ > getItemOrNullObject(key: number or string)|Obtém um objeto column por nome ou ID. Se a coluna não existir, retornará um objeto null.|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_Método_ > getCount()|Obtém a quantidade de linhas na tabela.|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > settings|Representa uma coleção de configurações associada à pasta de trabalho. Somente leitura.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relação_ > nomes|Coleção de nomes com escopo para a planilha atual. Somente leitura.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Método_ > getUsedRangeOrNullObject(valuesOnly: bool)|O intervalo usado é o menor intervalo que abrange todas as células que têm um valor ou uma formatação atribuída a elas. Se a planilha inteira estiver em branco, esta função retornará um objeto null.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getCount(visibleOnly: bool)|Obtém o número de planilhas na coleção.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Método_ > getItemOrNullObject(key: string)|Obtém um objeto worksheet usando o Nome ou ID dele. Se a planilha não existir, retornará um objeto null.|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Quais são as novidades na API JavaScript do Excel 1.3

A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.3.

|Objeto| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_Método_ > Delete()|Especifica a associação.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > add(range: Range or string, bindingType: string, id: string)|Adiciona uma nova associação a um intervalo específico.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > addFromNamedItem(name: string, bindingType: string, id: string)|Adiciona uma nova associação com base em um item nomeado na pasta de trabalho.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > addFromSelection(bindingType: string, id: string)|Adiciona uma nova associação com base na seleção atual.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > getItemOrNull(id: string)|Obtém um objeto de associação pela ID. Se o objeto em associação não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Método_ > getItemOrNull(name: string)|Obtém um gráfico usando o respectivo nome. Quando houver vários gráficos com o mesmo nome, o sistema retornará o primeiro deles.|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Método_ > getItemOrNull(name: string)|Obtém um objeto NamedItem usando o respectivo nome. Se o objeto nameditem não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Propriedade_ > nome|Nome da Tabela Dinâmica.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relação_ > planilha|A planilha que contém a Tabela Dinâmica atual. Somente leitura.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Método_ > refresh()|Atualiza a Tabela Dinâmica.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Propriedade_ > itens|Uma coleção de objetos de Tabela Dinâmica. Somente leitura.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getItem(name: string)|Obtém uma Tabela Dinâmica por nome.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Método_ > getItemOrNull(name: string)|Obtém uma Tabela Dinâmica por nome. Se a Tabela Dinâmica não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[range](/javascript/api/excel/excel.range)|_Método_ > getIntersectionOrNull(anotherRange: Range or string)|Obtém o objeto de intervalo que representa a interseção retangular dos intervalos determinados. Se nenhuma interseção for encontrada, retornará um objeto null.|1.3|
|[range](/javascript/api/excel/excel.range)|_Método_ > getVisibleView()|Representa as linhas visíveis do intervalo atual.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > cellAddresses|Representa os endereços de célula da RangeView. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > columnCount|Retorna o número de colunas visíveis. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > formulas|Representa a fórmula em notação A1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > formulasLocal|Representa a fórmula em notação A1, na formatação de número da localidade e no idioma do usuário.  Por exemplo, a fórmula "=SUM(A1, introduced in 1.5)" em inglês seria "=SOMA(A1; 1,5)" em português.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > formulasR1C1|Representa a fórmula em notação no estilo L1C1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > index|Retorna um valor que representa o índice da RangeView. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > numberFormat|Representa o código de formato de número do Excel para determinada célula.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > rowCount|Retorna o número de linhas visíveis. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > texto|Valores de texto do intervalo especificado. O valor de texto não depende da largura da célula. A substituição pelo sinal #, que ocorre na interface de usuário do Excel, não afeta o valor de texto retornado pela API. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > valueTypes|Representa o tipo de dados de cada célula. Somente leitura. Os valores possíveis são: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Propriedade_ > values|Representa os valores brutos da exibição do intervalo especificado. Os dados retornados podem ser dos tipos: cadeia de caracteres, número ou booliano. Células que contêm um erro retornarão a cadeia de caracteres de erro.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Relação_ > rows|Representa uma coleção de exibições de tabelas associadas ao intervalo. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Método_ > getRange()|Obtém o intervalo pai associado à RangeView atual.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Propriedade_ > itens|Uma coleção de objetos rangeView. Somente leitura.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Método_ > getItemAt(index: número)|Obtém uma linha de RangeView através de seu índice. Indexado com zero.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Propriedade_ > key|Retorna a chave que representa a id da configuração. Somente leitura.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Método_ > Delete()|Exclui a configuração.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Propriedade_ > itens|Uma coleção de objetos de configuração. Somente leitura.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItem(key: string)|Obtém uma entrada de configuração por meio da tecla.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > getItemOrNull(key: string)|Obtém uma entrada de configuração por meio da tecla. Se o objeto de configuração não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Método_ > set(key: string, value: string)|Define na pasta de trabalho ou adiciona a ela a configuração especificada.|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relação_ > settingCollection|Obtém o objeto Setting, que representa as associações que geraram o evento settingsChanged.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > highlightFirstColumn|Indica se a primeira coluna contém uma formatação especial.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > highlightLastColumn|Indica se a última coluna contém uma formatação especial.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > showBandedColumns|Indica se as colunas mostram formatação em faixas nas quais as colunas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > showBandedRows|Indica se as linhas mostram formatação em faixas nas quais as linhas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|1.3|
|[table](/javascript/api/excel/excel.table)|_Propriedade_ > showFilterButton|Indica se os botões de filtro estão visíveis na parte superior de cada cabeçalho da coluna. Essa configuração só será permitida se a tabela tiver uma linha de cabeçalho.|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Método_ > getItemOrNull(key: number or string)|Obtém uma tabela pelo nome ou ID. Se a tabela não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Método_ > getItemOrNull(key: number or string)|Obtém um objeto de coluna por nome ou ID. Se a coluna não existir, a propriedade do objeto isNull retornado será true.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > pivotTables|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho. Somente leitura.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > settings|Representa uma coleção de configurações associada à pasta de trabalho. Somente leitura.|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relação_ > pivotTables|Coleção de Tabelas Dinâmicas que fazem parte da planilha. Somente leitura.|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Quais são as novidades na API JavaScript do Excel 1.2

A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.2.

|Objeto| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > id|Obtém um gráfico com base em sua posição na coleção. Somente leitura.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Relação_ > planilha|A planilha que contém o gráfico atual. Somente leitura.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Método_ > getImage(height: number, width: number, fittingMode: string)|Processa o gráfico como uma imagem codificada em base64, dimensionando o gráfico para se ajustar às dimensões especificadas.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Relação_ > criteria|O filtro aplicado no momento à coluna fornecida. Somente leitura.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > apply(criteria: FilterCriteria)|Aplica os critérios de filtro determinados à coluna fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyBottomItemsFilter(count: number)|Aplica um filtro "Item Inferior" à coluna para obter o número de elementos fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyBottomPercentFilter(percent: number)]|Aplica um filtro "Percentual Inferior" à coluna para obter a porcentagem de elementos fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyCellColorFilter(color: string)|Aplica um filtro "Cor da Célula" à coluna para obter a cor fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyCustomFilter (criteria1: string, criteria2: string, oper: string)|Aplica um filtro "Ícone" à coluna para obter as cadeias de caracteres de critérios fornecidas.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyDynamicFilter(criteria: string)|Aplica um filtro "Dinâmico" à coluna.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyFontColorFilter(color: string)|Aplica um filtro "Cor da Fonte" à coluna para obter a cor fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyIconFilter(icon: Icon)|Aplica um filtro "Ícone" à coluna para obter o ícone fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyTopItemsFilter(count: number)|Aplica um filtro "Item Superior" à coluna para obter o número de elementos fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyTopPercentFilter(percent: number)|Aplica um filtro "Percentual Superior" à coluna para obter a porcentagem de elementos fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > applyValuesFilter (valores: ())|Aplica um filtro "Valores" à coluna para obter os valores fornecidos.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Método_ > clear()|Limpa o filtro na coluna fornecida.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > color|A cadeia HTML de cor usada para filtrar células. Usada com a filtragem "cellColor" e "fontColor".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > criterion1|O primeiro critério usado para filtrar os dados. Usado como um operador no caso de filtragem "custom".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > criterion2|O segundo critério usado para filtrar os dados. Só é usado como um operador no caso de filtragem "custom".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > dynamicCriteria|Os critérios dinâmicos do conjunto Excel.DynamicFilterCriteria a serem aplicados nessa coluna. Usados com a filtragem "dynamic". Os valores possíveis são: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > filterOn|A propriedade usada pelo filtro para determinar se os valores devem ficar visíveis. Os valores possíveis são: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > operator|O operador usado para combinar o critério 1 e 2 ao usar a filtragem "custom". Os valores possíveis são: "And", "Or".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Propriedade_ > values|O conjunto de valores a serem usados como parte da filtragem "values".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Relação_ > icon|O ícone usado para filtrar células. Usado com a filtragem "icon".|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Propriedade_ > date|A data no formato ISO8601 usada para filtrar os dados.|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Propriedade_ > specificity|Como a data específica deve ser usada para manter os dados. Por exemplo, se a data for 2005-04-02 e a especificidade estiver definida como "mês", a operação de filtragem manterá todas as linhas com uma data do mês de abril de 2009. Os valores possíveis são: Ano, segunda-feira, dia, hora, minuto, segundo.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Propriedade_ > formulaHidden|Indica se o Excel ocultará a fórmula para as células no intervalo. Um valor nulo indica que o intervalo inteiro não tem configuração uniforme de fórmula oculta.|1.2|
|[formatProtection](/javascript/api/excel/excel.formatprotection)|_Propriedade_ > locked|Indica se o Excel bloqueia as células no objeto. Um valor nulo indica que o intervalo inteiro não tem configuração de bloqueio uniforme.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Propriedade_ > index|Representa o índice do ícone no conjunto fornecido.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Propriedade_ > set|Representa o conjunto do qual ícone faz parte. Os valores possíveis são: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > columnHidden|Representa se todas as colunas do intervalo atual estão ocultas.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > formulasR1C1|Representa a fórmula em notação no estilo L1C1.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > hidden|Representa se todas as células do intervalo atual estão ocultas. Somente leitura.|1.2|
|[range](/javascript/api/excel/excel.range)|_Propriedade_ > rowHidden|Representa se todas as linhas do intervalo atual estão ocultas.|1.2|
|[range](/javascript/api/excel/excel.range)|_Relação_ > sort|Representa a classificação de intervalo do intervalo atual. Somente leitura.|1.2|
|[range](/javascript/api/excel/excel.range)|_Método_ > merge(across: bool)|Mescla as células do intervalo em uma região da planilha.|1.2|
|[range](/javascript/api/excel/excel.range)|_Método_ > unmerge()|Desfaz a mesclagem das células do intervalo em células separadas.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > columnWidth|Obtém ou define a largura de todas as colunas dentro do intervalo. Se as larguras das colunas não forem uniformes, será retornado null.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Propriedade_ > rowHeight|Obtém ou define a altura de todas as linhas do intervalo. Se as alturas das linhas não forem uniformes, será retornado null.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Relação_ > protection|Retorna o objeto de proteção de formato para um intervalo. Somente leitura.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Método_ > autofitColumns()|Altera a largura das colunas do intervalo atual para obter o melhor ajuste, com base nos dados atuais nas colunas.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Método_ > autofitRows()|Altera a altura das linhas do intervalo atual para obter o melhor ajuste, com base nos dados atuais nas colunas.|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_Propriedade_ > address|Representa as linhas visíveis do intervalo atual.|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_Método_apply(fields: SortField, matchCase: bool, hasHeaders: bool, orientation: string, method: string)|Executa uma operação de classificação.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > ascending|Indica se a classificação é feita de forma crescente.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > color|Representa a cor que é o destino da condição se a classificação estiver na cor da fonte ou da célula.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > dataOption|Representa as opções de classificação adicionais para esse campo. Os valores possíveis são: Normal, TextAsNumber.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > key|Representa a coluna (ou linha, dependendo da orientação da classificação) em que a condição está. Representado como um deslocamento da primeira coluna (ou linha).|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Propriedade_ > sortOn|Representa o tipo de classificação dessa condição. Os valores possíveis são: Valor, CellColor, FontColor, Ícone.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Relação_ > icon|Representa o ícone que é o destino da condição se a classificação está no ícone da célula.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relação_ > sort|Representa a classificação da tabela. Somente leitura.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relação_ > planilha|A planilha que contém a tabela atual. Somente leitura.|1.2|
|[table](/javascript/api/excel/excel.table)|_Método_ > clearFilters()|Limpa todos os filtros aplicados à tabela no momento.|1.2|
|[table](/javascript/api/excel/excel.table)|_Método_ > convertToRange()|Converte a tabela em um intervalo de células normal. Todos os dados são preservados.|1.2|
|[table](/javascript/api/excel/excel.table)|_Método_ > reapplyFilters()|Aplica novamente todos os filtros à tabela.|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_Relação_ > filter|Recupera o filtro aplicado à coluna. Somente leitura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Propriedade_ > matchCase|Indica se o uso de maiúsculas ou minúsculas afetou a última classificação da tabela. Somente leitura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Propriedade_ > method|Indica o último método de ordenação de caracteres chineses usado para classificar a tabela. Somente leitura. Os valores possíveis são: PinYin, StrokeCount.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Relação_ > fields|Representa as condições atuais usadas para a última classificação da tabela. Somente leitura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Método_apply(fields: SortField[], matchCase: bool, method: string)|Executa uma operação de classificação.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Método_ > clear()|Limpa a classificação que está na tabela. Essa ação não modifica a ordenação da tabela, mas limpa o estado dos botões do cabeçalho.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Método_ > reapply()|Reaplica os parâmetros de classificação atuais à tabela.|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_Relação_ > funções|Representa uma instância de aplicativo do Excel que contém essa pasta de trabalho. Somente leitura.|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relação_ > protection|Retorna o objeto de proteção da planilha para uma planilha. Somente leitura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Propriedade_ > protected|Indica se a planilha está protegida. Somente Leitura. Somente leitura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Relação_ > options|Opções de proteção da planilha. Somente leitura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Método_ > protect(options: WorksheetProtectionOptions)|Protege uma planilha. Falhará se uma planilha estiver protegida.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Método_ > unprotect()|Desprotege uma planilha.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowAutoFilter|Indica a opção de proteção de planilha para permitir a utilização do recurso de filtro automático.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowDeleteColumns|Indica a opção de proteção de planilha para permitir a exclusão de colunas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowDeleteRows|Indica a opção de proteção de planilha para permitir a exclusão de linhas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowFormatCells|Indica a opção de proteção de planilha para permitir a formatação de células.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowFormatColumns|Indica a opção de proteção de planilha para permitir a formatação de colunas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowFormatRows|Indica a opção de proteção de planilha para permitir a formatação de linhas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowInsertColumns|Indica a opção de proteção de planilha para permitir a inserção de colunas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowInsertHyperlinks|Indica a opção de proteção de planilha para permitir a inserção de hiperlinks.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowInsertRows|Indica a opção de proteção de planilha para permitir a inserção de linhas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowPivotTables|Indica a opção de proteção de planilha para permitir a utilização do recurso de Tabela Dinâmica.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Propriedade_ > allowSort|Indica a opção de proteção de planilha para permitir a utilização do recurso de classificação.|1.2|

## <a name="excel-javascript-api-11"></a>API JavaScript do Excel 1.1

A API JavaScript do Excel 1.1 é a primeira versão da API. Para saber mais sobre a API, confira [a API JavaScript do Excel](/javascript/api/excel) nos tópicos de referência.

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos da API e de hosts do Office](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos Suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
