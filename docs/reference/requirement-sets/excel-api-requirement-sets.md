# <a name="excel-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Excel

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação em tempo de execução para determinar se um host do Office oferece suporte a APIs que um suplemento precisa. Para obter mais informações, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Os suplementos do Excel são executados em várias versões do Office, incluindo o Office 2016 ou posterior para Windows, Office para iPad, Office para Mac e Office Online. A tabela a seguir lista os conjuntos de requisito do Excel, os aplicativos host do Office que suportam a cada conjunto de requisitos e as versões de compilação ou número para esses aplicativos.

> [!NOTE]
> Qualquer API que está marcada como **Beta** não está pronta para a produção de usuário final. Podemos disponibilizá-las para que os desenvolvedores as testem em ambientes de teste e de desenvolvimento. Elas não devem ser usadas contra documentos críticos de produção/de negócios.
> 
> Para os conjuntos de requisitos que são marcados como **Beta**, use a versão especificada (ou posterior) do software do Office e também a biblioteca Beta da CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Entradas não marcadas como **Beta** geralmente estão disponíveis e você pode usar a biblioteca de Produção na CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.

|  Conjunto de requisitos  |  Office 365 para Windows\*  |  Office 365 para iPad  |  Office 365 para Mac  | Office Online  |  Servidor do Office Online  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Beta  | [Visite nossa página sobre a especificação aberta da API JavaScript do Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)! |
| ExcelApi1.8  | Versão 1808 (Build 10730.20102) ou posterior | 2.17 ou posterior | 16.17 ou posterior | Setembro de 2018 | Em breve |
| ExcelApi1.7  | Versão 1801 (Build 9001.2171) ou posterior   | 2.9 ou posterior | 16.9 ou posterior | Abril de 2018 | Em breve |
| ExcelApi1.6  | Versão 1704 (Build 8201.2001) ou posterior   | 2.2 ou posterior |15.36 ou posterior| Abril de 2017 | Em breve|
| ExcelApi1.5  | Versão 1703 (Build 8067.2070) ou posterior   | 2.2 ou posterior |15.36 ou posterior| Março de 2017 | Em breve|
| ExcelApi1.4  | Versão 1701 (Build 7870.2024) ou posterior   | 2.2 ou posterior |15.36 ou posterior| Janeiro de 2017 | Em breve|
| ExcelApi1.3  | Versão 1608 (Build 7369.2055) ou posterior | 1.27 ou posterior |  15.27 ou posterior| Setembro de 2016 | Versão 1608 (Build 7601.6800) ou posterior|
| ExcelApi1.2  | Versão 1601 (Build 6741.2088) ou posterior | 1.21 ou posterior | 15.22 ou posterior| Janeiro de 2016 ||
| ExcelApi1.1  | Versão 1509 (Build 4266.1001) ou posterior | 1.19 ou posterior | 15.20 ou posterior| Janeiro de 2016 ||

> [!NOTE]
> O número da versão para Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão contém apenas o conjunto de requisitos ExcelApi 1.1.

Para obter mais informações sobre versões, números da versão e servidor do Office Online, consulte:

- [Versão e números da versão de lançamentos do canal de atualizações para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Qual versão do Office estou usando?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Onde você pode encontrar a versão e o número da versão de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Visão geral sobre o Servidor do Office Online](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="whats-new-in-excel-javascript-api-18"></a>Novidades na API JavaScript do Excel 1.8

Os recursos do conjunto de requisitos da API JavaScript do Excel 1.8 incluem APIs para tabelas dinâmicas, validação de dados, gráficos, eventos para gráficos, opções de desempenho e criação da pasta de trabalho.

### <a name="pivottable"></a>Tabela dinâmica

A onda 2 das APIs da Tabela Dinâmica permite que os suplementos definam as hierarquias de uma Tabela Dinâmica. Agora você pode controlar os dados e o modo como eles são agregados. Nosso [artigo sobre Tabela Dinâmica](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-pivottables) tem mais informações sobre a nova funcionalidade da Tabela Dinâmica.

### <a name="data-validation"></a>Validação de dados

A validação de dados permite que você controle o que o usuário insere em uma planilha. Você pode limitar as células para conjuntos de resposta predefinida ou dar avisos pop-up sobre uma entrada indesejável. Saiba mais sobre como [adicionar a validação de dados aos intervalos](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-data-validation) atualmente.

### <a name="charts"></a>Gráficos

Outra rodada das APIs de gráfico traz ainda maior controle programático sobre elementos do gráfico. Agora, você tem maior acesso à legenda, eixos, linha de tendência e área de plotagem.

### <a name="events"></a>Eventos

Mais [eventos](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events) foram adicionados aos gráficos. Faça seu suplemento reagir aos usuários interagindo com o gráfico. Você também pode disparar [eventos de alternância](https://docs.microsoft.com/office/dev/add-ins/excel/performance#enable-and-disable-events) em toda a pasta de trabalho.


|Object| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Method_ > [createWorkbook(base64File: string)](/javascript/api/excel/excel.application)|Cria uma nova pasta de trabalho oculta usando um arquivo .xlsx  codificado em base64 opcional.|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Property_ > formula1|Obtém ou define a Formula1, ou seja, o valor mínimo ou o valor dependente do operador.|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Property_ > formula2|Obtém ou define a Formula2, ou seja, o valor máximo ou o valor dependente do operador.|1.8|
|[basicDataValidation](/javascript/api/excel/excel.basicdatavalidation)|_Relationship_ > operator|O operador a ser usado para validar os dados.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > categoryLabelLevel|Retorna ou define uma constante da enumeração ChartCategoryLabelLevel referindo-se ao nível de onde os rótulos de categoria estão sendo originados. Leitura/gravação.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Propriedade_ > plotVisibleOnly|Verdadeiro se apenas as células visíveis forem plotadas. Falso se tanto as células visíveis quanto as ocultas forem plotadas. ReadWrite.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > seriesNameLevel|Retorna ou define uma constante da enumeração ChartSeriesNameLevel referindo-se ao nível de onde os nomes da série estão sendo originados. Leitura/gravação.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > showDataLabelsOverMaximum|Representa se deve mostrar os rótulos de dados quando o valor é maior que o valor máximo no eixo do valor.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > style|Retorna ou define o estilo de gráfico para o gráfico. ReadWrite.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Relationship_ > displayBlanksAs|Retorna ou define a forma como as células vazias são plotadas em um gráfico. ReadWrite.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Relationship_ > plotArea|Representa a plotArea do gráfico. Somente leitura.|1.8|
|[chart](/javascript/api/excel/excel.chart)|_Relationship_ > plotBy|Retorna ou define a maneira como as linhas ou colunas são usadas como uma série de dados no gráfico. ReadWrite.|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Property_ > chartId|Obtém a identificação do gráfico que é ativado.|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Property_ > type|Obtém o tipo de evento.|1.8|
|[chartActivatedEventArgs](/javascript/api/excel/excel.chartactivatedeventargs)|_Property_ > worksheetId|Obtém a identificação da planilha em que o gráfico é ativado.|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Property_ > chartId|Obtém a identificação do gráfico que é adicionado à planilha.|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Property_ > type|Obtém o tipo de evento.|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Property_ > worksheetId|Obtém a identificação da planilha em que o gráfico é adicionado.|1.8|
|[chartAddedEventArgs](/javascript/api/excel/excel.chartaddedeventargs)|_Relationship_ > source|Obtém a origem do evento.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > isBetweenCategories|Indica se o eixo dos valores cruza o eixo das categorias entre as categorias.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > multiLevel|Indica se um eixo tem vários níveis ou não.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > numberFormat|Representa o código de formatação para o rótulo de escala do eixo.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > offset|Representa a distância entre os níveis dos rótulos e a distância entre o primeiro nível e a linha de eixo. O valor deve ser um inteiro de 0 a 1.000.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > positionAt|Representa a posição do eixo especificado por onde o outro eixo irá cruzar. Você deve usar o método SetPositionAt(double) para definir essa propriedade. Somente leitura.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > textOrientation|Representa a orientação do texto do rótulo de escala do eixo. O valor deve ser um inteiro de -90 a 90 ou 180 para o texto orientado verticalmente.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > alignment|Representa o alinhamento para o rótulo de escala do eixo especificado.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > position|Representa a posição do eixo especificado por onde o outro eixo cruza.|1.8|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Método_ > [setPositionAt(value: double)](/javascript/api/excel/excel.chartaxis)|Define a posição do eixo especificado por onde o outro eixo irá cruzar.|1.8|
|[chartAxisFormat](/javascript/api/excel/excel.chartaxisformat)|_Relationship_ > fill|Representa formatação de preenchimento do gráfico. Somente leitura.|1.8|
|[chartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|_Method_ > [setFormula(formula: string)](/javascript/api/excel/excel.chartaxistitle)|Um valor da sequência de caracteres que representa a fórmula do título do eixo de gráfico usando a notação de estilo A1.|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Relationship_ > border|Representa o formato da borda, que inclui cor, estilo da linha e peso. Somente leitura.|1.8|
|[chartAxisTitleFormat](/javascript/api/excel/excel.chartaxistitleformat)|_Relationship_ > fill|Representa formatação de preenchimento do gráfico. Somente leitura.|1.8|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Method_ > [clear()](/javascript/api/excel/excel.chartborder)|Desmarque o formato da borda de um elemento do gráfico.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > autoText|Valor booleano que representa se rótulo de dados gera automaticamente o texto adequado com base no contexto.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > formula|Valor da sequência de caracteres que representa a fórmula do rótulo de dados do gráfico usando a notação de estilo A1.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > height|Retorna a altura, em pontos, do rótulo de dados do gráfico. Somente leitura. Nulo se o rótulo de dados do gráfico não estiver visível. Somente leitura.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > left|Representa a distância, em pontos, da borda esquerda do rótulo de dados do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > numberFormat|Valor da sequência de caracteres que representa o código de formatação do rótulo de dados.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > text|Sequência de caracteres que representa o texto do rótulo de dados em um gráfico.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > textOrientation|Representa a orientação do texto do rótulo de dados do gráfico. O valor deve ser um inteiro de -90 a 90 ou 180 para o texto orientado verticalmente.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > top|Representa a distância, em pontos, da borda superior do rótulo de dados do gráfico até o topo da área do gráfico. Nulo se o rótulo de dados do gráfico não estiver visível.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > width|Retorna a largura, em pontos, do rótulo de dados do gráfico. Somente leitura. Nulo se o rótulo de dados do gráfico não estiver visível. Somente leitura.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relationship_ > format|Representa o formato de rótulo de dados do gráfico. Somente leitura.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relationship_ > horizontalAlignment|Representa o alinhamento horizontal para o rótulo de dados do gráfico.|1.8|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Relationship_ > verticalAlignment|Representa o alinhamento vertical do rótulo de dados do gráfico.|1.8|
|[chartDataLabelFormat](/javascript/api/excel/excel.chartdatalabelformat)|_Relationship_ > border|Representa o formato da borda, que inclui cor, estilo da linha e peso. Somente leitura.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Property_ > autoText|Indica se os rótulos de dados geram automaticamente o texto apropriado com base no contexto.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Property_ > numberFormat|Representa o código de formato para os rótulos de dados.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Property_ > textOrientation|Representa a orientação do texto dos rótulos de dados. O valor deve ser um inteiro de -90 a 90 ou de 0 a 180 para o texto orientado verticalmente.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Relationship_ > horizontalAlignment|Representa o alinhamento horizontal para o rótulo de dados do gráfico.|1.8|
|[chartDataLabels](/javascript/api/excel/excel.chartdatalabels)|_Relationship_ > verticalAlignment|Representa o alinhamento vertical do rótulo de dados do gráfico.|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Property_ > chartId|Obtém a id do gráfico que for desativado.|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Property_ > type|Obtém o tipo de evento.|1.8|
|[chartDeactivatedEventArgs](/javascript/api/excel/excel.chartdeactivatedeventargs)|_Property_ > worksheetId|Obtém a identificação da planilha em que o gráfico foi desativado.|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Property_ > chartId|Obtém a id do gráfico que será excluído da planilha.|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Property_ > type|Obtém o tipo de evento.|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Property_ > worksheetId|Obtém a id da planilha em que o gráfico foi excluído.|1.8|
|[chartDeletedEventArgs](/javascript/api/excel/excel.chartdeletedeventargs)|_Relationship_ > source|Obtém a origem do evento.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Property_ > height|Representa a altura de legendEntry na legenda do gráfico. Somente leitura.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Property_ > index|Representa o índice de legendEntry na legenda do gráfico. Somente leitura.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Property_ > left|Representa a parte esquerda de legendEntry de um gráfico. Somente leitura.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Property_ > top|Representa a parte superior de legendEntry de um gráfico. Somente leitura.|1.8|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Property_ > width|Representa a largura de legendEntry na legenda do gráfico. Somente leitura.|1.8|
|[chartLegendFormat](/javascript/api/excel/excel.chartlegendformat)|_Relationship_ > border|Representa o formato da borda, que inclui cor, estilo da linha e peso. Somente leitura.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > height|Representa o valor de altura de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > insideHeight|Representa o valor insideHeight de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > insideLeft|Representa o valor insideLeft de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > insideTop|Representa o valor insideTop de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > insideWidth|Representa o valor insideWidth de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > left|Representa o valor à esquerda de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > top|Representa o valor superior de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Property_ > width|Representa o valor de largura de plotArea.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Relationship_ > format|Representa a formatação de plotArea de um gráfico. Somente leitura.|1.8|
|[chartPlotArea](/javascript/api/excel/excel.chartplotarea)|_Relationship_ > position|Representa a posição de plotArea.|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Relationship_ > border|Representa os atributos de borda de plotArea de um gráfico. Somente leitura.|1.8|
|[chartPlotAreaFormat](/javascript/api/excel/excel.chartplotareaformat)|_Relationship_ > fill|Representa o formato de preenchimento de um objeto, que inclui informações sobre a formatação do plano de fundo. Somente leitura.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > explosion|Retorna ou define o valor de explosão para um gráfico de pizza ou um gráfico de rosca. Retorna 0 (zero) se não houver explosão (quando a ponta da fatia estiver no centro da pizza). ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > firstSliceAngle|Retorna ou define o ângulo da primeira fatia do gráfico de pizza ou do gráfico de rosca, em graus (no sentido horário a partir da vertical). Essa propriedade se aplica somente aos gráficos de pizza, de pizza 3D e de rosca. Pode ter um valor de 0 a 360. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > invertIfNegative|Verdadeiro se o Microsoft Excel inverter o padrão no item quando este corresponder a um número negativo. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > overlap|Especifica como as barras e colunas são posicionadas. Pode ser um valor entre -100 e 100. Essa propriedade se aplica somente aos gráficos de barras 2D e de colunas 2D. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > secondPlotSize|Retorna ou define o tamanho da seção secundária de uma pizza do gráfico de pizza ou de um barra do gráfico de pizza, como um percentual do tamanho da pizza principal. Pode ser um valor de 5 a 200. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > varyByCategories|Verdadeiro se o Microsoft Word atribuir uma cor ou padrão diferente para cada marcador de dados. O gráfico deve conter apenas uma série. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relationship_ > axisGroup|Retorna ou define o grupo da série especificada. ReadWrite.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relationship_ > dataLabels|Representa uma coleção de todos os dataLabels da série. Somente leitura.|1.8|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relationship_ > splitType|Retorna ou define a maneira como as duas seções de uma pizza do gráfico de pizza ou de uma barra do gráfico de pizza são divididas. ReadWrite.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > backwardPeriod|Representa o número de períodos em que a linha de tendência se estende para trás.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > forwardPeriod|Representa o número de períodos em que a linha de tendência se estende para frente.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > showEquation|Verdadeiro se a equação para a linha de tendência for exibida no gráfico.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > showRSquared|Verdadeiro se o R-quadrado da linha de tendência for exibido no gráfico.|1.8|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Relationship_ > label|Representa o rótulo de linha de tendência de um gráfico. Somente leitura.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > autoText|Valor booleano que representa se rótulo de linha de tendência gera automaticamente o texto adequado com base no contexto.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > formula|Valor da sequência de caracteres que representa a fórmula do rótulo de linha de tendência do gráfico usando a notação de estilo A1.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > height|Retorna a altura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível. Somente leitura.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > left|Representa a distância, em pontos, da borda esquerda do rótulo de linha de tendência do gráfico até a borda esquerda da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > numberFormat|Valor da sequência de caracteres que representa o código de formatação do rótulo de linha de tendência.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > text|Sequência de caracteres que representa o texto do rótulo de linha de tendência em um gráfico.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > textOrientation|Representa a orientação do texto do rótulo de linha de tendência do gráfico. O valor deve ser um inteiro de -90 a 90 ou 180 para o texto orientado verticalmente.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > top|Representa a distância, em pontos, da borda superior do rótulo de linha de tendência do gráfico até o topo da área do gráfico. Nulo se o rótulo de linha de tendência do gráfico não estiver visível.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Property_ > width|Retorna a largura, em pontos, do rótulo de linha de tendência do gráfico. Somente leitura. Nulo se o rótulo de linha de tendência do gráfico não estiver visível. Somente leitura.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relationship_ > format|Representa o formato de rótulo de linha de tendência do gráfico. Somente leitura.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relationship_ > horizontalAlignment|Representa o alinhamento horizontal para o rótulo de linha de tendência do gráfico.|1.8|
|[chartTrendlineLabel](/javascript/api/excel/excel.charttrendlinelabel)|_Relationship_ > verticalAlignment|Representa o alinhamento vertical do rótulo de linha de tendência do gráfico.|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relationship_ > border|Representa o formato da borda, que inclui cor, estilo da linha e peso. Somente leitura.|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relationship_ > fill|Representa o formato de preenchimento do rótulo da linha de tendência de gráfico atual. Somente leitura.|1.8|
|[chartTrendlineLabelFormat](/javascript/api/excel/excel.charttrendlinelabelformat)|_Relationship_ > font|Representa os atributos de fonte (nome, tamanho, cor etc.) para um rótulo de linha de tendência do gráfico. Somente leitura.|1.8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_Property_ > fakeFileId|Transmite dados adicionais ao lado do cliente, por exemplo, worksheetId para TableSelectionChangedEvent.|1.8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_Property_ > fileBase64|Transmite dados adicionais ao lado do cliente, por exemplo, worksheetId para TableSelectionChangedEvent.|1.8|
|[createWorkbookPostProcessAction](/javascript/api/excel/excel.createworkbookpostprocessaction)|_Relationship_ > actionType|Transmite dados adicionais ao lado do cliente, por exemplo, worksheetId para TableSelectionChangedEvent.|1.8|
|[customDataValidation](/javascript/api/excel/excel.customdatavalidation)|_Property_ > formula| Uma fórmula de validação de dados personalizados. Isso cria regras especiais de entrada, como impedir duplicatas ou limitar o total em um intervalo de células.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Property_ > id|ID de DataPivotHierarchy. Somente leitura.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Property_ > name|Nome de DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Property_ > numberFormat|Formato de número de DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Property_ > position|Posição de DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relationship_ > field|Retorna os Campos Dinâmicos associados a DataPivotHierarchy. Somente leitura.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relationship_ > showAs|Determina se os dados devem ser exibidos como um cálculo de resumo específico ou não.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Relationship_ > summarizeBy|Determina se deve mostrar todos os itens de DataPivotHierarchy.|1.8|
|[dataPivotHierarchy](/javascript/api/excel/excel.datapivothierarchy)|_Method_ > [setToDefault()](/javascript/api/excel/excel.datapivothierarchy#settodefault)|Redefine o DataPivotHierarchy para seus valores padrão.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Property_ > items|Uma coleção de objetos dataPivotHierarchy. Somente leitura.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Method_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Adiciona o PivotHierarchy ao eixo atual.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Method_ > [getCount()](/javascript/api/excel/excel.datapivothierarchycollection)|Obtém o número de hierarquias originais na coleção.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Method_ > [getItem(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Obtém um DataPivotHierarchy por seu nome ou id.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Method_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.datapivothierarchycollection)|Obtém um DataPivotHierarchy pelo nome. Se o DataPivotHierarchy não existir, um objeto nulo será retornado.|1.8|
|[dataPivotHierarchyCollection](/javascript/api/excel/excel.datapivothierarchycollection)|_Method_ > [remove(DataPivotHierarchy: DataPivotHierarchy)](/javascript/api/excel/excel.datapivothierarchycollection)|Remove o PivotHierarchy do eixo atual.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Property_ > ignoreBlanks|Ignorar espaços em branco: nenhuma validação de dados será realizada em células vazias, o padrão é verdadeiro.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Property_ > valid|Representa se todos os valores de célula são válidos de acordo com as regras de validação de dados. Somente leitura.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relationship_ > errorAlert|Alerta de erro quando o usuário insere dados inválidos.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relationship_ > prompt|Avisa quando os usuários selecionam uma célula.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relationship_ > rule|Regra de validação de dados que contém os tipos diferentes de critérios de validação de dados.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Relationship_ > type|Para obter detalhes sobre o tipo de validação de dados, consulte [Excel.DataValidationType](/javascript/api/excel/excel.datavalidationtype). Somente leitura.|1.8|
|[dataValidation](/javascript/api/excel/excel.datavalidation)|_Method_ > [clear()](/javascript/api/excel/excel.datavalidation)|Desmarca a validação de dados do intervalo atual.|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Property_ > message|Representa a mensagem de alerta de erro.|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Property_ > showAlert|Determina se uma caixa de diálogo de alerta de erro deve ser mostrada ou não quando um usuário insere dados inválidos. O padrão é verdadeiro.|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Property_ > title|Representa o título da caixa de diálogo de alerta de erro.|1.8|
|[dataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)|_Relationship_ > style|Representa o tipo de alerta da validação de dados. Consulte [Excel.DataValidationAlertStyle](/javascript/api/excel/excel.datavalidationalertstyle) para obter detalhes.|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Property_ > message|Representa a mensagem de solicitação.|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Property_ > showPrompt|Determina se a solicitação deve ser mostrada ou não quando o usuário seleciona uma célula com validação de dados.|1.8|
|[dataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)|_Property_ > title|Representa o título para a solicitação.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > custom|Critérios de validação de dados personalizados.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > date|Critérios de validação de dados de data.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > decimal|Critérios de validação de dados decimais.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > list|Critérios de validação de dados de lista.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > textLength|Critérios de validação de dados TextLength.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > time|Critérios de validação de dados de hora.|1.8|
|[dataValidationRule](/javascript/api/excel/excel.datavalidationrule)|_Relationship_ > wholeNumber|Critérios de validação de dados WholeNumber.|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Property_ > formula1|Obtém ou define a Formula1, ou seja, o valor mínimo ou o valor dependente do operador.|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Property_ > formula2|Obtém ou define a Formula2, ou seja, o valor máximo ou o valor dependente do operador.|1.8|
|[dateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation)|_Relationship_ > operator|O operador a ser usado para validar os dados.|1.8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_Property_ > isEnableEvents{|Transmite dados adicionais ao lado do cliente, por exemplo, worksheetId para TableSelectionChangedEvent.|1.8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_Relationship_ > actionType|Transmite dados adicionais ao lado do cliente, por exemplo, worksheetId para TableSelectionChangedEvent.|1.8|
|[enableEventsPostProcessAction](/javascript/api/excel/excel.enableeventspostprocessaction)|_Relationship_ > controlId|Transmite dados adicionais ao lado do cliente, por exemplo, worksheetId para TableSelectionChangedEvent.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Property_ > enableMultipleFilterItems|Determina se vários itens de filtro devem ser permitidos.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Property_ > id|ID de FilterPivotHierarchy. Somente leitura.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Property_ > name|Nome de FilterPivotHierarchy.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Property_ > position|Posição de FilterPivotHierarchy.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Relationship_ > fields|Retorna os Campos Dinâmicos associados a FilterPivotHierarchy. Somente leitura.|1.8|
|[filterPivotHierarchy](/javascript/api/excel/excel.filterpivothierarchy)|_Method_ > [setToDefault()](/javascript/api/excel/excel.filterpivothierarchy)|Redefine o FilterPivotHierarchy para seus valores padrão.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Property_ > items|Uma coleção de objetos filterPivotHierarchy. Somente leitura.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Method_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Adiciona o PivotHierarchy ao eixo atual. Se a hierarquia estiver presente em outro lugar na linha, coluna ou eixo filtro, ela será removida desse local.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Method_ > [getCount()](/javascript/api/excel/excel.filterpivothierarchycollection)|Obtém o número de hierarquias originais na coleção.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Method_ > [getItem(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Obtém um FilterPivotHierarchy por seu nome ou id.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Method_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.filterpivothierarchycollection)|Obtém um FilterPivotHierarchy pelo nome. Se o FilterPivotHierarchy não existir, um objeto nulo será retornado.|1.8|
|[filterPivotHierarchyCollection](/javascript/api/excel/excel.filterpivothierarchycollection)|_Method_ > [remove(filterPivotHierarchy: FilterPivotHierarchy)](/javascript/api/excel/excel.filterpivothierarchycollection)|Remove o PivotHierarchy do eixo atual.|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Property_ > inCellDropDown|Exibe a lista no menu suspenso da célula ou não. O padrão é verdadeiro.|1.8|
|[listDataValidation](/javascript/api/excel/excel.listdatavalidation)|_Property_ > source|Origem da lista para validação de dados|1.8|
|[openWorkbookPostProcessAction](/javascript/api/excel/excel.openworkbookpostprocessaction)|_Property_ > fakeFileId|Transmite dados adicionais ao lado do cliente, por exemplo, worksheetId para TableSelectionChangedEvent.|1.8|
|[openWorkbookPostProcessAction](/javascript/api/excel/excel.openworkbookpostprocessaction)|_Relationship_ > actionType|Transmite dados adicionais ao lado do cliente, por exemplo, worksheetId para TableSelectionChangedEvent.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Property_ > id|ID do Campo Dinâmico. Somente leitura.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Property_ > name|Nome do Campo Dinâmico|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Property_ > showAllItems|Determina se todos os itens do Campo Dinâmico devem ser mostrados.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Relationship_ > items|Retorna os Campos Dinâmicos associado ao Campo Dinâmico. Somente leitura.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Relationship_ > subtotals|Subtotais do Campo Dinâmico.|1.8|
|[pivotField](/javascript/api/excel/excel.pivotfield)|_Method_ > [sortByLabels(sortby: SortBy)](/javascript/api/excel/excel.pivotfield)|Classifica o Campo Dinâmico. Se um DataPivotHierarchy for especificado, então a classificação será aplicada com base nele; caso contrário, a classificação será baseada no Campo Dinâmico em si.|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Property_ > items|Uma coleção de objetos pivotField. Somente leitura.|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Method_ > [getCount()](/javascript/api/excel/excel.pivotfieldcollection)|Obtém o número de hierarquias originais na coleção.|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Method_ > [getItem(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Obtém um PivotHierarchy por seu nome ou id.|1.8|
|[pivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection)|_Method_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotfieldcollection)|Obtém um PivotHierarchy pelo nome. Se o PivotHierarchy não existir, um objeto nulo será retornado.|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Property_ > id|ID de PivotHierarchy. Somente leitura.|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Property_ > name|Nome de PivotHierarchy.|1.8|
|[pivotHierarchy](/javascript/api/excel/excel.pivothierarchy)|_Relationship_ > fields|Retorna os Campos Dinâmicos associados a PivotHierarchy. Somente leitura.|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Property_ > items|Uma coleção de objetos pivotHierarchy. Somente leitura.|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Method_ > [getCount()](/javascript/api/excel/excel.pivothierarchycollection)|Obtém o número de hierarquias originais na coleção.|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Method_ > [getItem(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Obtém um PivotHierarchy por seu nome ou id.|1.8|
|[pivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection)|_Method_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivothierarchycollection)|Obtém um PivotHierarchy pelo nome. Se o PivotHierarchy não existir, um objeto nulo será retornado.|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Property_ > id|ID de PivotItem. Somente leitura.|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Property_ > isExpanded|Determina se o item é expandido para mostrar itens filhos ou se ele foi recolhido e os itens filhos estão ocultos.|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Property_ > name|Nome de PivotItem.|1.8|
|[pivotItem](/javascript/api/excel/excel.pivotitem)|_Property_ > visible|Determina se o PivotItem está visível ou não.|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Property_ > items|Uma coleção de objetos pivotItem. Somente leitura.|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Method_ > [getCount()](/javascript/api/excel/excel.pivotitemcollection)|Obtém o número de hierarquias originais na coleção.|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Method_ > [getItem(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Obtém um PivotHierarchy por seu nome ou id.|1.8|
|[pivotItemCollection](/javascript/api/excel/excel.pivotitemcollection)|_Method_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.pivotitemcollection)|Obtém um PivotHierarchy pelo nome. Se o PivotHierarchy não existir, um objeto nulo será retornado.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Property_ > showColumnGrandTotals|Verdadeiro se o relatório de tabela dinâmica mostrar os totais gerais para as colunas.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Property_ > showRowGrandTotals|Verdadeiro se o relatório de tabela dinâmica mostrar os totais gerais para as linhas.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Property_ > subtotalLocation|Esta propriedade indica o SubtotalLocationType de todos os campos na tabela dinâmica. Se os campos tiverem diferentes estados, isso será nulo. Os valores possíveis são: AtTop, AtBottom.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Relationship_ > layoutType|Esta propriedade indica o PivotLayoutType de todos os campos na tabela dinâmica. Se os campos tiverem diferentes estados, isso será nulo.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Method_ > [getColumnLabelRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo onde residem os rótulos de coluna da tabela dinâmica.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Method_ > [getDataBodyRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo onde residem os valores de dados da tabela dinâmica.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout.md)|_Method_ > [getFilterAxisRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo da área de filtragem da tabela dinâmica.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Method_ > [getRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo sobre o qual a tabela dinâmica existe, excluindo a área de filtragem.|1.8|
|[pivotLayout](/javascript/api/excel/excel.pivotlayout)|_Method_ > [getRowLabelRange()](/javascript/api/excel/excel.pivotlayout)|Retorna o intervalo onde residem os rótulos de linha da tabela dinâmica.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > columnHierarchies|Hierarquias originais de coluna da tabela dinâmica. Somente leitura.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > dataHierarchies|Hierarquias originais de dados da tabela dinâmica. Somente leitura.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > filterHierarchies|Hierarquias originais de filtragem da tabela dinâmica. Somente leitura.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > hierarchies|Hierarquias originais da tabela dinâmica. Somente leitura.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > layout|PivotLayout descrevendo o layout e a estrutura visual da tabela dinâmica. Somente leitura.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > rowHierarchies|Hierarquias originais de linha da tabela dinâmica. Somente leitura.|1.8|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Method_ > [delete()](/javascript/api/excel/excel.pivottable)|Exclui a tabela dinâmica.|1.8|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Method_ > [add(name: string, source: object, destination: object)](/javascript/api/excel/excel.pivottablecollection)|Adiciona uma tabela dinâmica com base nos dados de origem especificados e os insere na célula superior esquerda do intervalo de destino.|1.8|
|[range](/javascript/api/excel/excel.range)|_Relationship_ > dataValidation|Retorna um objeto de validação de dados. Somente leitura.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Property_ > id|ID de RowColumnPivotHierarchy. Somente leitura.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Property_ > name|Nome de RowColumnPivotHierarchy.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Property_ > position|Posição de RowColumnPivotHierarchy.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Relationship_ > fields|Retorna os Campos Dinâmicos associados a RowColumnPivotHierarchy. Somente leitura.|1.8|
|[rowColumnPivotHierarchy](/javascript/api/excel/excel.rowcolumnpivothierarchy)|_Method_ > [setToDefault()](/javascript/api/excel/excel.rowcolumnpivothierarchy)|Redefine o RowColumnPivotHierarchy para seus valores padrão.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Property_ > items|Uma coleção de objetos rowColumnPivotHierarchy. Somente leitura.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Method_ > [add(pivotHierarchy: PivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Adiciona o PivotHierarchy ao eixo atual. Se a hierarquia estiver presente em outro lugar na linha, coluna,|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Method_ > [getCount()](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Obtém o número de hierarquias originais na coleção.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Method_ > [getItem(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Obtém um RowColumnPivotHierarchy por seu nome ou id.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Method_ > [getItemOrNullObject(name: string)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Obtém um RowColumnPivotHierarchy pelo nome. Se o RowColumnPivotHierarchy não existir, um objeto nulo será retornado.|1.8|
|[rowColumnPivotHierarchyCollection](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|_Method_ > [remove(rowColumnPivotHierarchy: RowColumnPivotHierarchy)](/javascript/api/excel/excel.rowcolumnpivothierarchycollection)|Remove o PivotHierarchy do eixo atual.|1.8|
|[runtime](/javascript/api/excel/excel.runtime)|_Property_ > enableEvents|Alternar os eventos de JavaScript no painel de tarefas atual ou no suplemento de conteúdos.|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relationship_ > baseField|O Campo Dinâmico base para fundamentar o cálculo de ShowAs, se aplicável com base no tipo ShowAsCalculation; caso contrário, é nulo.|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relationship_ > baseItem|O Item base para fundamentar o cálculo de ShowAs, se aplicável com base no tipo ShowAsCalculation; caso contrário, é nulo.|1.8|
|[showAsRule](/javascript/api/excel/excel.showasrule)|_Relationship_ > calculation|Cálculo de ShowAs a ser usado para o Campo Dinâmico de dados.|1.8|
|[style](/javascript/api/excel/excel.style)|_Property_ > autoIndent|Indica se o texto será recuado automaticamente quando o alinhamento do texto em uma célula estiver definido como distribuição igual.|1.8|
|[style](/javascript/api/excel/excel.style)|_Property_ > textOrientation|A orientação do texto para o estilo.|1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > automatic|Se Automatic estiver definido como verdadeiro, todos os outros valores serão ignorados ao definir Subtotals.|1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > average| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > count| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Propriedade_ > countNumbers| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > max| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > min| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > product| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > standardDeviation| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > standardDeviationP| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > sum| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > variance| |1.8|
|[subtotals](/javascript/api/excel/excel.subtotals)|_Property_ > varianceP| |1.8|
|[table](/javascript/api/excel/excel.table)|_Property_ > legacyId|Retorna uma id numérica. Somente leitura.|1.8|
|[workbook](/javascript/api/excel/excel.workbook)|_Property_ > readOnly|Verdadeiro se a pasta de trabalho for aberta no modo Somente leitura. Somente leitura.|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Property_ > id|Retorna um valor que identifica exclusivamente o objeto WorkbookCreated. Somente leitura.|1.8|
|[workbookCreated](/javascript/api/excel/excel.workbookcreated)|_Method_ > [open()](/javascript/api/excel/excel.workbookcreated)|Abra a pasta de trabalho.|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > showGridlines|Obtém ou define o sinalizador de linhas de grade da planilha.|1.8|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > showHeadings|Obtém ou define o sinalizador de títulos da planilha.|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Property_ > type|Obtém o tipo de evento.|1.8|
|[worksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|_Property_ > worksheetId|Obtém a identificação da planilha que é calculada.|1.8|

## <a name="whats-new-in-excel-javascript-api-17"></a>Novidades na API JavaScript do Excel 1.7

Os recursos do conjunto de requisitos da API JavaScript do Excel 1.7 incluem APIs para gráficos, eventos, planilhas, intervalos, propriedades de documento, itens denominados, opções de proteção e estilos.

### <a name="customize-charts"></a>Personalizar gráficos

Com as novas APIs de gráfico, você pode criar tipos de gráfico adicionais, adicionar uma série de dados a um gráfico, definir o título do gráfico, adicionar um título do eixo, adicionar uma unidade de exibição, adicionar uma linha de tendência com média móvel, alterar uma linha de tendência para linear e muito mais. Eis alguns exemplos:

* Eixo de gráfico - obter, definir, formatar e remover a unidade de eixo, rótulo e título em um gráfico.
* Série de gráfico - adicionar, definir e excluir uma série em um gráfico.  Alterar os marcadores de série, as ordens de plotagem e o dimensionamento.
* Linhas de tendência do gráfico - adicionar, obter e formatar as linhas de tendência em um gráfico.
* Legenda do gráfico - formatar a fonte da legenda em um gráfico.
* Ponto do gráfico - definir a cor do ponto do gráfico.
* Subsequência de caracteres do título do gráfico - obter e definir a subsequência de caracteres do título de um gráfico.
* Tipo de gráfico - opção para criar mais tipos de gráfico.

### <a name="events"></a>Eventos

As APIs de eventos do Excel fornecem uma variedade de manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando um evento específico ocorre. Você pode projetar essa função para executar quaisquer ações que seu cenário exigir. Para obter uma lista de eventos que estão atualmente disponíveis, consulte [Trabalhar com eventos usando a API JavaScript do  Excel](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-events).

### <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Personalizar a aparência das planilhas e dos intervalos

Ao usar as novas APIs, você pode personalizar a aparência das planilhas de várias maneiras:

* Congele painéis para manter colunas ou linhas específicas visíveis quando você rolar pela planilha. Por exemplo, se a primeira linha em sua planilha contiver cabeçalhos, você pode congelar essa linha para que os cabeçalhos de coluna permaneçam visíveis enquanto rola para baixo pela planilha.
* Modificar a cor da guia da planilha.
* Adicionar títulos da planilha.


Você pode personalizar a aparência dos intervalos de várias maneiras:

* Defina o estilo de célula para um intervalo para garantir que todas as células no intervalo possuam uma formatação consistente. Um estilo de célula é um conjunto definido de características de formatação, como fontes e tamanhos de fonte, formatos de número, bordas da célula e sombreamento da célula. Use qualquer um dos estilos de célula internas do Excel ou crie seu próprio estilo de célula personalizado.
* Defina a orientação do texto para um intervalo.
* Adicione ou modifique um hiperlink em um intervalo que se conecta a outro local na pasta de trabalho ou a um local externo.

### <a name="manage-document-properties"></a>Gerenciar propriedades do documento

Ao usar as APIs de propriedades do documento, você pode acessar as propriedades do documento internas e também criar e gerenciar propriedades do documento personalizadas para armazenar o estado da pasta de trabalho e a unidade de fluxo de trabalho e lógica de negócios.

### <a name="copy-worksheets"></a>Copiar planilhas

Ao usar as APIs de cópia de planilha, você pode copiar os dados e o formato de uma planilha para uma nova planilha dentro da mesma pasta de trabalho e reduzir a quantidade de transferência de dados necessária.

### <a name="handle-ranges-with-ease"></a>Manipular intervalos com facilidade

Ao usar as diversas APIs de intervalo, você pode fazer coisas como obter a região ao redor, obter um intervalo redimensionado e muito mais. Essas APIs deve tornar tarefas como manipulação e endereçamento de intervalo muito mais eficientes.

Além disso:

* Opções de proteção de pasta de trabalho e planilha - use essas APIs para proteger os dados em uma planilha e estrutura de pasta de trabalho.
* Atualizar um item nomeado - use essa API para atualizar um item nomeado.
* Obter a célula ativa - use essa API para obter a célula ativa de uma pasta de trabalho.

|Object| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > chartType|Representa o tipo de gráfico. Os valores possíveis são: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie etc..|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > id|A id exclusiva do gráfico. Somente leitura.|1.7|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > showAllFieldButtons|Representa se todos os botões de campos em um gráfico dinâmico devem ser exibidos.|1.7|
|[chartAreaFormat](/javascript/api/excel/excel.chartareaformat)|_Relationship_ > border|Representa o formato da borda da área do gráfico, que inclui cor, estilo da linha e peso. Somente leitura.|1.7|
|[chartAxes](/javascript/api/excel/excel.chartaxes)|_Method_ > getItem(type: string, group: string)|Retorna o eixo específico identificado pelo tipo e grupo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > axisBetweenCategories|Indica se o eixo dos valores cruza o eixo das categorias entre as categorias.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > axisGroup|Representa o grupo para o eixo especificado. Somente leitura. Os valores possíveis são: Primary, Secondary.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > categoryType|Retorna ou define o tipo de eixo das categorias. Os valores possíveis são: Automatic, TextAxis, DateAxis.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > crosses|Representa o eixo especificado onde o outro eixo cruza. Os valores possíveis são: Automatic, Maximum, Minimum, Custom.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > crossesAt|Representa o eixo especificado por onde o outro eixo cruza. Somente leitura. A definição para essa propriedade deve usar o método SetCrossesAt(double). Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > customDisplayUnit|Representa o valor da unidade de exibição do eixo personalizado. Somente leitura. Para definir essa propriedade, use o método SetCustomDisplayUnit(double). Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > displayUnit|Representa a unidade de exibição do eixo. Os valores possíveis são: None, Hundreds, Thousands, TenThousands, HundredThousands, Millions, TenMillions, HundredMillions, Billions, Trillions, Custom.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > height|Representa a altura, em pontos, do eixo do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > left|Representa a distância, em pontos, da borda esquerda do eixo à esquerda da área do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > logBase|Representa a base do logaritmo ao usar escalas logarítmicas.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > reversePlotOrder|Indica se o Microsoft Excel plota os pontos de dados dos últimos aos primeiros.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > scaleType|Representa o tipo de escala do eixo dos valores. Os valores possíveis são: Linear, Logarithmic.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > showDisplayUnitLabel|Representa se o rótulo da unidade de exibição do eixo está visível.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Propriedade_ > tickLabelSpacing|Representa o número de categorias ou séries entre os rótulos de marca de escala. Pode ser um valor de 1 a 31.999 ou uma sequência de caracteres vazia para configuração automática. O valor retornado sempre será um número.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > tickMarkSpacing|Representa o número de categorias ou séries entre marcas de escala.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > top|Representa a distância, em pontos, da borda superior do eixo ao topo da área do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > type|Representa o tipo de eixo. Somente leitura. Os valores possíveis são: Invalid, Category, Value, Series.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > visible|Um valor booleano representa a visibilidade do eixo.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Property_ > width|Representa a largura, em pontos, do eixo do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relacionamento_ > baseTimeUnit|Retorna ou define a unidade base para o eixo das categorias especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > majorTickMark|Representa o tipo de marca de escala primária para o eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > majorTimeUnitScale|Retorna ou define o valor de escala de unidades primária para o eixo das categorias quando a propriedade CategoryType estiver definida como TimeScale.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > minorTickMark|Representa o tipo de marca de escala secundária para o eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > minorTimeUnitScale|Retorna ou define o valor de escala de unidades secundária para o eixo das categorias quando a propriedade CategoryType estiver definida como TimeScale.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Relationship_ > tickLabelPosition|Especifica a posição dos rótulos de marcas de escala no eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Method_ > setCategoryNames(sourceData: Range)|Define todos os nomes de categoria para o eixo especificado.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Method_ > setCrossesAt(value: double)|Define o eixo especificado por onde o outro eixo irá cruzar.|1.7|
|[chartAxis](/javascript/api/excel/excel.chartaxis)|_Method_ > setCustomDisplayUnit(value: double)|Define a unidade de exibição do eixo para um valor personalizado.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Property_ > color|Código de cor HTML que representa a cor das bordas do gráfico.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Property_ > weight|Representa a espessura da borda, em pontos.|1.7|
|[chartBorder](/javascript/api/excel/excel.chartborder)|_Relationship_ > lineStyle|Representa o estilo da linha da borda.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > position|Valor DataLabelPosition que representa a posição do rótulo de dados. Os valores possíveis são: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > separator|Sequência de caracteres que representa o separador usado para o rótulo de dados em um gráfico.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > showBubbleSize|Valor booleano que representa se o tamanho da bolha do rótulo de dados está visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > showCategoryName|Valor booleano que representa se o nome da categoria do rótulo de dados está visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > showLegendKey|Valor booleano que representa se o código de legenda do rótulo de dados está visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > showPercentage|Valor booleano que representa se o percentual do rótulo de dados está visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > showSeriesName|Valor booleano que representa se o nome da série do rótulo de dados está visível ou não.|1.7|
|[chartDataLabel](/javascript/api/excel/excel.chartdatalabel)|_Property_ > showValue|Valor booleano que determina se o valor do rótulo de dados está visível ou não.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Property_ > height|Representa a altura da legenda do gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Property_ > left|Representa esquerda da legenda de um gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Property_ > showShadow|Representa se a legenda tem sombra no gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Property_ > top|Representa o topo da legenda de um gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Property_ > width|Representa a largura da legenda do gráfico.|1.7|
|[chartLegend](/javascript/api/excel/excel.chartlegend)|_Relationship_ > legendEntries|Representa uma coleção de legendEntries na legenda. Somente leitura.|1.7|
|[chartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|_Property_ > visible|Representa a parte visível de uma entrada de legenda do gráfico.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Property_ > items|Uma coleção de objetos chartLegendEntry. Somente leitura.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Method_ > getCount()|Retorna o número de legendEntry na coleção.|1.7|
|[chartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|_Method_ > getItemAt(index: number)|Retorna um legendEntry no índice fornecido.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Property_ > hasDataLabel|Representa se um ponto de dados tem datalabel. Não se aplica a gráficos de superfície.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Property_ > markerBackgroundColor|Representação do código de cores HTML da cor do plano de fundo do marcador do ponto de dados. Exemplo: #FF0000 representa vermelho.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Property_ > markerForegroundColor|Representação do código de cores HTML da cor de primeiro plano do ponto de dados. Exemplo: #FF0000 representa vermelho.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Property_ > markerSize|Representa o tamanho do marcador do ponto de dados.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Property_ > markerStyle|Representa o estilo de marcador de um ponto de dados do gráfico. Os valores possíveis são: Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|1.7|
|[chartPoint](/javascript/api/excel/excel.chartpoint)|_Relationship_ > dataLabel|Retorna o rótulo de dados de um ponto do gráfico. Somente leitura.|1.7|
|[chartPointFormat](/javascript/api/excel/excel.chartpointformat)|_Relationship_ > border|Representa o formato de borda de um ponto de dados do gráfico, que inclui informações de cor, estilo e espessura. Somente leitura.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > chartType|Representa o tipo de gráfico de uma série. Os valores possíveis são: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie etc..|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > doughnutHoleSize|Representa o tamanho do orifício de rosca uma série de gráfico.  Válido somente em gráficos de rosca e doughnutExploded.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > filtered|Valor booleano que representa se a série está filtrada ou não. Não se aplica a gráficos de superfície.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > gapWidth|Representa a largura do intervalo de uma série de gráfico.  Válido apenas em gráficos de barras e de colunas, assim como|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > hasDataLabels|Valor booleano que representa se a série tem rótulos de dados ou não.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > markerBackgroundColor|Representa a cor do plano de fundo de marcadores de uma série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > markerForegroundColor|Representa a cor de primeiro plano de marcadores de uma série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > markerSize|Representa o tamanho do marcador de uma série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > markerStyle|Representa o estilo de marcador de uma série de gráfico. Os valores possíveis são: Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > plotOrder|Representa a ordem de plotagem de uma série de gráfico dentro do grupo de gráficos.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > showShadow|Valor booleano que representa se a série tem sombra ou não.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Property_ > smooth|Valor booleano que representa se a série é suave ou não. Somente para gráficos de linhas e de dispersão.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relationship_ > dataLabels|Representa uma coleção de todos os dataLabels da série. Somente leitura.|ApiSet.InProgressFeatures.ChartingAPI|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Relationship_ > trendlines|Representa uma coleção de linhas de tendência na série. Somente leitura.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Method_ > delete()|Exclui a série de gráfico.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Method_ > setBubbleSizes(sourceData: Range)|Define os tamanhos de bolha para uma série de gráfico. Funciona somente para gráficos de bolhas.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Método_ > setValues(sourceData: Range)|Define os valores de uma série de gráfico. Para o gráfico de dispersão, isso significa os valores do eixo Y.|1.7|
|[chartSeries](/javascript/api/excel/excel.chartseries)|_Method_ > setXAxisValues(sourceData: Range)|Define os valores do eixo X de uma série de gráfico. Funciona somente para gráficos de dispersão.|1.7|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Method_ > add(name: string, index: number)|Adiciona uma nova série à coleção.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > height|Retorna a altura, em pontos, do título do gráfico. Somente leitura. Nulo se o título do gráfico não estiver visível. Somente leitura.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > horizontalAlignment|Representa o alinhamento horizontal para o título do gráfico. Os valores possíveis são: Center, Left, Justify, Distributed, Right.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > left|Representa a distância, em pontos, da borda esquerda do título do gráfico até a borda esquerda da área do gráfico. Nulo se o título do gráfico não estiver visível.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > position|Representa a posição do título do gráfico. Os valores possíveis são: Top, Automatic, Bottom, Right, Left.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > showShadow|Representa um valor booleano que determina se o título do gráfico tem uma sombra.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > textOrientation|Representa a orientação do texto do título do gráfico. O valor deve ser um inteiro de -90 a 90 ou 180 para o texto orientado verticalmente.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > top|Representa a distância, em pontos, da borda superior do título do gráfico ao topo da área do gráfico. Nulo se o título do gráfico não estiver visível.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > verticalAlignment|Representa o alinhamento vertical do título do gráfico. Os valores possíveis são: Center, Bottom, Top, Justify, Distributed.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Property_ > width|Retorna a altura, em pontos, do título do gráfico. Somente leitura. Nulo se o título do gráfico não estiver visível. Somente leitura.|1.7|
|[chartTitle](/javascript/api/excel/excel.charttitle)|_Method_ > setFormula(formula: string)|Define um valor da sequência de caracteres que representa a fórmula do título do gráfico usando a notação de estilo A1.|1.7|
|[chartTitleFormat](/javascript/api/excel/excel.charttitleformat)|_Relationship_ > border|Representa o formato da borda do título do gráfico, que inclui cor, estilo da linha e peso. Somente leitura.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > backward|Representa o número de períodos em que a linha de tendência se estende para trás.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > displayEquation|Verdadeiro se a equação para a linha de tendência for exibida no gráfico.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > displayRSquared|Verdadeiro se o R-quadrado da linha de tendência for exibido no gráfico.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > forward|Representa o número de períodos em que a linha de tendência se estende para frente.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > intercept|Representa o valor de interceptação da linha de tendência. Pode ser definido como um valor numérico ou uma sequência de caracteres vazia (para valores automáticos). O valor retornado sempre será um número.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > movingAveragePeriod|Representa o período de uma linha de tendência do gráfico, apenas para a linha de tendência com o tipo MovingAverage.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > name|Representa o nome da linha de tendência. Pode ser definido como um valor da sequência de caracteres ou pode ser definido como um valor nulo que representa os valores automáticos. O valor retornado sempre é uma sequência de caracteres|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > polynomialOrder|Representa a ordem de uma linha de tendência do gráfico, apenas para a linha de tendência com tipo Polynomial.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Property_ > type|Representa o tipo de uma linha de tendência do gráfico. Os valores possíveis são: Linear, Exponential, Logarithmic, MovingAverage, Polynomial, Power.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Relationship_ > format|Representa a formatação de uma linha de tendência do gráfico. Somente leitura.|1.7|
|[chartTrendline](/javascript/api/excel/excel.charttrendline)|_Method_ > delete()|Exclui o objeto trendline.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Property_ > items|Uma coleção de objetos chartTrendline. Somente leitura.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Method_ > add(type: string)|Adiciona uma nova linha de tendência à coleção de linha de tendência.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Method_ > getCount()|Retorna o número de linhas de tendência na coleção.|1.7|
|[chartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|_Method_ > getItem(index: number)|Obtém o objeto trendline pelo índice, que é a ordem de inserção na matriz de itens.|1.7|
|[chartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|_Relationship_ > line|Representa a formatação de linha do gráfico. Somente leitura.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Property_ > key|Obtém a chave da propriedade personalizada. Somente leitura. Somente leitura.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Property_ > type|Obtém o tipo de valor da propriedade personalizada. Somente leitura. Somente leitura. Os valores possíveis são: Number, Boolean, Date, String, Float.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Property_ > value|Obtém ou define o valor da propriedade personalizada.|1.7|
|[customProperty](/javascript/api/excel/excel.customproperty)|_Method_ > delete()|Exclui a propriedade personalizada.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Property_ > items|Uma coleção de objetos customProperty. Somente leitura.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Method_ > add(key: string, value: object)|Cria uma nova propriedade personalizada ou define uma existente.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Method_ > deleteAll()|Exclui todas as propriedades personalizadas nesta coleção.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Method_ > getCount()|Obtém a contagem das propriedades personalizadas.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Method_ > getItem(key: string)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. É gerado se a propriedade personalizada não existir.|1.7|
|[customPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|_Method_ > getItemOrNullObject(key: string)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Retorna um objeto nulo se a propriedade personalizada não existir.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Property_ > items|Uma coleção de objetos dataConnection. Somente leitura.|1.7|
|[dataConnectionCollection](/javascript/api/excel/excel.dataconnectioncollection)|_Method_ > refreshAll()|Atualiza todas as conexões de dados na coleção.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > author|Obtém ou define o autor da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > category|Obtém ou define a categoria da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > comments|Obtém ou define os comentários da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > company|Obtém ou define a empresa da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > keywords|Obtém ou define as palavras-chave da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > lastAuthor|Obtém o último autor da pasta de trabalho. Somente leitura. Somente leitura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > manager|Obtém ou define o gerente da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > revisionNumber|Obtém o número de revisão da pasta de trabalho. Somente leitura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > subject|Obtém ou define o assunto da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Property_ > title|Obtém ou define o título da pasta de trabalho.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relationship_ > creationDate|Obtém a data de criação da pasta de trabalho. Somente leitura. Somente leitura.|1.7|
|[documentProperties](/javascript/api/excel/excel.documentproperties)|_Relationship_ > custom|Obtém a coleção das propriedades personalizadas da pasta de trabalho. Somente leitura. Somente leitura.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Property_ > formula|Obtém ou define a fórmula do item nomeado.  A fórmula sempre começa com um sinal de '='.|1.7|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relationship_ > arrayValues|Retorna um objeto que contém os valores e tipos de item nomeado. Somente leitura.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Property_ > types|Representa os tipos de cada item na matriz de itens nomeados Somente leitura. Os valores possíveis são: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.7|
|[namedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|_Property_ > values|Representa os valores de cada item na matriz de itens nomeados. Somente leitura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Property_ > isEntireColumn|Representa se o intervalo atual é uma coluna inteira. Somente leitura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Property_ > isEntireRow|Representa se o intervalo atual é uma linha inteira. Somente leitura.|1.7|
|[range](/javascript/api/excel/excel.range)|_Property_ > numberFormatLocal|Representa o código de formato de número do Excel para o intervalo especificado como uma sequência de caracteres no idioma do usuário.|1.7|
|[range](/javascript/api/excel/excel.range)|_Property_ > style|Representa o estilo do intervalo atual. Isso retornará nulo ou uma sequência de caracteres.|1.7|
|[range](/javascript/api/excel/excel.range)|_Method_ > getAbsoluteResizedRange(numRows: number, numColumns: number)|Obtém um objeto Range com a mesma célula superior esquerda que o objeto Range atual, porém com os números especificados de linhas e colunas.|1.7|
|[range](/javascript/api/excel/excel.range)|_Method_ > getImage()|Renderiza o intervalo como uma imagem codificado em base64.|1.7|
|[range](/javascript/api/excel/excel.range)|_Method_ > getSurroundingRegion()|Retorna um objeto Range que representa a região ao redor da célula superior esquerda nesse intervalo. Uma região ao redor é um intervalo limitado por qualquer combinação de linhas e colunas em branco em relação a esse intervalo.|1.7|
|[range](/javascript/api/excel/excel.range)|_Method_ > showCard()|Exibe o cartão para uma célula ativa se ela tiver o conteúdo de valor sofisticado.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Property_ > textOrientation|Obtém ou define a orientação do texto de todas as células dentro do intervalo.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Property_ > useStandardHeight|Determina se a altura da linha do objeto Range equivale à altura padrão da planilha.|1.7|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Property_ > useStandardWidth|Determina se a largura da coluna do objeto Range equivale à largura padrão da planilha.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Property_ > address|Representa o destino da url para o hiperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Property_ > document..|Representa o documento... destino do hiperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Property_ > screenTip|Representa a sequência de caracteres exibida ao passar o mouse sobre o hiperlink.|1.7|
|[rangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|_Property_ > textToDisplay|Representa a sequência de caracteres que é exibida na parte superior esquerda da maioria das células do intervalo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > addIndent|Indica se o texto será recuado automaticamente quando o alinhamento do texto em uma célula estiver definido como distribuição igual.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > autoIndent|Indica se o texto será recuado automaticamente quando o alinhamento do texto em uma célula estiver definido como distribuição igual.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > builtIn|Indica se o estilo for um estilo interno. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > formulaHidden|Indica se a fórmula ficará oculta quando a planilha for protegida.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > horizontalAlignment|Representa o alinhamento horizontal para o estilo. Os valores possíveis são: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includeAlignment|Indica se o estilo inclui as propriedades AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel e TextOrientation.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includeBorder|Indica se o estilo inclui as propriedades de borda Color, ColorIndex, LineStyle e Weight.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includeFont|Indica se o estilo incluir as propriedades de fonte Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript e Underline.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includeNumber|Indica se o estilo inclui a propriedade NumberFormat.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includePatterns|Indica se o estilo inclui as propriedades de interior Color, ColorIndex, InvertIfNegative, Pattern, PatternColor e PatternColorIndex.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > includeProtection|Indica se o estilo inclui as propriedades de proteção FormulaHidden e Locked.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > indentLevel|Um inteiro de 0 a 250 que indica o nível de recuo para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > locked|Indica se o objeto está bloqueado quando a planilha for protegida.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > name|O nome do estilo. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > numberFormat|O código de formato do formato de número para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > numberFormatLocal|O código de formatação localizado do formato de número para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > orientation|A orientação do texto para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > readingOrder|A ordem de leitura para o estilo. Os valores possíveis são: Context, LeftToRight, RightToLeft.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > shrinkToFit|Indica se o texto é automaticamente reduzido para se ajustar à largura da coluna disponível.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > textOrientation|A orientação do texto para o estilo.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > verticalAlignment|Representa o alinhamento vertical para o estilo. Os valores possíveis são: Top, Center, Bottom, Justify, Distributed.|1.7|
|[style](/javascript/api/excel/excel.style)|_Property_ > wrapText|Indica se o Microsoft Excel faz quebra automática do texto no objeto.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relationship_ > borders|Uma coleção Border de quatro objetos Border que representam o estilo das quatro bordas. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relationship_ > fill|O preenchimento do estilo. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Relationship_ > font|Um objeto Font que representa a fonte do estilo. Somente leitura.|1.7|
|[style](/javascript/api/excel/excel.style)|_Method_ > delete()|Exclui esse estilo.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Property_ > items|Uma coleção de objetos style. Somente leitura.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Method_ > add(name: string)]|Adiciona um novo estilo à coleção.|1.7|
|[styleCollection](/javascript/api/excel/excel.stylecollection)|_Method_ > getItem(key: string)|Obtém um estilo pelo nome.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Property_ > address|Obtém o endereço que representa a área alterada de uma tabela em uma planilha específica.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Property_ > changeType|Obtém o tipo de alteração que representa como o evento Changed é disparado. Os valores possíveis são: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Property_ > source|Obtém a origem do evento. Os valores possíveis são: Local, Remote.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Property_ > tableId|Obtém a id da tabela na qual os dados foram alterados.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Property_ > type|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|_Property_ > worksheetId|Obtém a id da planilha na qual os dados foram alterados.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Property_ > address|Obtém o endereço do intervalo que representa a área selecionada da tabela em uma planilha específica.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Property_ > isInsideTable|Indica se a seleção estiver dentro de uma tabela, o endereço será inútil se IsInsideTable for falso.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Property_ > tableId|Obtém a id da tabela na qual a seleção foi alterada.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Property_ > type|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[tableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|_Property_ > worksheetId|Obtém a id da planilha na qual a seleção foi alterada.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Property_ > name|Obtém o nome da pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > dataConnections|Atualiza todas as conexões de dados na pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > properties|Obtém as propriedades de pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > protection|Retorna o objeto de proteção de pasta de trabalho para uma pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > styles|Representa uma coleção de estilos associados à pasta de trabalho. Somente leitura.|1.7|
|[workbook](/javascript/api/excel/excel.workbook)|_Method_ > getActiveCell()|Obtém a célula atualmente ativa da pasta de trabalho.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Property_ > protected|Indica se a pasta de trabalho está protegida. Somente Leitura. Somente leitura.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Method_ > protect(password: string)|Protege uma pasta de trabalho. Falha se a pasta de trabalho foi protegida.|1.7|
|[workbookProtection](/javascript/api/excel/excel.workbookprotection)|_Method_ > unprotect(password: string)|Desprotege uma pasta de trabalho.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > gridlines|Obtém ou define o sinalizador de linhas de grade da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > headings|Obtém ou define o sinalizador de títulos da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > showHeadings|Obtém ou define o sinalizador de títulos da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > standardHeight|Retorna a altura padrão de todas as linhas da planilha, em pontos. Somente leitura.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > standardWidth|Retorna ou define a largura padrão de todas as colunas da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Property_ > tabColor|Obtém ou define a cor da guia da planilha.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relationship_ > freezePanes|Obtém um objeto que pode ser usado para manipular os painéis congelados na planilha Somente leitura.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > copy(positionType: WorksheetPositionType, relativeTo: Worksheet)|Copia uma planilha e a coloca na posição especificada. Retorna a planilha copiada.|1.7|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)|Obtém o objeto Range que começa em um determinado índice de linha e índice de coluna e que abrange um determinado número de linhas e colunas.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Property_ > type|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|_Property_ > worksheetId|Obtém a id da planilha que está ativada.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Property_ > source|Obtém a origem do evento. Os valores possíveis são: Local, Remote.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Property_ > type|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|_Property_ > worksheetId|Obtém a id da planilha que foi adicionada à pasta de trabalho.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Property_ > address|Obtém o endereço de intervalo que representa a área alterada de uma planilha específica.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Property_ > changeType|Obtém o tipo de alteração que representa como o evento Changed é disparado. Os valores possíveis são: Others, RangeEdited, RowInserted, RowDeleted, ColumnInserted, ColumnDeleted, CellInserted, CellDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Property_ > source|Obtém a origem do evento. Os valores possíveis são: Local, Remote.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Property_ > type|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|_Property_ > worksheetId|Obtém a id da planilha na qual os dados foram alterados.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Property_ > type|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|_Property_ > worksheetId|Obtém a id da planilha que foi desativada.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Property_ > source|Obtém a origem do evento. Os valores possíveis são: Local, Remote.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Property_ > type|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|_Property_ > worksheetId|Obtém a id da planilha que foi excluída da pasta de trabalho.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Method_ > freezeAt(frozenRange: Range or string)|Define as células congeladas no modo de exibição da planilha ativa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Method_ > freezeColumns(count: number)|Congela a(s) primeira(s) coluna(s) da planilha in-loco.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Method_ > freezeRows(count: number)|Congele a(s) linha(s) superior(es) da planilha in-loco.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Method_ > getLocation()|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Method_ > getLocationOrNullObject()|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|1.7|
|[worksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|_Method_ > unfreeze()|Remove todos os painéis congelados na planilha.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowEditObjects|Representa a opção de proteção da planilha que permite a edição de objetos.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowEditScenarios|Representa a opção de proteção da planilha que permite a edição de cenários.|1.7|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Relationship_ > selectionMode|Representa a opção de proteção da planilha do modo de seleção.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Property_ > address|Obtém o endereço do intervalo que representa a área selecionada de uma planilha específica.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Property_ > type|Obtém o tipo de evento. Os valores possíveis são: WorksheetDataChanged, WorksheetSelectionChanged, WorksheetAdded, WorksheetActivated, WorksheetDeactivated, TableDataChanged, TableSelectionChanged, WorksheetDeleted.|1.7|
|[worksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|_Property_ > worksheetId|Obtém a id da planilha na qual a seleção foi alterada.|1.7|


## <a name="whats-new-in-excel-javascript-api-16"></a>Novidades na API JavaScript do Excel 1.6 

### <a name="conditional-formatting"></a>Formatação condicional

Introduz a formatação condicional de um intervalo. Permite os seguintes tipos de formatação condicional:

* Escala de cores
* Barra de dados
* Conjunto de ícones
* Personalizado

Além disso:

* Retorna o intervalo ao qual o formato condicional foi aplicado. 
* Remoção da formatação condicional. 
* Fornece a prioridade e a funcionalidade stopifTrue. 
* Obtém a coleção de toda a formatação condicional em um determinado intervalo. 
* Limpa todos os formatos condicionais ativos no intervalo especificado atual. 

|Object| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[application](/javascript/api/excel/excel.application)|_Method_ > suspendApiCalculationUntilNextSync()|Suspende o cálculo até que o próximo "context.sync()" seja chamado. Uma vez definido, é responsabilidade do desenvolvedor recalcular a pasta de trabalho para garantir que todas as dependências sejam propagadas.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relationship_ > format|Retorna um objeto format, que encapsula a fonte de formatos condicionais, o preenchimento, as bordas e outras propriedades. Somente leitura.|1.6|
|[cellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)|_Relationship_ > rule|Representa o objeto Rule neste formato condicional.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Property_ > threeColorScale|Caso verdadeiro, a escala de cores terá três pontos (mínimo, médio, máximo). Caso contrário, terá dois (mínimo, máximo). Somente leitura.|1.6|
|[colorScaleConditionalFormat](/javascript/api/excel/excel.colorscaleconditionalformat)|_Relationship_ > criteria|Os critérios da escala de cores. O ponto médio é opcional ao se usar uma escala de cores de dois pontos.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Property_ > formula1|A fórmula, se necessário, para avaliar a regra de formato condicional.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Property_ > formula2|A fórmula, se necessário, para avaliar a regra de formato condicional.|1.6|
|[conditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)|_Property_ > operator|O operador do formato condicional de texto. Os valores possíveis são: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relationship_ > maximum|O critério de escala de cores de ponto máximo.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relationship_ > midpoint|O critério de escala de cores de ponto médio, caso a escala de cores seja uma escala de três cores.|1.6|
|[conditionalColorScaleCriteria](/javascript/api/excel/excel.conditionalcolorscalecriteria)|_Relationship_ > minimum|O critério de escala de cores de ponto mínimo.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Property_ > color|Representação do código de cores HTML da cor da escala de cores. Por exemplo, #FF0000 representa vermelho.|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Property_ > formula|Um número, uma fórmula ou nulo (se Type for LowestValue).|1.6|
|[conditionalColorScaleCriterion](/javascript/api/excel/excel.conditionalcolorscalecriterion)|_Property_ > type|No que a fórmula condicional de ícone deve se basear. Os valores possíveis são: Invalid, LowestValue, HighestValue, Number, Percent, Formula, Percentile.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > borderColor|Código de cores HTML que representa a cor da linha da borda, no formato #RRGGBB (p. ex., "FFA500") ou como uma cor HTML nomeada (p. ex., "cor-de-laranja").|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > fillColor|Código de cores HTML que representa a cor de preenchimento, no formato #RRGGBB (p. ex., "FFA500") ou como uma cor HTML nomeada (p. ex., "cor-de-laranja").|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > matchPositiveBorderColor|Representação booleana para indicar se o DataBar negativo tem ou não a mesma cor de borda que o DataBar positivo.|1.6|
|[conditionalDataBarNegativeFormat](/javascript/api/excel/excel.conditionaldatabarnegativeformat)|_Property_ > matchPositiveFillColor|Representação booleana para indicar se o DataBar negativo tem ou não a mesma cor de preenchimento que o DataBar positivo.|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Property_ > borderColor|Código de cores HTML que representa a cor da linha da borda, no formato #RRGGBB (p. ex., "FFA500") ou como uma cor HTML nomeada (p. ex., "cor-de-laranja").|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Property_ > fillColor|Código de cores HTML que representa a cor de preenchimento, no formato #RRGGBB (p. ex., "FFA500") ou como uma cor HTML nomeada (p. ex., "cor-de-laranja").|1.6|
|[conditionalDataBarPositiveFormat](/javascript/api/excel/excel.conditionaldatabarpositiveformat)|_Property_ > gradientFill|Representação booleana para indicar se o DataBar tem um gradiente ou não.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Property_ > formula|A fórmula, se necessário, para avaliar a regra do databar.|1.6|
|[conditionalDataBarRule](/javascript/api/excel/excel.conditionaldatabarrule)|_Property_ > type|O tipo de regra para o databar. Os valores possíveis são: LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Property_ > id|A prioridade do formato condicional dentro do ConditionalFormatCollection atual. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Property_ > priority|A prioridade (ou índice) dentro da coleção de formatos condicionais na qual esse formato condicional se encontra atualmente. A alteração  disso também|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Property_ > stopIfTrue|Se as condições desse formato condicional forem atendidas, nenhum formato de prioridade mais baixa terá efeito nessa célula.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Property_ > type|Um tipo de formato condicional. É possível definir somente um por vez. Somente Leitura. Somente leitura. Os valores possíveis são: Custom, DataBar, ColorScale, IconSet.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > cellValue|Retornará as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo de CellValue. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > cellValueOrNullObject|Retornará as propriedades do formato condicional do valor da célula se o formato condicional atual for um tipo de CellValue. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > colorScale|Retornará as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > colorScaleOrNullObject|Retornará as propriedades de formato condicional ColorScale se o formato condicional atual for um tipo ColorScale. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > custom|Retornará as propriedades do formato condicional personalizado se o formato condicional atual for um tipo custom. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > customOrNullObject|Retornará as propriedades do formato condicional personalizado se o formato condicional atual for um tipo custom. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > dataBar|Retornará as propriedades da barra de dados se o formato condicional atual for uma barra de dados. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > dataBarOrNullObject|Retornará as propriedades da barra de dados se o formato condicional atual for uma barra de dados. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > iconSet|Retornará as propriedades do formato condicional de IconSet se o formato condicional atual for um tipo IconSet. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > iconSetOrNullObject|Retornará as propriedades do formato condicional de IconSet se o formato condicional atual for um tipo IconSet. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > preset|Retornará o formato condicional de critérios predefinidos, como as propriedades above averagebelow averageunique valuescontains blanknonblankerrornoerror. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > presetOrNullObject|Retornará o formato condicional de critérios predefinidos, como as propriedades above averagebelow averageunique valuescontains blanknonblankerrornoerror. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > textComparison|Retornará as propriedades do formato condicional do texto específico se o formato condicional atual for um tipo text. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > textComparisonOrNullObject|Retornará as propriedades do formato condicional do texto específico se o formato condicional atual for um tipo text. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > topBottom|Retornará as propriedades do formato condicional TopBottom se o formato condicional atual for um tipo TopBottom. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Relationship_ > topBottomOrNullObject|Retornará as propriedades do formato condicional TopBottom se o formato condicional atual for um tipo TopBottom. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Method_ > delete()|Exclui esse formato condicional.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Method_ > getRange()|Retornará o intervalo ao qual o formato condicional está aplicado ou um objeto nulo se o intervalo for descontínuo. Somente leitura.|1.6|
|[conditionalFormat](/javascript/api/excel/excel.conditionalformat)|_Method_ > getRangeOrNullObject()|Retornará o intervalo ao qual o formato condicional está aplicado ou um objeto nulo se o intervalo for descontínuo. Somente leitura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Property_ > items|Uma coleção de objetos conditionalFormat. Somente leitura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Method_ > add(type: string)|Adiciona um novo formato condicional à coleção na prioridade firsttop.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Method_ > clearAll()|Limpa todos os formatos condicionais ativos no intervalo especificado atual.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Method_ > getCount()|Retorna o número de formatos condicionais na pasta de trabalho. Somente leitura.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Method_ > getItem(id: string)|Retorna um formato condicional para a ID especificada.|1.6|
|[conditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|_Method_ > getItemAt(index: number)|Retorna um formato condicional no índice fornecido.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Property_ > formula|A fórmula, se necessário, para avaliar a regra de formato condicional.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Property_ > formulaLocal|A fórmula, se necessário, para avaliar a regra de formato condicional no idioma do usuário.|1.6|
|[conditionalFormatRule](/javascript/api/excel/excel.conditionalformatrule)|_Property_ > formulaR1C1|A fórmula, se necessário, para avaliar a regra de formato condicional em notação de estilo R1C1.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Property_ > formula|Um número ou uma fórmula, dependendo do tipo.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Property_ > operator|GreaterThan ou GreaterThanOrEqual para cada tipo rule para o formato condicional Icon. Os valores possíveis são: Invalid, GreaterThan, GreaterThanOrEqual.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relationship_ > customIcon|O ícone personalizado para o critério atual caso seja diferente do IconSet padrão; caso contrário, será retornado nulo.|1.6|
|[conditionalIconCriterion](/javascript/api/excel/excel.conditionaliconcriterion)|_Relationship_ > type|No que a fórmula condicional de ícone deve se basear.|1.6|
|[conditionalPresetCriteriaRule](/javascript/api/excel/excel.conditionalpresetcriteriarule)|_Property_ > criterion|O critério do formato condicional. Os valores possíveis são: Invalid, Blanks, NonBlanks, Errors, NonErrors, Yesterday, Today, Tomorrow, LastSevenDays, LastWeek, ThisWeek, NextWeek, LastMonth, ThisMonth, NextMonth, AboveAverage, BelowAverage, EqualOrAboveAverage, EqualOrBelowAverage, OneStdDevAboveAverage, OneStdDevBelowAverage, TwoStdDevAboveAverage, TwoStdDevBelowAverage, ThreeStdDevAboveAverage, ThreeStdDevBelowAverage, UniqueValues, DuplicateValues.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Property_ > color|Código de cores HTML que representa a cor da linha da borda, no formato #RRGGBB (p. ex., "FFA500") ou como uma cor HTML nomeada (p. ex., "cor-de-laranja").|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Property_ > id|Representa o identificador da borda. Somente leitura. Os valores possíveis são: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Property_ > sideIndex|Valor constante que indica o lado específico da borda. Somente leitura. Os valores possíveis são: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](/javascript/api/excel/excel.conditionalrangeborder)|_Property_ > style|Uma das constantes de estilo da linha especificando o estilo da linha da borda. Os valores possíveis são: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Property_ > count|Número de objetos de borda da coleção. Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Property_ > items|Uma coleção de objetos conditionalRangeBorder. Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relationship_ > bottom|Obtém a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relationship_ > left|Obtém a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relationship_ > right|Obtém a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Relationship_ > top|Obtém a borda superior Somente leitura.|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Method_ > getItem(index: string)|Obtém um objeto border usando seu nome|1.6|
|[conditionalRangeBorderCollection](/javascript/api/excel/excel.conditionalrangebordercollection)|_Method_ > getItemAt(index: number)|Obtém um objeto border usando seu índice.|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Property_ > color|Código de cores HTML que representa a cor do preenchimento, no formato #RRGGBB (p. ex., "FFA500") ou como uma cor HTML nomeada (p. ex., "cor-de-laranja").|1.6|
|[conditionalRangeFill](/javascript/api/excel/excel.conditionalrangefill)|_Method_ > clear()|Redefine o preenchimento.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Property_ > bold|Representa o status de negrito da fonte.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Property_ > color|Representação do código de cores HTML para a cor do texto. Por exemplo, #FF0000 representa vermelho.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Property_ > italic|Representa o status de itálico da fonte.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Property_ > strikethrough|Representa o status de tachado da fonte.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Property_ > underline|Tipo de sublinhado aplicado à fonte. Os valores possíveis são: None, Single, Double.|1.6|
|[conditionalRangeFont](/javascript/api/excel/excel.conditionalrangefont)|_Method_ > clear()|Redefine os formatos de fonte.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Property_ > numberFormat|Representa o código de formato de número do Excel para o intervalo específico. Desmarcado se nulo for passado.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relationship_ > borders|Coleção de objetos border que se aplica ao intervalo do formato condicional geral. Somente leitura.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relationship_ > fill|Retorna o objeto fill definido no intervalo do formato condicional geral. Somente leitura.|1.6|
|[conditionalRangeFormat](/javascript/api/excel/excel.conditionalrangeformat)|_Relationship_ > font|Retorna o objeto font definido no intervalo do formato condicional geral. Somente leitura.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Property_ > operator|O operador do formato condicional do texto. Os valores possíveis são: Invalid, Contains, NotContains, BeginsWith, EndsWith.|1.6|
|[conditionalTextComparisonRule](/javascript/api/excel/excel.conditionaltextcomparisonrule)|_Property_ > text|O valor Text do formato condicional.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Property_ > rank|A classificação entre 1 e 1.000 para classificações numéricas ou 1 e 100 para classificações percentuais.|1.6|
|[conditionalTopBottomRule](/javascript/api/excel/excel.conditionaltopbottomrule)|_Property_ > type|Formata valores com base na classificação superior ou inferior. Os valores possíveis são: Invalid, TopItems, TopPercent, BottomItems, BottomPercent.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relationship_ > format|Retorna um objeto format, que encapsula a fonte de formatos condicionais, o preenchimento, as bordas e outras propriedades. Somente leitura.|1.6|
|[customConditionalFormat](/javascript/api/excel/excel.customconditionalformat)|_Relationship_ > rule|Representa o objeto Rule neste formato condicional. Somente leitura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Property_ > axisColor|Código de cores HTML que representa a cor da linha Axis, no formato #RRGGBB (p. ex., "FFA500") ou como uma cor HTML nomeada (p. ex., "cor-de-laranja").|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Property_ > axisFormat|Representação de como o eixo é determinado para uma barra de dados do Excel. Os valores possíveis são: Automatic, None, CellMidPoint.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Property_ > barDirection|Representa a direção em que o gráfico de barras de dados deve ser baseado. Os valores possíveis são: Context, LeftToRight, RightToLeft.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Property_ > showDataBarOnly|Caso verdadeiro, oculta os valores das células onde a barra de dados é aplicada.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relationship_ > lowerBoundRule|A regra para o que constitui o limite inferior (e como calculá-lo, se aplicável) para uma barra de dados.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relationship_ > negativeFormat|Representação de todos os valores à esquerda do eixo em uma barra de dados do Excel. Somente leitura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relationship_ > positiveFormat|Representação de todos os valores à direita do eixo em uma barra de dados do Excel. Somente leitura.|1.6|
|[dataBarConditionalFormat](/javascript/api/excel/excel.databarconditionalformat)|_Relationship_ > upperBoundRule|A regra para o que constitui o limite superior (e como calculá-lo, se aplicável) para uma barra de dados.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Property_ > reverseIconOrder|Caso verdadeiro, inverte as ordens de ícones para IconSet. Observe que não será possível definir isso se ícones personalizados forem usados.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Property_ > showIconOnly|Caso verdadeiro, oculta os valores e mostra somente ícones.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Property_ > style|Caso definido, exibe a opção IconSet do formato condicional. Os valores possíveis são: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.6|
|[iconSetConditionalFormat](/javascript/api/excel/excel.iconsetconditionalformat)|_Relationship_ > criteria|Uma matriz de Criteria e IconSets para as regras e os possíveis ícones personalizados para ícones condicionais. Observe que, para o primeiro critério, apenas o ícone personalizado pode ser modificado, enquanto o tipo, a fórmula e o operador serão ignorados quando definidos.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relationship_ > format|Retorna um objeto format, que encapsula a fonte de formatos condicionais, o preenchimento, as bordas e outras propriedades. Somente leitura.|1.6|
|[presetCriteriaConditionalFormat](/javascript/api/excel/excel.presetcriteriaconditionalformat)|_Relationship_ > rule|A regra do formato condicional.|1.6|
|[range](/javascript/api/excel/excel.range)|_Relationship_ > conditionalFormats|Coleção de ConditionalFormats que formam uma intersecção do intervalo. Somente leitura.|1.6|
|[range](/javascript/api/excel/excel.range)|_Method_ > calculate()|Calcula um intervalo de células em uma planilha.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relationship_ > format|Retorna um objeto format, que encapsula a fonte de formatos condicionais, o preenchimento, as bordas e outras propriedades. Somente leitura.|1.6|
|[textConditionalFormat](/javascript/api/excel/excel.textconditionalformat)|_Relationship_ > rule|A regra do formato condicional.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relationship_ > format|Retorna um objeto format, que encapsula a fonte de formatos condicionais, o preenchimento, as bordas e outras propriedades. Somente leitura.|1.6|
|[topBottomConditionalFormat](/javascript/api/excel/excel.topbottomconditionalformat)|_Relationship_ > rule|Os critérios do formato condicional TopBottom.|1.6|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > internalTest|Apenas para uso interno. Somente leitura.|1.6|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > calculate(markAllDirty: bool)|Calcula todas as células em uma planilha.|1.6|

##  <a name="whats-new-in-excel-javascript-api-15"></a>Novidades na API JavaScript do Excel 1.5

### <a name="custom-xml-part"></a>Parte XML personalizada

* Adição de uma coleção de partes XML personalizadas ao objeto workbook.
* Obter parte XML personalizada usando ID
* Obter uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace específico.
* Obter uma sequência de caracteres XML associada a uma parte.
* Fornecer id e namespace de uma parte.
* Adicionar uma nova parte XML personalizada à pasta de trabalho.
* Definir a parte XML inteira.
* Excluir uma parte XML personalizada.
* Excluir um atributo com o nome especificado do elemento identificado por xpath.
* Consultar o conteúdo XML por xpath.
* Inserir, atualizar e excluir o atributo.

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

|Object| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Property_ > id|ID da parte XML personalizada. Somente leitura.|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Property_ > namespaceUri|URI do namespace da parte XML personalizada. Somente leitura.|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Method_ > delete()|Exclui a parte XML personalizada.|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Method_ > getXml()|Obtém o conteúdo XML completo da parte XML personalizada.|1.5|
|[customXmlPart](/javascript/api/excel/excel.customxmlpart)|_Method_ > setXml(xml: string)|Define o conteúdo XML personalizado da parte XML completa.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Property_ > items|Uma coleção de objetos customXmlPart. Somente leitura.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Method_ > add(type: string)|Adicionar uma nova parte XML personalizada à pasta de trabalho.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Method_ > getByNamespace(namespaceUri: string)|Obtém uma nova coleção com escopo de partes XML personalizadas cujos namespaces correspondem ao namespace especificado.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Method_ > getCount()|Obtém o número de partes CustomXml na coleção.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Method_ > getItem(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1.5|
|[customXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|_Method_ > getItemOrNullObject(key: string)|Obtém uma parte XML personalizada com base em sua ID.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Property_ > items|Uma coleção de objetos customXmlPartScoped. Somente leitura.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Method_ > getCount()|Obtém o número de partes CustomXML nesta coleção.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Method_ > getItem(id: string)|Obtém uma parte XML personalizada com base em sua ID.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Method_ > getItemOrNullObject(key: string)|Obtém uma parte XML personalizada com base em sua ID.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Method_ > getOnlyItem()|Se a coleção contiver exatamente um item, esse método o retornará.|1.5|
|[customXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|_Method_ > getOnlyItemOrNullObject()|Se a coleção contiver exatamente um item, esse método o retornará.|1.5|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > customXmlParts|Representa a coleção de partes XML personalizadas contidas nesta pasta de trabalho. Somente leitura.|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > getNext(visibleOnly: bool)|Obtém a planilha posterior a esta. Se não houver nenhuma planilha após esta, este método gerará um erro.|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > getNextOrNullObject(visibleOnly: bool)|Obtém a planilha posterior a esta. Se não houver nenhuma planilha após esta, este método retornará um objeto nulo.|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > getPrevious(visibleOnly: bool)|Obtém a planilha anterior a esta. Se não houver nenhuma planilha anterior, este método gerará um erro.|1.5|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > getPreviousOrNullObject(visibleOnly: bool)|Obtém a planilha anterior a esta. Se não houver nenhuma planilha anterior, este método retornará um objeto nulo.|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Method_ > getFirst(visibleOnly: bool)|Obtém a primeira planilha na coleção.|1.5|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Method_ > getLast(visibleOnly: bool)|Obtém a última planilha na coleção.|1.5|

## <a name="whats-new-in-excel-javascript-api-14"></a>Novidades na API JavaScript do Excel 1.4
A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.4.

### <a name="named-item-add-and-new-properties"></a>Adicionar item nomeado e novas propriedades

Novas propriedades:

* `comment`
* `scope` itens com escopo de planilha ou pasta de trabalho
* `worksheet` retorna a planilha que o item nomeado tem como escopo.

Novos métodos:

* `add(name: string, reference: Range or string, comment: string)`Adiciona um novo nome à coleção do escopo específico.
* `addFormulaLocal(name: string, formula: string, comment: string)` Adiciona um novo nome à coleção do escopo específico usando a localidade do usuário para a fórmula.

### <a name="settings-api-in-in-excel-namespace"></a>Configurações de API no namespace do Excel

O objeto [Setting](/javascript/api/excel/excel.setting) representa um par chave-valor de uma configuração persistentes ao documento. Agora, adicionamos APIs relacionadas às configurações no namespace do Excel. Isso não oferece uma nova funcionalidade de rede. No entanto, assim é mais fácil manter a sintaxe de API com base em lote prometido para reduzir a dependência em tarefas comuns relacionadas à API para Excel.

As APIs incluem `getItem()` para obter a entrada de configuração por meio da chave, `add()` para adicionar o par de configuração chave:valor especificado na pasta de trabalho.

### <a name="others"></a>Outros

* Definir o nome da coluna de tabela (a versão anterior permite somente leitura).
* Adicionar coluna de tabela ao fim da tabela (a versão anterior permite apenas qualquer lugar, exceto o último).
* Adicione várias linhas a uma tabela de cada vez (a versão anterior só permite uma linha por vez).
* `range.getColumnsAfter(count: number)` e `range.getColumnsBefore(count: number)` para obter determinado número de colunas à direita/esquerda do objeto Range atual.
* Obter item ou função de objeto null: Essa funcionalidade permite obter o objeto utilizando uma chave. Se o objeto não existir, a propriedade isNullObject do objeto retornado será verdadeira. Isso permite que os desenvolvedores verifiquem se existe um objeto ou não sem ter de lidar com ele por meio da manipulação de exceção. Disponível na planilha, item nomeado, associação, série de gráficos etc.

    ```javascript
    worksheet.GetItemOrNullObject()
    ```

|Object| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Method_ > getCount()|Obtém o número de associações da coleção.|1.4|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Method_ > getItemOrNullObject(key: string)|Obtém um objeto binding pela ID. Se o objeto binding não existir, retornará um objeto null.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Method_ > getCount()|Retorna o número de gráficos da planilha.|1.4|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Method_ > getItemOrNullObject(key: string)|Obtém um gráfico usando seu nome. Quando houver vários gráficos com o mesmo nome, o primeiro deles será retornado.|1.4|
|[chartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|_Method_ > getCount()|Retorna o número de pontos do gráfico da série.|1.4|
|[chartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|_Method_ > getCount()|Retorna o número de série da coleção.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Property_ > comment|Representa o comentário associado a esse nome.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Property_ > scope|Indica se o nome tem escopo para a pasta de trabalho ou uma planilha específica. Somente leitura. Os valores possíveis são: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relationship_ > worksheet|Retorna a planilha em que o item nomeado tem escopo. Gerará um erro se os itens tiverem escopo para a pasta de trabalho. Somente leitura.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Relationship_ > worksheetOrNullObject|Retorna a planilha em que o item nomeado tem escopo. Retornará um objeto null se o item tiver escopo para a pasta de trabalho. Somente leitura.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Method_ > delete()|Exclui o nome fornecido.|1.4|
|[namedItem](/javascript/api/excel/excel.nameditem)|_Method_ > getRangeOrNullObject()|Retorna o objeto Range associado ao nome. Retornará um objeto null se o tipo do item nomeado não for um intervalo.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Method_ > add(name: string, reference: Range or string, comment: string)|Adiciona um novo nome à coleção do escopo específico.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Method_ > addFormulaLocal(name: string, formula: string, comment: string)|Adiciona um novo nome à coleção do escopo específico usando a localidade do usuário para a fórmula.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Method_ > getCount()|Obtém o número de itens nomeados na coleção.|1.4|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Method_ > getItemOrNullObject(key: string)|Obtém um objeto NamedItem usando seu nome. Se o objeto NamedItem não existir, um objeto null será retornado.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Method_ > getCount()|Obtém o número de tabelas dinâmicas na coleção.|1.4|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Method_ > getItemOrNullObject(key: string)|Obtém uma Tabela Dinâmica por nome. Se a tabela dinâmica não existir, um objeto null será retornado.|1.4|
|[range](/javascript/api/excel/excel.range)|_Method_ > getIntersectionOrNullObject(anotherRange: Range or string)|Obtém o objeto range que representa a intersecção retangular dos intervalos específicos. Se nenhuma intersecção for encontrada, um objeto null será retornado.|1.4|
|[range](/javascript/api/excel/excel.range)|_Method_ > getUsedRangeOrNullObject(valuesOnly: bool)|Retorna o intervalo usado do objeto range específico. Se não houver nenhuma célula usada no intervalo, esta função retornará um objeto null.|1.4|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Method_ > getCount()|Obtém o número de objetos RangeView na coleção.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Property_ > key|Retorna a chave que representa a id da configuração. Somente leitura.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Property_ > value|Representa o valor armazenado para esta configuração.|1.4|
|[setting](/javascript/api/excel/excel.setting)|_Method_ > delete()|Exclui a configuração.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Property_ > items|Uma coleção de objetos setting. Somente leitura.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > add(key: string, value: (any))|Define ou adiciona a configuração especificada na pasta de trabalho.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > getCount()|Obtém o número de Configurações na coleção.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > getItem(key: string)|Obtém uma entrada Setting por meio da chave.|1.4|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > getItemOrNullObject(key: string)|Obtém uma entrada de configuração por meio da chave. Se a Configuração não existir, um objeto null será retornado.|1.4|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relationship_ > settings|Obtém o objeto Setting, que representa a associação que gerou o evento settingsChanged.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Method_ > getCount()|Obtém o número de tabelas na coleção.|1.4|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Method_ > getItemOrNullObject(key: number or string)|Obtém uma tabela pelo nome ou ID. Se a tabela não existir, um objeto null será retornado.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Method_ > getCount()|Obtém o número de colunas na tabela.|1.4|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Method_ > getItemOrNullObject(key: number or string)|Obtém um objeto column por nome ou ID. Se a coluna não existir, um objeto null será retornado.|1.4|
|[tableRowCollection](/javascript/api/excel/excel.tablerowcollection)|_Method_ > getCount()|Obtém a quantidade de linhas na tabela.|1.4|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > settings|Representa uma coleção de configurações associadas à pasta de trabalho. Somente leitura.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relationship_ > names|Coleção de nomes com escopo para a planilha atual. Somente leitura.|1.4|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Method_ > getUsedRangeOrNullObject(valuesOnly: bool)|O intervalo usado é o menor intervalo que abrange todas as células que têm um valor ou uma formatação atribuída a elas. Se a planilha inteira estiver em branco, esta função retornará um objeto null.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Method_ > getCount(visibleOnly: bool)|Obtém o número de planilhas na coleção.|1.4|
|[worksheetCollection](/javascript/api/excel/excel.worksheetcollection)|_Method_ > getItemOrNullObject(key: string)|Obtém um objeto worksheet usando seu Nome ou ID. Se a planilha não existir, um objeto null será retornado.|1.4|

## <a name="whats-new-in-excel-javascript-api-13"></a>Novidades na API JavaScript do Excel 1.3

A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.3.

|Object| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[binding](/javascript/api/excel/excel.binding)|_Method_ > delete()|Exclui a associação.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Method_ > add(range: Range or string, bindingType: string, id: string)|Adiciona uma nova associação a um intervalo específico.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Method_ > addFromNamedItem(name: string, bindingType: string, id: string)|Adiciona uma nova associação com base em um item nomeado na pasta de trabalho.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Método_ > addFromSelection (bindingType: cadeia de caracteres, id: cadeia de caracteres)|Adiciona uma nova associação com base na seleção atual.|1.3|
|[bindingCollection](/javascript/api/excel/excel.bindingcollection)|_Method_ > getItemOrNull(id: string)|Obtém um objeto binding pela ID. Se o objeto binding não existir, a propriedade isNull do objeto retornado será verdadeira.|1.3|
|[chartCollection](/javascript/api/excel/excel.chartcollection)|_Method_ > getItemOrNull(id: string)|Obtém um gráfico usando seu nome. Quando houver vários gráficos com o mesmo nome, o primeiro deles será retornado.|1.3|
|[namedItemCollection](/javascript/api/excel/excel.nameditemcollection)|_Method_ > getItemOrNull(id: string)|Obtém um objeto NamedItem usando seu nome. Se o objeto nameditem não existir, a propriedade isNull do objeto retornado será verdadeira.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Property_ > name|Nome da Tabela Dinâmica.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Relationship_ > worksheet|A planilha que contém a Tabela Dinâmica atual. Somente leitura.|1.3|
|[pivotTable](/javascript/api/excel/excel.pivottable)|_Method_ > refresh()|Atualiza a Tabela Dinâmica.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Property_ > items|Uma coleção de objetos pivotTable. Somente leitura.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Method_ > getItem(key: string)|Obtém uma Tabela Dinâmica por nome.|1.3|
|[pivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|_Method_ > getItemOrNull(id: string)|Obtém uma Tabela Dinâmica por nome. Se a Tabela Dinâmica não existir, a propriedade isNull do objeto retornado será verdadeira.|1.3|
|[range](/javascript/api/excel/excel.range)|_Method_ > getIntersectionOrNull(anotherRange: Range or string)|Obtém o objeto range que representa a intersecção retangular dos intervalos específicos. Se nenhuma intersecção for encontrada, um objeto null será retornado.|1.3|
|[range](/javascript/api/excel/excel.range)|_Method_ > getVisibleView()|Representa as linhas visíveis do intervalo atual.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > cellAddresses|Representa os endereços de célula da RangeView. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > columnCount|Retorna o número de colunas visíveis. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > formulas|Representa a fórmula em notação de estilo A1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > formulasLocal|Representa a fórmula em notação de estilo A1, no idioma do usuário e na localidade de formatação de número.  Por exemplo, a fórmula "=SUM(A1, introduced in 1.5)" em inglês seria "=SOMA(A1; 1,5)" em alemão.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > formulasR1C1|Representa a fórmula em notação de estilo L1C1.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > index|Retorna um valor que representa o índice de RangeView. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > numberFormat|Representa o código de formato de número do Excel para a célula específica.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > rowCount|Retorna o número de linhas visíveis. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > text|Valores Text do intervalo especificado. O valor Text não depende da largura da célula. A substituição pelo sinal #, que ocorre na interface de usuário do Excel, não afeta o valor de texto retornado pela API. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > valueTypes|Representa o tipo de dados de cada célula. Somente leitura. Os valores possíveis são: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Property_ > values|Representa os valores brutos da exibição do intervalo especificado. Os dados retornados podem ser dos tipos: sequência de caracteres, número ou booleano. A célula que contém um erro retornará a cadeia de caracteres de erro.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Relationship_ > rows|Representa uma coleção de exibições de intervalo associadas ao intervalo. Somente leitura.|1.3|
|[rangeView](/javascript/api/excel/excel.rangeview)|_Method_ > getRange()|Obtém o intervalo pai associado à RangeView atual.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Property_ > items|Uma coleção de objetos rangeView. Somente leitura.|1.3|
|[rangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|_Method_ > getItemAt(index: number)|Obtém uma linha RangeView através de seu índice. Indexado com zero.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Property_ > key|Retorna a chave que representa a id da configuração. Somente leitura.|1.3|
|[setting](/javascript/api/excel/excel.setting)|_Method_ > delete()|Exclui a configuração.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Property_ > items|Uma coleção de objetos setting. Somente leitura.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > getItem(key: string)|Obtém uma entrada Setting por meio da chave.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > getItemOrNull(key: string)|Obtém uma entrada Setting por meio da chave. Se Setting não existir, a propriedade isNull do objeto retornado será verdadeira.|1.3|
|[settingCollection](/javascript/api/excel/excel.settingcollection)|_Method_ > set(key: string, value: string)|Define ou adiciona a configuração especificada na pasta de trabalho.|1.3|
|[settingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|_Relationship_ > settingCollection|Obtém o objeto Setting, que representa a associação que gerou o evento settingsChanged.|1.3|
|[table](/javascript/api/excel/excel.table)|_Property_ > highlightFirstColumn|Indica se a primeira coluna contém uma formatação especial.|1.3|
|[table](/javascript/api/excel/excel.table)|_Property_ > highlightLastColumn|Indica se a última coluna contém uma formatação especial.|1.3|
|[table](/javascript/api/excel/excel.table)|_Property_ > showBandedColumns|Indica se as colunas mostram formatação em tiras nas quais as colunas ímpares são realçadas de modo diferente das colunas pares, tornando a leitura da tabela mais fácil.|1.3|
|[table](/javascript/api/excel/excel.table)|_Property_ > showBandedRows|Indica se as linhas mostram formatação em tiras nas quais as linhas ímpares são realçadas de modo diferente das linhas pares, tornando a leitura da tabela mais fácil.|1.3|
|[table](/javascript/api/excel/excel.table)|_Property_ > showFilterButton|Indica se os botões de filtro estão visíveis na parte superior de cada cabeçalho da coluna. Essa configuração só será permitida se a tabela tiver uma linha de cabeçalho.|1.3|
|[tableCollection](/javascript/api/excel/excel.tablecollection)|_Method_ > getItemOrNull(key: number or string)|Obtém uma tabela pelo nome ou ID. Se a tabela não existir, a propriedade isNull do objeto retornado será verdadeira.|1.3|
|[tableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|_Method_ > getItemOrNull(key: number or string)|Obtém um objeto column por nome ou ID. Se column não existir, a propriedade isNull do objeto retornado será verdadeira.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > pivotTables|Representa uma coleção de Tabelas Dinâmicas associadas à pasta de trabalho. Somente leitura.|1.3|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > settings|Representa uma coleção de configurações associadas à pasta de trabalho. Somente leitura.|1.3|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relationship_ > pivotTables|Coleção de Tabelas Dinâmicas que fazem parte da planilha. Somente leitura.|1.3|

## <a name="whats-new-in-excel-javascript-api-12"></a>Novidades na API JavaScript do Excel 1.2

A seguir estão as novas adições às APIs JavaScript do Excel no conjunto de requisitos 1.2.

|Object| Novidades| Descrição|Conjunto de requisitos|
|:----|:----|:----|:----|
|[chart](/javascript/api/excel/excel.chart)|_Property_ > id|Obtém um gráfico com base em sua posição na coleção. Somente leitura.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Relationship_ > worksheet|A planilha que contém o gráfico atual. Somente leitura.|1.2|
|[chart](/javascript/api/excel/excel.chart)|_Method_ > getImage(height: number, width: number, fittingMode: string)|Processa o gráfico como uma imagem codificada em base64, dimensionando o gráfico para se ajustar às dimensões especificadas.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Relationship_ > criteria|O filtro aplicado no momento à coluna fornecida. Somente leitura.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > apply(criteria: FilterCriteria)|Aplica os critérios de filtro determinados à coluna fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyBottomItemsFilter(count: number)|Aplica um filtro "Bottom Item" à coluna para obter o número de elementos fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyBottomPercentFilter(percent: number)]|Aplica um filtro "Bottom Percent" à coluna para obter o percentual de elementos fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyCellColorFilter(color: string)|Aplica um filtro "Cell Color" à coluna para obter a cor fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyCustomFilter(criteria1: string, criteria2: string, oper: string)|Aplica um filtro "Icon" à coluna para obter as sequências de caracteres de critérios fornecidas.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyDynamicFilter(criteria: string)|Aplica um filtro "Dynamic" à coluna.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyFontColorFilter(color: string)|Aplica um filtro "Font Color" à coluna para obter a cor fornecida.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyIconFilter(icon: Icon)|Aplica um filtro "Ícone" à coluna para obter o ícone fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyTopItemsFilter(count: number)|Aplica um filtro "Top Item" à coluna para obter o número de elementos fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyTopPercentFilter(percent: number)|Aplica um filtro "Top Percent" à coluna para obter o percentual de elementos fornecido.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > applyValuesFilter(values: ())|Aplica um filtro "Values" à coluna para obter os valores fornecidos.|1.2|
|[filter](/javascript/api/excel/excel.filter)|_Method_ > clear()|Desmarca o filtro na coluna determinada.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > color|A sequência de caracteres de cores HTML usada para filtrar células. Usada com as filtragens "cellColor" e "fontColor".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > criterion1|O primeiro critério usado para filtrar os dados. Usado como um operador no caso da filtragem "custom".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > criterion2|O segundo critério usado para filtrar os dados. Usado apenas como um operador no caso da filtragem "custom".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > dynamicCriteria|Os critérios dinâmicos do conjunto Excel.DynamicFilterCriteria a serem aplicados nessa coluna. Usados com a filtragem "dynamic". Os valores possíveis são: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > filterOn|A propriedade usada pelo filtro para determinar se os valores devem ficar visíveis. Os valores possíveis são: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > operator|O operador usado para combinar os critérios 1 e 2 ao usar a filtragem "custom". Os valores possíveis são: And, Or.|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Property_ > values|O conjunto de valores a serem usados como parte da filtragem "values".|1.2|
|[filterCriteria](/javascript/api/excel/excel.filtercriteria)|_Relationship_ > icon|O ícone usado para filtrar células. Usado com a filtragem "icon".|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Property_ > date|A data no formato ISO8601 usada para filtrar os dados.|1.2|
|[filterDatetime](/javascript/api/excel/excel.filterdatetime)|_Property_ > specificity|Como a data específica deve ser usada para manter os dados. Por exemplo, se a data for 2005-04-02 e a especificidade estiver definida como "mês", a operação de filtragem manterá todas as linhas com uma data do mês de abril de 2009. Os valores possíveis são: Year, Monday, Day, Hour, Minute, Second.|1.2|
|[FormatProtection](/javascript/api/excel/excel.formatprotection)|_Property_ > formulaHidden|Indica se o Excel oculta a fórmula para as células no intervalo. Um valor nulo indica que o intervalo inteiro não tem configuração uniforme da fórmula oculta.|1.2|
|[FormatProtection](/javascript/api/excel/excel.formatprotection)|_Property_ > locked|Indica se o Excel bloqueia as células no objeto. Um valor nulo indica que o intervalo inteiro não tem configuração uniforme de bloqueio.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Property_ > index|Representa o índice do ícone no conjunto fornecido.|1.2|
|[icon](/javascript/api/excel/excel.icon)|_Property_ > set|Representa o conjunto do qual ícone faz parte. Os valores possíveis são: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](/javascript/api/excel/excel.range)|_Property_ > columnHidden|Representa se todas as colunas do intervalo atual estão ocultas.|1.2|
|[range](/javascript/api/excel/excel.range)|_Property_ > formulasR1C1|Representa a fórmula em notação de estilo L1C1.|1.2|
|[range](/javascript/api/excel/excel.range)|_Property_ > hidden|Representa se todas as células do intervalo atual estão ocultas. Somente leitura.|1.2|
|[range](/javascript/api/excel/excel.range)|_Property_ > rowHidden|Representa se todas as linhas do intervalo atual estão ocultas.|1.2|
|[range](/javascript/api/excel/excel.range)|_Relationship_ > sort|Representa a classificação de intervalo do intervalo atual. Somente leitura.|1.2|
|[range](/javascript/api/excel/excel.range)|_Method_ > merge(across: bool)|Mescla as células do intervalo em uma região da planilha.|1.2|
|[range](/javascript/api/excel/excel.range)|_Method_ > unmerge()|Desfaz a mesclagem das células do intervalo em células separadas.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Property_ > columnWidth|Obtém ou define a largura de todas as colunas dentro do intervalo. Se as larguras das colunas não forem uniformes, nulo será retornado.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Property_ > rowHeight|Obtém ou define a altura de todas as linhas do intervalo. Se as alturas das linhas não forem uniformes, nulo será retornado.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Relationship_ > protection|Retorna o objeto de proteção de formato para um intervalo. Somente leitura.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Method_ > autofitColumns()|Altera a largura das colunas do intervalo atual para obter o melhor ajuste, com base nos dados atuais nas colunas.|1.2|
|[rangeFormat](/javascript/api/excel/excel.rangeformat)|_Method_ > autofitRows()|Altera a altura das linhas do intervalo atual para obter o melhor ajuste, com base nos dados atuais nas colunas.|1.2|
|[rangeReference](/javascript/api/excel/excel.rangereference)|_Property_ > address|Representa as linhas visíveis do intervalo atual.|1.2|
|[rangeSort](/javascript/api/excel/excel.rangesort)|_Method_ > apply(fields: SortField, matchCase: bool, hasHeaders: bool, orientation: string, method: string)|Executa uma operação de classificação.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Property_ > ascending|Indica se a classificação é feita de forma crescente.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Property_ > color|Representa a cor que é o destino da condição se a classificação estiver na cor da fonte ou da célula.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Property_ > dataOption|Representa as opções de classificação adicionais para esse campo. Os valores possíveis são: Normal, TextAsNumber.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Property_ > key|Representa a coluna (ou linha, dependendo da orientação da classificação) em que a condição está. Representado como um deslocamento da primeira coluna (ou linha).|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Property_ > sortOn|Representa o tipo de classificação dessa condição. Os valores possíveis são: Value, CellColor, FontColor, Icon.|1.2|
|[sortField](/javascript/api/excel/excel.sortfield)|_Relationship_ > icon|Representa o ícone que é o destino da condição se a classificação estiver no ícone da célula.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relationship_ > sort|Representa a classificação da tabela. Somente leitura.|1.2|
|[table](/javascript/api/excel/excel.table)|_Relationship_ > worksheet|A planilha que contém a tabela atual. Somente leitura.|1.2|
|[table](/javascript/api/excel/excel.table)|_Method_ > clearFilters()|Limpa todos os filtros aplicados à tabela no momento.|1.2|
|[table](/javascript/api/excel/excel.table)|_Method_ > convertToRange()|Converte a tabela em um intervalo de células normal. Todos os dados são preservados.|1.2|
|[table](/javascript/api/excel/excel.table)|_Method_ > reapplyFilters()|Reaplica todos os filtros à tabela no momento.|1.2|
|[tableColumn](/javascript/api/excel/excel.tablecolumn)|_Relationship_ > filter|Recupera o filtro aplicado à coluna. Somente leitura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Property_ > matchCase|Indica se o uso de maiúsculas ou minúsculas afetou a última classificação da tabela. Somente leitura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Property_ > method|Representa o último método de ordenação de caracteres chineses usado para classificar a tabela. Somente leitura. Os valores possíveis são: PinYin, StrokeCount.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Relationship_ > fields|Representa as condições atuais usadas para a última classificação da tabela. Somente leitura.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Method_ > apply(fields: SortField, matchCase: bool, method: string)|Executa uma operação de classificação.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Method_ > clear()|Desmarca a classificação que está na tabela no momento. Essa ação não modifica a ordenação da tabela, mas desmarca o estado dos botões do cabeçalho.|1.2|
|[tableSort](/javascript/api/excel/excel.tablesort)|_Method_ > reapply()|Reaplica os parâmetros de classificação atuais à tabela.|1.2|
|[workbook](/javascript/api/excel/excel.workbook)|_Relationship_ > functions|Representa uma instância de aplicativo do Excel que contém essa pasta de trabalho. Somente leitura.|1.2|
|[worksheet](/javascript/api/excel/excel.worksheet)|_Relationship_ > protection|Retorna o objeto de proteção da planilha para uma planilha. Somente leitura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Property_ > protected|Indica se a planilha está protegida. Somente Leitura. Somente leitura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Relationship_ > options|Opções de proteção da planilha. Somente leitura.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Method_ > protect(options: WorksheetProtectionOptions)|Protege uma planilha. Falhará se uma planilha estiver protegida.|1.2|
|[worksheetProtection](/javascript/api/excel/excel.worksheetprotection)|_Method_ > unprotect()|Desprotege uma planilha.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowAutoFilter|Representa a opção de proteção de planilha para permitir a utilização do recurso de filtro automático.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowDeleteColumns|Indica a opção de proteção de planilha para permitir a exclusão de colunas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowDeleteRows|Indica a opção de proteção de planilha para permitir a exclusão de linhas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowFormatCells|Indica a opção de proteção de planilha para permitir a formatação de células.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowFormatColumns|Indica a opção de proteção de planilha para permitir a formatação de colunas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowFormatRows|Indica a opção de proteção de planilha para permitir a formatação de linhas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowInsertColumns|Indica a opção de proteção de planilha para permitir a inserção de colunas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowInsertHyperlinks|Representa a opção de proteção de planilha para permitir a inserção de hiperlinks.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowInsertRows|Representa a opção de proteção de planilha para permitir a inserção de linhas.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowPivotTables|Representa a opção de proteção de planilha para permitir a utilização do recurso de Tabela Dinâmica.|1.2|
|[worksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|_Property_ > allowSort|Representa a opção de proteção de planilha para permitir a utilização do recurso de classificação.|1.2|

## <a name="excel-javascript-api-11"></a>API JavaScript do Excel 1.1

A API JavaScript do Excel 1.1 é a primeira versão da API. Para saber mais sobre a API, confira os tópicos de referência da [API JavaScript do Excel](/javascript/api/excel).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Especificar requisitos de API e hosts do Office](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Manifesto XML dos suplementos do Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
