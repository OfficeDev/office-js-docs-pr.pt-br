---
title: Conjunto de requisitos de API JavaScript do Excel 1,7
description: Detalhes sobre o conjunto de requisitos ExcelApi 1,7
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c84d099982225bae11cb3deba8a0503da0695aed
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771985"
---
# <a name="whats-new-in-excel-javascript-api-17"></a>Quais são as novidades na API JavaScript do Excel 1.7

O conjunto de requisitos 1.7 da API JavaScript do Excel incluei APIs para gráficos, eventos, planilhas, intervalos, propriedades do documento, itens nomeados, opções de proteção e estilos.

## <a name="customize-charts"></a>Personalize gráficos

Com as novas APIs de gráficos, você pode criar tipos degráficos adicionais, adicionar uma série de dados a um gráfico, definir o título do gráfico, adicionar um título de eixo, adicionar unidade de exibição, adicionar uma linha de tendência com média móvel, alterar uma linha de tendência para linear e muito mais. Estes são alguns exemplos:

* Eixo gráfico - obtenha, defina, formate e remova unidade de eixo, etiqueta e título em um gráfico.
* Série de gráficos - adicione, defina e exclua uma série em um gráfico.  Alterar marcadores da série, pedidos de plotagem e dimensionamento.
* Gráfico de linhas de tendências: adicione, receba e formate linhas de tendências em um gráfico.
* Legenda do gráfico - formate a fonte de legenda de um gráfico.
* Ponto do gráfico - defina a cor do ponto do gráfico.
* Subtítulo do título do gráfico - obtenha e defina a subseqüência do título para um gráfico.
* Tipo de gráfico - opção para criar mais tipos de gráfico.

## <a name="events"></a>Eventos

As APIs de eventos JavaScript do Excel fornecem diversos,  manipuladores de eventos que permitem que o suplemento execute automaticamente uma função designada quando ocorre um evento específico. Você pode criar essa função para executar as ações que seu cenário exige. Para obter uma lista de eventos que estão disponíveis, confira [trabalhar com eventos usando as API JavaScript do Excel](/office/dev/add-ins/excel/excel-add-ins-events).

## <a name="customize-the-appearance-of-worksheets-and-ranges"></a>Personalizar a aparência de planilhas e intervalos

Nas novas APIs você pode personalizar a aparência das planilhas de várias maneiras:

* Congele painéis para manter linhas ou colunas específicas visíveis durante a rolagem na planilha. Por exemplo, se a primeira linha da planilha inclui cabeçalhos, você pode congelá-la para que os cabeçalhos das colunas permaneçam visíveis enquanto rola para baixo na planilha.
* Modificar a cor da guia de planilha.
* Adicione títulos de planilha.

Você pode personalizar a aparência de intervalos de várias maneiras:

* Defina o estilo de célula para um intervalo para garantir que todas as células no intervalo tenham formatação consistente. Um estilo de célula é um conjunto definido de características de formatação, como fontes e tamanhos de fonte, formatos numéricos, bordas de célula e sombreamento de célula. Use qualquer um dos estilos de célula internas do Excel ou crie seu próprio estilo de célula personalizado.
* Defina a orientação de texto para um intervalo.
* Adicione ou modifique um hiperlink em um intervalo vinculado a outro local na pasta de trabalho ou a um local externo.

## <a name="manage-document-properties"></a>Gerenciar propriedades dos documentos

Usando as APIs de propriedades do documento, você pode acessar as propriedades do documento interno e também criar e gerenciar propriedades personalizadas do documento para armazenar o estado da pasta de trabalho e direcionar o fluxo de trabalho e a lógica comercial.

## <a name="copy-worksheets"></a>Copiar planilhas

Usando a cópia da planilha APIs, você pode copiar os dados e o formato de uma planilha para uma nova planilha na mesma pasta de trabalho e reduzir a quantidade de transferência de dados necessária.

## <a name="handle-ranges-with-ease"></a>Lidar com intervalos com facilidade

Usando várias APIs de intervalo, você pode fazer coisas como obter região ao redor, obter um intervalo redimensionado e muito mais. Essas APIs devem tornar as tarefas, como manipulação de intervalo e endereçamento, muito mais eficientes.

Além disso:

* Opções de proteção de pasta de trabalho e planilha - use estas APIs para proteger dados em uma planilha e a estrutura da pasta de trabalho.
* Atualizar um item nomeado - usar esta API para atualizar um item nomeado.
* Obter a célula ativa - usar esta API para acessar a célula ativa da pasta de trabalho.

## <a name="api-list"></a>Lista de APIs

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Chart](/javascript/api/excel/excel.chart)|[chartType](/javascript/api/excel/excel.chart#charttype)|Representa o tipo de gráfico. Confira Excel. ChartType para obter detalhes.|
||[id](/javascript/api/excel/excel.chart#id)|Id exclusiva do gráfico. Somente leitura.|
||[showAllFieldButtons](/javascript/api/excel/excel.chart#showallfieldbuttons)|Representa se deseja exibir todos os botões de campo em um Gráfico Dinâmico.|
|[ChartAreaFormat](/javascript/api/excel/excel.chartareaformat)|[Borderô](/javascript/api/excel/excel.chartareaformat#border)|Representa o formato da borda da área do gráfico, que inclui cores, LineStyle e Weight. Somente leitura.|
|[ChartAreaFormatData](/javascript/api/excel/excel.chartareaformatdata)|[Borderô](/javascript/api/excel/excel.chartareaformatdata#border)|Representa o formato da borda da área do gráfico, que inclui cores, LineStyle e Weight. Somente leitura.|
|[ChartAreaFormatLoadOptions](/javascript/api/excel/excel.chartareaformatloadoptions)|[Borderô](/javascript/api/excel/excel.chartareaformatloadoptions#border)|Representa o formato da borda da área do gráfico, que inclui cores, LineStyle e Weight.|
|[ChartAreaFormatUpdateData](/javascript/api/excel/excel.chartareaformatupdatedata)|[Borderô](/javascript/api/excel/excel.chartareaformatupdatedata#border)|Representa o formato da borda da área do gráfico, que inclui cores, LineStyle e Weight.|
|[ChartAxes](/javascript/api/excel/excel.chartaxes)|[getItem (tipo \| : "categoria \| " inválida "" valor " \| " Series ", Group?:" Primary " \| " Secondary ")](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Retorna o eixo específico identificado por tipo e grupo.|
||[getItem (tipo: Excel. ChartAxisType, Group?: Excel. ChartAxisGroup)](/javascript/api/excel/excel.chartaxes#getitem-type--group-)|Retorna o eixo específico identificado por tipo e grupo.|
|[ChartAxis](/javascript/api/excel/excel.chartaxis)|[baseTimeUnit](/javascript/api/excel/excel.chartaxis#basetimeunit)|Retorna ou define a unidade base para o eixo da categoria especificada.|
||[categoryType](/javascript/api/excel/excel.chartaxis#categorytype)|Retorna ou define o tipo de eixo de categoria.|
||[displayUnit](/javascript/api/excel/excel.chartaxis#displayunit)|Representa a unidade de exibição de eixo. Consulte Excel. ChartAxisDisplayUnit para obter detalhes.|
||[logBase](/javascript/api/excel/excel.chartaxis#logbase)|Representa a base do logaritmo ao usar escalas logarítmicas.|
||[majorTickMark](/javascript/api/excel/excel.chartaxis#majortickmark)|Representa o tipo de marca de escala principal para o eixo especificado. Consulte Excel. ChartAxisTickMark para obter detalhes.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxis#majortimeunitscale)|Retorna ou define o valor de escala de unidades principais para o eixo das categorias quando a propriedade CategoryType estiver definida como escala de tempo.|
||[minorTickMark](/javascript/api/excel/excel.chartaxis#minortickmark)|Representa o tipo de marca de escala secundária para o eixo especificado. Consulte Excel. ChartAxisTickMark para obter detalhes.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxis#minortimeunitscale)|Retorna ou define o valor da escala unitária secundária para o eixo da categoria quando a propriedade CategoryType estiver definida como TimeScale.|
||[axisGroup](/javascript/api/excel/excel.chartaxis#axisgroup)|Representa o grupo para o eixo especificado. Consulte Excel. ChartAxisGroup para obter detalhes. Somente leitura.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxis#customdisplayunit)|Representa o valor da unidade de exibição do eixo personalizado. Somente leitura. Para definir essa propriedade, use o método de SetCustomDisplayUnit(duplo).|
||[height](/javascript/api/excel/excel.chartaxis#height)|Representa a altura, em pontos, do eixo do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|
||[left](/javascript/api/excel/excel.chartaxis#left)|Representa a distância, em pontos, da borda esquerda do eixo à esquerda da área do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|
||[top](/javascript/api/excel/excel.chartaxis#top)|Representa a distância, em pontos, da borda superior do eixo a parte superior da área do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|
||[tipo](/javascript/api/excel/excel.chartaxis#type)|Representa o tipo de eixo. Consulte Excel. ChartAxisType para obter detalhes.|
||[width](/javascript/api/excel/excel.chartaxis#width)|Representa a largura, em pontos, do eixo do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxis#reverseplotorder)|Representa se o Microsoft Excel plota os pontos de dados do último para o primeiro.|
||[scaleType](/javascript/api/excel/excel.chartaxis#scaletype)|Representa o tipo de escala do eixo dos valores. Consulte Excel. ChartAxisScaleType para obter detalhes.|
||[setcategorynames (sourceData: intervalo)](/javascript/api/excel/excel.chartaxis#setcategorynames-sourcedata-)|Define todos os nomes de categoria para o eixo especificado.|
||[setCustomDisplayUnit (valor: número)](/javascript/api/excel/excel.chartaxis#setcustomdisplayunit-value-)|Definirá a unidade de exibição de eixo a um valor personalizado.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxis#showdisplayunitlabel)|Indica se a etiqueta de unidade de exibição de eixo está visível.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxis#ticklabelposition)|Representa a posição dos rótulos de marcas de escala no eixo especificado. Consulte Excel. ChartAxisTickLabelPosition para obter detalhes.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxis#ticklabelspacing)|Representa o número série ou categorias entre os rótulos de marcas de escala. Pode ser um valor de 1 a 31999 ou uma cadeia de caracteres vazia para configuração automática. O valor retornado sempre é um número.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxis#tickmarkspacing)|Representa o número de série ou categorias entre as marcas de escala.|
||[visible](/javascript/api/excel/excel.chartaxis#visible)|Um valor booliano representa a visibilidade do eixo.|
|[ChartAxisData](/javascript/api/excel/excel.chartaxisdata)|[axisGroup](/javascript/api/excel/excel.chartaxisdata#axisgroup)|Representa o grupo para o eixo especificado. Consulte Excel. ChartAxisGroup para obter detalhes. Somente leitura.|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxisdata#basetimeunit)|Retorna ou define a unidade base para o eixo da categoria especificada.|
||[categoryType](/javascript/api/excel/excel.chartaxisdata#categorytype)|Retorna ou define o tipo de eixo de categoria.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxisdata#customdisplayunit)|Representa o valor da unidade de exibição do eixo personalizado. Somente leitura. Para definir essa propriedade, use o método de SetCustomDisplayUnit(duplo).|
||[displayUnit](/javascript/api/excel/excel.chartaxisdata#displayunit)|Representa a unidade de exibição de eixo. Consulte Excel. ChartAxisDisplayUnit para obter detalhes.|
||[height](/javascript/api/excel/excel.chartaxisdata#height)|Representa a altura, em pontos, do eixo do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|
||[left](/javascript/api/excel/excel.chartaxisdata#left)|Representa a distância, em pontos, da borda esquerda do eixo à esquerda da área do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|
||[logBase](/javascript/api/excel/excel.chartaxisdata#logbase)|Representa a base do logaritmo ao usar escalas logarítmicas.|
||[majorTickMark](/javascript/api/excel/excel.chartaxisdata#majortickmark)|Representa o tipo de marca de escala principal para o eixo especificado. Consulte Excel. ChartAxisTickMark para obter detalhes.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisdata#majortimeunitscale)|Retorna ou define o valor de escala de unidades principais para o eixo das categorias quando a propriedade CategoryType estiver definida como escala de tempo.|
||[minorTickMark](/javascript/api/excel/excel.chartaxisdata#minortickmark)|Representa o tipo de marca de escala secundária para o eixo especificado. Consulte Excel. ChartAxisTickMark para obter detalhes.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisdata#minortimeunitscale)|Retorna ou define o valor da escala unitária secundária para o eixo da categoria quando a propriedade CategoryType estiver definida como TimeScale.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisdata#reverseplotorder)|Representa se o Microsoft Excel plota os pontos de dados do último para o primeiro.|
||[scaleType](/javascript/api/excel/excel.chartaxisdata#scaletype)|Representa o tipo de escala do eixo dos valores. Consulte Excel. ChartAxisScaleType para obter detalhes.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisdata#showdisplayunitlabel)|Indica se a etiqueta de unidade de exibição de eixo está visível.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisdata#ticklabelposition)|Representa a posição dos rótulos de marcas de escala no eixo especificado. Consulte Excel. ChartAxisTickLabelPosition para obter detalhes.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisdata#ticklabelspacing)|Representa o número série ou categorias entre os rótulos de marcas de escala. Pode ser um valor de 1 a 31999 ou uma cadeia de caracteres vazia para configuração automática. O valor retornado sempre é um número.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisdata#tickmarkspacing)|Representa o número de série ou categorias entre as marcas de escala.|
||[top](/javascript/api/excel/excel.chartaxisdata#top)|Representa a distância, em pontos, da borda superior do eixo a parte superior da área do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|
||[tipo](/javascript/api/excel/excel.chartaxisdata#type)|Representa o tipo de eixo. Consulte Excel. ChartAxisType para obter detalhes.|
||[visible](/javascript/api/excel/excel.chartaxisdata#visible)|Um valor booliano representa a visibilidade do eixo.|
||[width](/javascript/api/excel/excel.chartaxisdata#width)|Representa a largura, em pontos, do eixo do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|
|[ChartAxisLoadOptions](/javascript/api/excel/excel.chartaxisloadoptions)|[axisGroup](/javascript/api/excel/excel.chartaxisloadoptions#axisgroup)|Representa o grupo para o eixo especificado. Consulte Excel. ChartAxisGroup para obter detalhes. Somente leitura.|
||[baseTimeUnit](/javascript/api/excel/excel.chartaxisloadoptions#basetimeunit)|Retorna ou define a unidade base para o eixo da categoria especificada.|
||[categoryType](/javascript/api/excel/excel.chartaxisloadoptions#categorytype)|Retorna ou define o tipo de eixo de categoria.|
||[excede](/javascript/api/excel/excel.chartaxisloadoptions#crosses)|[Preterido; mantido para compatibilidade com as soluções de terceiros existentes]. Em vez `Position` disso, use.|
||[crossesAt](/javascript/api/excel/excel.chartaxisloadoptions#crossesat)|[Preterido; mantido para compatibilidade com as soluções de terceiros existentes]. Em vez `PositionAt` disso, use.|
||[customDisplayUnit](/javascript/api/excel/excel.chartaxisloadoptions#customdisplayunit)|Representa o valor da unidade de exibição do eixo personalizado. Somente leitura. Para definir essa propriedade, use o método de SetCustomDisplayUnit(duplo).|
||[displayUnit](/javascript/api/excel/excel.chartaxisloadoptions#displayunit)|Representa a unidade de exibição de eixo. Consulte Excel. ChartAxisDisplayUnit para obter detalhes.|
||[height](/javascript/api/excel/excel.chartaxisloadoptions#height)|Representa a altura, em pontos, do eixo do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|
||[left](/javascript/api/excel/excel.chartaxisloadoptions#left)|Representa a distância, em pontos, da borda esquerda do eixo à esquerda da área do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|
||[logBase](/javascript/api/excel/excel.chartaxisloadoptions#logbase)|Representa a base do logaritmo ao usar escalas logarítmicas.|
||[majorTickMark](/javascript/api/excel/excel.chartaxisloadoptions#majortickmark)|Representa o tipo de marca de escala principal para o eixo especificado. Consulte Excel. ChartAxisTickMark para obter detalhes.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisloadoptions#majortimeunitscale)|Retorna ou define o valor de escala de unidades principais para o eixo das categorias quando a propriedade CategoryType estiver definida como escala de tempo.|
||[minorTickMark](/javascript/api/excel/excel.chartaxisloadoptions#minortickmark)|Representa o tipo de marca de escala secundária para o eixo especificado. Consulte Excel. ChartAxisTickMark para obter detalhes.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisloadoptions#minortimeunitscale)|Retorna ou define o valor da escala unitária secundária para o eixo da categoria quando a propriedade CategoryType estiver definida como TimeScale.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisloadoptions#reverseplotorder)|Representa se o Microsoft Excel plota os pontos de dados do último para o primeiro.|
||[scaleType](/javascript/api/excel/excel.chartaxisloadoptions#scaletype)|Representa o tipo de escala do eixo dos valores. Consulte Excel. ChartAxisScaleType para obter detalhes.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisloadoptions#showdisplayunitlabel)|Indica se a etiqueta de unidade de exibição de eixo está visível.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisloadoptions#ticklabelposition)|Representa a posição dos rótulos de marcas de escala no eixo especificado. Consulte Excel. ChartAxisTickLabelPosition para obter detalhes.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisloadoptions#ticklabelspacing)|Representa o número série ou categorias entre os rótulos de marcas de escala. Pode ser um valor de 1 a 31999 ou uma cadeia de caracteres vazia para configuração automática. O valor retornado sempre é um número.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisloadoptions#tickmarkspacing)|Representa o número de série ou categorias entre as marcas de escala.|
||[top](/javascript/api/excel/excel.chartaxisloadoptions#top)|Representa a distância, em pontos, da borda superior do eixo a parte superior da área do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|
||[tipo](/javascript/api/excel/excel.chartaxisloadoptions#type)|Representa o tipo de eixo. Consulte Excel. ChartAxisType para obter detalhes.|
||[visible](/javascript/api/excel/excel.chartaxisloadoptions#visible)|Um valor booliano representa a visibilidade do eixo.|
||[width](/javascript/api/excel/excel.chartaxisloadoptions#width)|Representa a largura, em pontos, do eixo do gráfico. Nulo se o eixo não estiver visível. Somente leitura.|
|[ChartAxisUpdateData](/javascript/api/excel/excel.chartaxisupdatedata)|[baseTimeUnit](/javascript/api/excel/excel.chartaxisupdatedata#basetimeunit)|Retorna ou define a unidade base para o eixo da categoria especificada.|
||[categoryType](/javascript/api/excel/excel.chartaxisupdatedata#categorytype)|Retorna ou define o tipo de eixo de categoria.|
||[displayUnit](/javascript/api/excel/excel.chartaxisupdatedata#displayunit)|Representa a unidade de exibição de eixo. Consulte Excel. ChartAxisDisplayUnit para obter detalhes.|
||[logBase](/javascript/api/excel/excel.chartaxisupdatedata#logbase)|Representa a base do logaritmo ao usar escalas logarítmicas.|
||[majorTickMark](/javascript/api/excel/excel.chartaxisupdatedata#majortickmark)|Representa o tipo de marca de escala principal para o eixo especificado. Consulte Excel. ChartAxisTickMark para obter detalhes.|
||[majorTimeUnitScale](/javascript/api/excel/excel.chartaxisupdatedata#majortimeunitscale)|Retorna ou define o valor de escala de unidades principais para o eixo das categorias quando a propriedade CategoryType estiver definida como escala de tempo.|
||[minorTickMark](/javascript/api/excel/excel.chartaxisupdatedata#minortickmark)|Representa o tipo de marca de escala secundária para o eixo especificado. Consulte Excel. ChartAxisTickMark para obter detalhes.|
||[minorTimeUnitScale](/javascript/api/excel/excel.chartaxisupdatedata#minortimeunitscale)|Retorna ou define o valor da escala unitária secundária para o eixo da categoria quando a propriedade CategoryType estiver definida como TimeScale.|
||[reversePlotOrder](/javascript/api/excel/excel.chartaxisupdatedata#reverseplotorder)|Representa se o Microsoft Excel plota os pontos de dados do último para o primeiro.|
||[scaleType](/javascript/api/excel/excel.chartaxisupdatedata#scaletype)|Representa o tipo de escala do eixo dos valores. Consulte Excel. ChartAxisScaleType para obter detalhes.|
||[showDisplayUnitLabel](/javascript/api/excel/excel.chartaxisupdatedata#showdisplayunitlabel)|Indica se a etiqueta de unidade de exibição de eixo está visível.|
||[tickLabelPosition](/javascript/api/excel/excel.chartaxisupdatedata#ticklabelposition)|Representa a posição dos rótulos de marcas de escala no eixo especificado. Consulte Excel. ChartAxisTickLabelPosition para obter detalhes.|
||[tickLabelSpacing](/javascript/api/excel/excel.chartaxisupdatedata#ticklabelspacing)|Representa o número série ou categorias entre os rótulos de marcas de escala. Pode ser um valor de 1 a 31999 ou uma cadeia de caracteres vazia para configuração automática. O valor retornado sempre é um número.|
||[tickMarkSpacing](/javascript/api/excel/excel.chartaxisupdatedata#tickmarkspacing)|Representa o número de série ou categorias entre as marcas de escala.|
||[visible](/javascript/api/excel/excel.chartaxisupdatedata#visible)|Um valor booliano representa a visibilidade do eixo.|
|[ChartBorder](/javascript/api/excel/excel.chartborder)|[color](/javascript/api/excel/excel.chartborder#color)|Código de cor HTML que representa a cor das bordas no gráfico.|
||[lineStyle](/javascript/api/excel/excel.chartborder#linestyle)|Representa o estilo de linha da borda. Consulte Excel. ChartLineStyle para obter detalhes.|
||[Set (Propriedades: Excel. ChartBorder)](/javascript/api/excel/excel.chartborder#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartBorderUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.chartborder#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[weight](/javascript/api/excel/excel.chartborder#weight)|Representa a espessura da borda, em pontos.|
|[ChartBorderData](/javascript/api/excel/excel.chartborderdata)|[color](/javascript/api/excel/excel.chartborderdata#color)|Código de cor HTML que representa a cor das bordas no gráfico.|
||[lineStyle](/javascript/api/excel/excel.chartborderdata#linestyle)|Representa o estilo de linha da borda. Consulte Excel. ChartLineStyle para obter detalhes.|
||[weight](/javascript/api/excel/excel.chartborderdata#weight)|Representa a espessura da borda, em pontos.|
|[ChartBorderLoadOptions](/javascript/api/excel/excel.chartborderloadoptions)|[$all](/javascript/api/excel/excel.chartborderloadoptions#$all)||
||[color](/javascript/api/excel/excel.chartborderloadoptions#color)|Código de cor HTML que representa a cor das bordas no gráfico.|
||[lineStyle](/javascript/api/excel/excel.chartborderloadoptions#linestyle)|Representa o estilo de linha da borda. Consulte Excel. ChartLineStyle para obter detalhes.|
||[weight](/javascript/api/excel/excel.chartborderloadoptions#weight)|Representa a espessura da borda, em pontos.|
|[ChartBorderUpdateData](/javascript/api/excel/excel.chartborderupdatedata)|[color](/javascript/api/excel/excel.chartborderupdatedata#color)|Código de cor HTML que representa a cor das bordas no gráfico.|
||[lineStyle](/javascript/api/excel/excel.chartborderupdatedata#linestyle)|Representa o estilo de linha da borda. Consulte Excel. ChartLineStyle para obter detalhes.|
||[weight](/javascript/api/excel/excel.chartborderupdatedata#weight)|Representa a espessura da borda, em pontos.|
|[ChartCollectionLoadOptions](/javascript/api/excel/excel.chartcollectionloadoptions)|[chartType](/javascript/api/excel/excel.chartcollectionloadoptions#charttype)|Para cada ITEM na coleção: representa o tipo do gráfico. Confira Excel. ChartType para obter detalhes.|
||[id](/javascript/api/excel/excel.chartcollectionloadoptions#id)|Para cada ITEM na coleção: a ID exclusiva do gráfico. Somente leitura.|
||[showAllFieldButtons](/javascript/api/excel/excel.chartcollectionloadoptions#showallfieldbuttons)|Para cada ITEM na coleção: indica se todos os botões de campo devem ser exibidos em um gráfico dinâmico.|
|[ChartData](/javascript/api/excel/excel.chartdata)|[chartType](/javascript/api/excel/excel.chartdata#charttype)|Representa o tipo de gráfico. Confira Excel. ChartType para obter detalhes.|
||[id](/javascript/api/excel/excel.chartdata#id)|Id exclusiva do gráfico. Somente leitura.|
||[showAllFieldButtons](/javascript/api/excel/excel.chartdata#showallfieldbuttons)|Representa se deseja exibir todos os botões de campo em um Gráfico Dinâmico.|
|[ChartDataLabel](/javascript/api/excel/excel.chartdatalabel)|[position](/javascript/api/excel/excel.chartdatalabel#position)|Valor de DataLabelPosition que representa a posição do rótulo de dados. Consulte Excel. ChartDataLabelPosition para obter detalhes.|
||[divisória](/javascript/api/excel/excel.chartdatalabel#separator)|Cadeia de caracteres que representa o separador usado para o rótulo de dados em um gráfico.|
||[Set (Propriedades: Excel. ChartDataLabel)](/javascript/api/excel/excel.chartdatalabel#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartDataLabelUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.chartdatalabel#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabel#showbubblesize)|Valor booliano que determina se o tamanho da bolha do rótulo de dados fica visível ou não.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabel#showcategoryname)|Valor booliano que determina se o nome da categoria do rótulo de dados fica visível ou não.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabel#showlegendkey)|Valor booliano que determina se o código de legenda do rótulo de dados fica visível ou não.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabel#showpercentage)|Valor booliano que determina se o percentual do rótulo de dados fica visível ou não.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabel#showseriesname)|Valor booliano que determina se o nome da série do rótulo de dados fica visível ou não.|
||[showValue](/javascript/api/excel/excel.chartdatalabel#showvalue)|Valor booliano que determina se o valor do rótulo de dados fica visível ou não.|
|[ChartDataLabelData](/javascript/api/excel/excel.chartdatalabeldata)|[position](/javascript/api/excel/excel.chartdatalabeldata#position)|Valor de DataLabelPosition que representa a posição do rótulo de dados. Consulte Excel. ChartDataLabelPosition para obter detalhes.|
||[divisória](/javascript/api/excel/excel.chartdatalabeldata#separator)|Cadeia de caracteres que representa o separador usado para o rótulo de dados em um gráfico.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabeldata#showbubblesize)|Valor booliano que determina se o tamanho da bolha do rótulo de dados fica visível ou não.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabeldata#showcategoryname)|Valor booliano que determina se o nome da categoria do rótulo de dados fica visível ou não.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabeldata#showlegendkey)|Valor booliano que determina se o código de legenda do rótulo de dados fica visível ou não.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabeldata#showpercentage)|Valor booliano que determina se o percentual do rótulo de dados fica visível ou não.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabeldata#showseriesname)|Valor booliano que determina se o nome da série do rótulo de dados fica visível ou não.|
||[showValue](/javascript/api/excel/excel.chartdatalabeldata#showvalue)|Valor booliano que determina se o valor do rótulo de dados fica visível ou não.|
|[ChartDataLabelLoadOptions](/javascript/api/excel/excel.chartdatalabelloadoptions)|[$all](/javascript/api/excel/excel.chartdatalabelloadoptions#$all)||
||[position](/javascript/api/excel/excel.chartdatalabelloadoptions#position)|Valor de DataLabelPosition que representa a posição do rótulo de dados. Consulte Excel. ChartDataLabelPosition para obter detalhes.|
||[divisória](/javascript/api/excel/excel.chartdatalabelloadoptions#separator)|Cadeia de caracteres que representa o separador usado para o rótulo de dados em um gráfico.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelloadoptions#showbubblesize)|Valor booliano que determina se o tamanho da bolha do rótulo de dados fica visível ou não.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelloadoptions#showcategoryname)|Valor booliano que determina se o nome da categoria do rótulo de dados fica visível ou não.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelloadoptions#showlegendkey)|Valor booliano que determina se o código de legenda do rótulo de dados fica visível ou não.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelloadoptions#showpercentage)|Valor booliano que determina se o percentual do rótulo de dados fica visível ou não.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelloadoptions#showseriesname)|Valor booliano que determina se o nome da série do rótulo de dados fica visível ou não.|
||[showValue](/javascript/api/excel/excel.chartdatalabelloadoptions#showvalue)|Valor booliano que determina se o valor do rótulo de dados fica visível ou não.|
|[ChartDataLabelUpdateData](/javascript/api/excel/excel.chartdatalabelupdatedata)|[position](/javascript/api/excel/excel.chartdatalabelupdatedata#position)|Valor de DataLabelPosition que representa a posição do rótulo de dados. Consulte Excel. ChartDataLabelPosition para obter detalhes.|
||[divisória](/javascript/api/excel/excel.chartdatalabelupdatedata#separator)|Cadeia de caracteres que representa o separador usado para o rótulo de dados em um gráfico.|
||[showBubbleSize](/javascript/api/excel/excel.chartdatalabelupdatedata#showbubblesize)|Valor booliano que determina se o tamanho da bolha do rótulo de dados fica visível ou não.|
||[showCategoryName](/javascript/api/excel/excel.chartdatalabelupdatedata#showcategoryname)|Valor booliano que determina se o nome da categoria do rótulo de dados fica visível ou não.|
||[showLegendKey](/javascript/api/excel/excel.chartdatalabelupdatedata#showlegendkey)|Valor booliano que determina se o código de legenda do rótulo de dados fica visível ou não.|
||[showPercentage](/javascript/api/excel/excel.chartdatalabelupdatedata#showpercentage)|Valor booliano que determina se o percentual do rótulo de dados fica visível ou não.|
||[showSeriesName](/javascript/api/excel/excel.chartdatalabelupdatedata#showseriesname)|Valor booliano que determina se o nome da série do rótulo de dados fica visível ou não.|
||[showValue](/javascript/api/excel/excel.chartdatalabelupdatedata#showvalue)|Valor booliano que determina se o valor do rótulo de dados fica visível ou não.|
|[ChartFormatString](/javascript/api/excel/excel.chartformatstring)|[font](/javascript/api/excel/excel.chartformatstring#font)|Representa os atributos de fonte, como nome da fonte, tamanho da fonte, cor etc. do objeto de caracteres do gráfico.|
||[Set (Propriedades: Excel. ChartFormatString)](/javascript/api/excel/excel.chartformatstring#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartFormatStringUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.chartformatstring#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[ChartFormatStringData](/javascript/api/excel/excel.chartformatstringdata)|[font](/javascript/api/excel/excel.chartformatstringdata#font)|Representa os atributos de fonte, como nome da fonte, tamanho da fonte, cor etc. do objeto de caracteres do gráfico.|
|[ChartFormatStringLoadOptions](/javascript/api/excel/excel.chartformatstringloadoptions)|[$all](/javascript/api/excel/excel.chartformatstringloadoptions#$all)||
||[font](/javascript/api/excel/excel.chartformatstringloadoptions#font)|Representa os atributos de fonte, como nome da fonte, tamanho da fonte, cor etc. do objeto de caracteres do gráfico.|
|[ChartFormatStringUpdateData](/javascript/api/excel/excel.chartformatstringupdatedata)|[font](/javascript/api/excel/excel.chartformatstringupdatedata#font)|Representa os atributos de fonte, como nome da fonte, tamanho da fonte, cor etc. do objeto de caracteres do gráfico.|
|[ChartLegend](/javascript/api/excel/excel.chartlegend)|[height](/javascript/api/excel/excel.chartlegend#height)|Representa a altura, em pontos, da legenda no gráfico. NULL se a legenda não estiver visível.|
||[left](/javascript/api/excel/excel.chartlegend#left)|Representa a esquerda, em pontos, de uma legenda de gráfico. NULL se a legenda não estiver visível.|
||[legendEntries](/javascript/api/excel/excel.chartlegend#legendentries)|Representa uma coleção de legendEntries na legenda. Somente leitura.|
||[Ocultar sombra](/javascript/api/excel/excel.chartlegend#showshadow)|Representa se a legenda tem uma sombra no gráfico.|
||[top](/javascript/api/excel/excel.chartlegend#top)|Representa o início de uma legenda do gráfico.|
||[width](/javascript/api/excel/excel.chartlegend#width)|Representa a largura, em pontos, da legenda no gráfico. NULL se a legenda não estiver visível.|
|[ChartLegendData](/javascript/api/excel/excel.chartlegenddata)|[height](/javascript/api/excel/excel.chartlegenddata#height)|Representa a altura, em pontos, da legenda no gráfico. NULL se a legenda não estiver visível.|
||[left](/javascript/api/excel/excel.chartlegenddata#left)|Representa a esquerda, em pontos, de uma legenda de gráfico. NULL se a legenda não estiver visível.|
||[legendEntries](/javascript/api/excel/excel.chartlegenddata#legendentries)|Representa uma coleção de legendEntries na legenda. Somente leitura.|
||[Ocultar sombra](/javascript/api/excel/excel.chartlegenddata#showshadow)|Representa se a legenda tem uma sombra no gráfico.|
||[top](/javascript/api/excel/excel.chartlegenddata#top)|Representa o início de uma legenda do gráfico.|
||[width](/javascript/api/excel/excel.chartlegenddata#width)|Representa a largura, em pontos, da legenda no gráfico. NULL se a legenda não estiver visível.|
|[ChartLegendEntry](/javascript/api/excel/excel.chartlegendentry)|[Set (Propriedades: Excel. ChartLegendEntry)](/javascript/api/excel/excel.chartlegendentry#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartLegendEntryUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.chartlegendentry#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[visible](/javascript/api/excel/excel.chartlegendentry#visible)|Representa o visível de uma entrada de legenda do gráfico.|
|[ChartLegendEntryCollection](/javascript/api/excel/excel.chartlegendentrycollection)|[getCount()](/javascript/api/excel/excel.chartlegendentrycollection#getcount--)|Retorna o número de legendEntry da coleção.|
||[getItemAt(index: number)](/javascript/api/excel/excel.chartlegendentrycollection#getitemat-index-)|Retorna legendEntry no índice fornecido.|
||[items](/javascript/api/excel/excel.chartlegendentrycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartLegendEntryCollectionLoadOptions](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions)|[$all](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#$all)||
||[visible](/javascript/api/excel/excel.chartlegendentrycollectionloadoptions#visible)|Para cada ITEM na coleção: representa o visível de uma entrada de legenda de gráfico.|
|[ChartLegendEntryData](/javascript/api/excel/excel.chartlegendentrydata)|[visible](/javascript/api/excel/excel.chartlegendentrydata#visible)|Representa o visível de uma entrada de legenda do gráfico.|
|[ChartLegendEntryLoadOptions](/javascript/api/excel/excel.chartlegendentryloadoptions)|[$all](/javascript/api/excel/excel.chartlegendentryloadoptions#$all)||
||[visible](/javascript/api/excel/excel.chartlegendentryloadoptions#visible)|Representa o visível de uma entrada de legenda do gráfico.|
|[ChartLegendEntryUpdateData](/javascript/api/excel/excel.chartlegendentryupdatedata)|[visible](/javascript/api/excel/excel.chartlegendentryupdatedata#visible)|Representa o visível de uma entrada de legenda do gráfico.|
|[ChartLegendLoadOptions](/javascript/api/excel/excel.chartlegendloadoptions)|[height](/javascript/api/excel/excel.chartlegendloadoptions#height)|Representa a altura, em pontos, da legenda no gráfico. NULL se a legenda não estiver visível.|
||[left](/javascript/api/excel/excel.chartlegendloadoptions#left)|Representa a esquerda, em pontos, de uma legenda de gráfico. NULL se a legenda não estiver visível.|
||[Ocultar sombra](/javascript/api/excel/excel.chartlegendloadoptions#showshadow)|Representa se a legenda tem uma sombra no gráfico.|
||[top](/javascript/api/excel/excel.chartlegendloadoptions#top)|Representa o início de uma legenda do gráfico.|
||[width](/javascript/api/excel/excel.chartlegendloadoptions#width)|Representa a largura, em pontos, da legenda no gráfico. NULL se a legenda não estiver visível.|
|[ChartLegendUpdateData](/javascript/api/excel/excel.chartlegendupdatedata)|[height](/javascript/api/excel/excel.chartlegendupdatedata#height)|Representa a altura, em pontos, da legenda no gráfico. NULL se a legenda não estiver visível.|
||[left](/javascript/api/excel/excel.chartlegendupdatedata#left)|Representa a esquerda, em pontos, de uma legenda de gráfico. NULL se a legenda não estiver visível.|
||[Ocultar sombra](/javascript/api/excel/excel.chartlegendupdatedata#showshadow)|Representa se a legenda tem uma sombra no gráfico.|
||[top](/javascript/api/excel/excel.chartlegendupdatedata#top)|Representa o início de uma legenda do gráfico.|
||[width](/javascript/api/excel/excel.chartlegendupdatedata#width)|Representa a largura, em pontos, da legenda no gráfico. NULL se a legenda não estiver visível.|
|[ChartLineFormat](/javascript/api/excel/excel.chartlineformat)|[lineStyle](/javascript/api/excel/excel.chartlineformat#linestyle)|Representa o estilo da linha. Consulte Excel. ChartLineStyle para obter detalhes.|
||[weight](/javascript/api/excel/excel.chartlineformat#weight)|Representa a espessura da linha, em pontos.|
|[ChartLineFormatData](/javascript/api/excel/excel.chartlineformatdata)|[lineStyle](/javascript/api/excel/excel.chartlineformatdata#linestyle)|Representa o estilo da linha. Consulte Excel. ChartLineStyle para obter detalhes.|
||[weight](/javascript/api/excel/excel.chartlineformatdata#weight)|Representa a espessura da linha, em pontos.|
|[ChartLineFormatLoadOptions](/javascript/api/excel/excel.chartlineformatloadoptions)|[lineStyle](/javascript/api/excel/excel.chartlineformatloadoptions#linestyle)|Representa o estilo da linha. Consulte Excel. ChartLineStyle para obter detalhes.|
||[weight](/javascript/api/excel/excel.chartlineformatloadoptions#weight)|Representa a espessura da linha, em pontos.|
|[ChartLineFormatUpdateData](/javascript/api/excel/excel.chartlineformatupdatedata)|[lineStyle](/javascript/api/excel/excel.chartlineformatupdatedata#linestyle)|Representa o estilo da linha. Consulte Excel. ChartLineStyle para obter detalhes.|
||[weight](/javascript/api/excel/excel.chartlineformatupdatedata#weight)|Representa a espessura da linha, em pontos.|
|[ChartLoadOptions](/javascript/api/excel/excel.chartloadoptions)|[chartType](/javascript/api/excel/excel.chartloadoptions#charttype)|Representa o tipo de gráfico. Confira Excel. ChartType para obter detalhes.|
||[id](/javascript/api/excel/excel.chartloadoptions#id)|Id exclusiva do gráfico. Somente leitura.|
||[showAllFieldButtons](/javascript/api/excel/excel.chartloadoptions#showallfieldbuttons)|Representa se deseja exibir todos os botões de campo em um Gráfico Dinâmico.|
|[ChartPoint](/javascript/api/excel/excel.chartpoint)|[hasDataLabel](/javascript/api/excel/excel.chartpoint#hasdatalabel)|Indica se um ponto de dados tem um rótulo de dados. Não aplicável para gráficos de superfície.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpoint#markerbackgroundcolor)|Representação do código de cor HTML da cor de fundo do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpoint#markerforegroundcolor)|Representação do código de cor HTML da cor de primeiro plano do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|
||[markerSize](/javascript/api/excel/excel.chartpoint#markersize)|Representa o tamanho do marcador do ponto de dados.|
||[markerStyle](/javascript/api/excel/excel.chartpoint#markerstyle)|Representa estilo do marcador de um ponto de dados do gráfico. Consulte Excel. ChartMarkerStyle para obter detalhes.|
||[dataLabel](/javascript/api/excel/excel.chartpoint#datalabel)|Retorna o rótulo de dados de um ponto de gráfico. Somente leitura.|
|[ChartPointData](/javascript/api/excel/excel.chartpointdata)|[dataLabel](/javascript/api/excel/excel.chartpointdata#datalabel)|Retorna o rótulo de dados de um ponto de gráfico. Somente leitura.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointdata#hasdatalabel)|Indica se um ponto de dados tem um rótulo de dados. Não aplicável para gráficos de superfície.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointdata#markerbackgroundcolor)|Representação do código de cor HTML da cor de fundo do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointdata#markerforegroundcolor)|Representação do código de cor HTML da cor de primeiro plano do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|
||[markerSize](/javascript/api/excel/excel.chartpointdata#markersize)|Representa o tamanho do marcador do ponto de dados.|
||[markerStyle](/javascript/api/excel/excel.chartpointdata#markerstyle)|Representa estilo do marcador de um ponto de dados do gráfico. Consulte Excel. ChartMarkerStyle para obter detalhes.|
|[ChartPointFormat](/javascript/api/excel/excel.chartpointformat)|[Borderô](/javascript/api/excel/excel.chartpointformat#border)|Representa o formato da borda de um ponto de dados do gráfico, que inclui informações de cor, estilo e peso. Somente leitura.|
|[ChartPointFormatData](/javascript/api/excel/excel.chartpointformatdata)|[Borderô](/javascript/api/excel/excel.chartpointformatdata#border)|Representa o formato da borda de um ponto de dados do gráfico, que inclui informações de cor, estilo e peso. Somente leitura.|
|[ChartPointFormatLoadOptions](/javascript/api/excel/excel.chartpointformatloadoptions)|[Borderô](/javascript/api/excel/excel.chartpointformatloadoptions#border)|Representa o formato da borda de um ponto de dados do gráfico, que inclui informações de cor, estilo e peso.|
|[ChartPointFormatUpdateData](/javascript/api/excel/excel.chartpointformatupdatedata)|[Borderô](/javascript/api/excel/excel.chartpointformatupdatedata#border)|Representa o formato da borda de um ponto de dados do gráfico, que inclui informações de cor, estilo e peso.|
|[ChartPointLoadOptions](/javascript/api/excel/excel.chartpointloadoptions)|[dataLabel](/javascript/api/excel/excel.chartpointloadoptions#datalabel)|Retorna o rótulo de dados de um ponto de gráfico.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointloadoptions#hasdatalabel)|Indica se um ponto de dados tem um rótulo de dados. Não aplicável para gráficos de superfície.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointloadoptions#markerbackgroundcolor)|Representação do código de cor HTML da cor de fundo do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointloadoptions#markerforegroundcolor)|Representação do código de cor HTML da cor de primeiro plano do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|
||[markerSize](/javascript/api/excel/excel.chartpointloadoptions#markersize)|Representa o tamanho do marcador do ponto de dados.|
||[markerStyle](/javascript/api/excel/excel.chartpointloadoptions#markerstyle)|Representa estilo do marcador de um ponto de dados do gráfico. Consulte Excel. ChartMarkerStyle para obter detalhes.|
|[ChartPointUpdateData](/javascript/api/excel/excel.chartpointupdatedata)|[dataLabel](/javascript/api/excel/excel.chartpointupdatedata#datalabel)|Retorna o rótulo de dados de um ponto de gráfico.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointupdatedata#hasdatalabel)|Indica se um ponto de dados tem um rótulo de dados. Não aplicável para gráficos de superfície.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointupdatedata#markerbackgroundcolor)|Representação do código de cor HTML da cor de fundo do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointupdatedata#markerforegroundcolor)|Representação do código de cor HTML da cor de primeiro plano do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|
||[markerSize](/javascript/api/excel/excel.chartpointupdatedata#markersize)|Representa o tamanho do marcador do ponto de dados.|
||[markerStyle](/javascript/api/excel/excel.chartpointupdatedata#markerstyle)|Representa estilo do marcador de um ponto de dados do gráfico. Consulte Excel. ChartMarkerStyle para obter detalhes.|
|[ChartPointsCollectionLoadOptions](/javascript/api/excel/excel.chartpointscollectionloadoptions)|[dataLabel](/javascript/api/excel/excel.chartpointscollectionloadoptions#datalabel)|Para cada ITEM na coleção: retorna o rótulo de dados de um ponto de gráfico.|
||[hasDataLabel](/javascript/api/excel/excel.chartpointscollectionloadoptions#hasdatalabel)|Para cada ITEM na coleção: indica se um ponto de dados tem um rótulo de dados. Não aplicável para gráficos de superfície.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerbackgroundcolor)|Para cada ITEM na coleção: representação do código de cor HTML da cor de plano de fundo do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|
||[markerForegroundColor](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerforegroundcolor)|Para cada ITEM na coleção: representação do código de cor HTML da cor de primeiro plano do marcador do ponto de dados. Por exemplo #FF0000 representa vermelho.|
||[markerSize](/javascript/api/excel/excel.chartpointscollectionloadoptions#markersize)|Para cada ITEM na coleção: representa o tamanho do marcador do ponto de dados.|
||[markerStyle](/javascript/api/excel/excel.chartpointscollectionloadoptions#markerstyle)|Para cada ITEM na coleção: representa o estilo de marcador de um ponto de dados do gráfico. Consulte Excel. ChartMarkerStyle para obter detalhes.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[chartType](/javascript/api/excel/excel.chartseries#charttype)|Representa o tipo de gráfico de uma série. Confira Excel. ChartType para obter detalhes.|
||[delete()](/javascript/api/excel/excel.chartseries#delete--)|Exclui a série de gráfico.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseries#doughnutholesize)|Representa o tamanho do furo de rosca de uma série de gráficos.  Válida apenas em gráficos de rosca e doughnutExploded.|
||[último](/javascript/api/excel/excel.chartseries#filtered)|Valor booliano representando se a série é filtrada ou não. Não aplicável para gráficos de superfície.|
||[gapWidth](/javascript/api/excel/excel.chartseries#gapwidth)|Representa a largura do espaçamento de uma série de gráfico.  Válida apenas sobre gráficos de barras e colunas, bem como|
||[hasDataLabels](/javascript/api/excel/excel.chartseries#hasdatalabels)|Valor booliano representando se a série tem rótulos de dados ou não.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseries#markerbackgroundcolor)|Representa a cor de fundo dos marcadores de uma série de gráficos.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseries#markerforegroundcolor)|Representa cor de primeiro plano dos marcadores de uma série de gráfico.|
||[markerSize](/javascript/api/excel/excel.chartseries#markersize)|Representa o tamanho do marcador de uma série de gráfico.|
||[markerStyle](/javascript/api/excel/excel.chartseries#markerstyle)|Representa o estilo do marcador de uma série de gráfico. Consulte Excel. ChartMarkerStyle para obter detalhes.|
||[plotOrder](/javascript/api/excel/excel.chartseries#plotorder)|Representa a ordem de plotagem de uma série de gráficos dentro do grupo de gráfico.|
||[Trendlines](/javascript/api/excel/excel.chartseries#trendlines)|Representa uma coleção de todas as linha de tendência da série. Somente leitura.|
||[setBubbleSizes (sourceData: Range)](/javascript/api/excel/excel.chartseries#setbubblesizes-sourcedata-)|Definir tamanhos das bolhas para uma série de gráfico. Funciona apenas para gráficos de bolhas.|
||[SetValues (sourceData: Range)](/javascript/api/excel/excel.chartseries#setvalues-sourcedata-)|Definir valores de uma série de gráficos. Para gráfico de dispersão, isso significa valores do eixo Y.|
||[setXAxisValues (sourceData: Range)](/javascript/api/excel/excel.chartseries#setxaxisvalues-sourcedata-)|Definir valores do eixo X para uma série de gráficos. Funciona apenas para gráficos de dispersão.|
||[Ocultar sombra](/javascript/api/excel/excel.chartseries#showshadow)|Valor booliano que representa se a série tem uma sombra ou não.|
||[suave](/javascript/api/excel/excel.chartseries#smooth)|Valor booliano representando se a série é suave ou não. Só se aplica a gráficos de linhas e de dispersão.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[Add (Name?: String, index?: Number)](/javascript/api/excel/excel.chartseriescollection#add-name--index-)|Adiciona uma nova série para o conjunto. A nova série adicionada não fica visível até que Set Values/x Axis Values/tamanho da bolha (dependendo do tipo de gráfico).|
|[ChartSeriesCollectionLoadOptions](/javascript/api/excel/excel.chartseriescollectionloadoptions)|[chartType](/javascript/api/excel/excel.chartseriescollectionloadoptions#charttype)|Para cada ITEM na coleção: representa o tipo de gráfico de uma série. Confira Excel. ChartType para obter detalhes.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#doughnutholesize)|Para cada ITEM na coleção: representa o tamanho do buraco de rosca de uma série de gráficos.  Válida apenas em gráficos de rosca e doughnutExploded.|
||[último](/javascript/api/excel/excel.chartseriescollectionloadoptions#filtered)|Para cada ITEM na coleção: valor booliano que representa se a série é filtrada ou não. Não aplicável para gráficos de superfície.|
||[gapWidth](/javascript/api/excel/excel.chartseriescollectionloadoptions#gapwidth)|Para cada ITEM na coleção: representa a largura do intervalo de uma série de gráficos.  Válida apenas sobre gráficos de barras e colunas, bem como|
||[hasDataLabels](/javascript/api/excel/excel.chartseriescollectionloadoptions#hasdatalabels)|Para cada ITEM na coleção: valor booliano que representa se a série tem rótulos de dados ou não.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerbackgroundcolor)|Para cada ITEM na coleção: representa a cor de plano de fundo de marcadores de uma série de gráficos.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerforegroundcolor)|Para cada ITEM na coleção: representa a cor de primeiro plano de marcadores de uma série de gráficos.|
||[markerSize](/javascript/api/excel/excel.chartseriescollectionloadoptions#markersize)|Para cada ITEM na coleção: representa o tamanho do marcador de uma série de gráficos.|
||[markerStyle](/javascript/api/excel/excel.chartseriescollectionloadoptions#markerstyle)|Para cada ITEM na coleção: representa o estilo de marcador de uma série de gráficos. Consulte Excel. ChartMarkerStyle para obter detalhes.|
||[plotOrder](/javascript/api/excel/excel.chartseriescollectionloadoptions#plotorder)|Para cada ITEM na coleção: representa a ordem de plotagem de uma série de gráfico dentro do grupo de gráficos.|
||[Ocultar sombra](/javascript/api/excel/excel.chartseriescollectionloadoptions#showshadow)|Para cada ITEM na coleção: valor booliano que representa se a série tem uma sombra ou não.|
||[suave](/javascript/api/excel/excel.chartseriescollectionloadoptions#smooth)|Para cada ITEM na coleção: valor booliano que representa se a série é suave ou não. Só se aplica a gráficos de linhas e de dispersão.|
|[ChartSeriesData](/javascript/api/excel/excel.chartseriesdata)|[chartType](/javascript/api/excel/excel.chartseriesdata#charttype)|Representa o tipo de gráfico de uma série. Confira Excel. ChartType para obter detalhes.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesdata#doughnutholesize)|Representa o tamanho do furo de rosca de uma série de gráficos.  Válida apenas em gráficos de rosca e doughnutExploded.|
||[último](/javascript/api/excel/excel.chartseriesdata#filtered)|Valor booliano representando se a série é filtrada ou não. Não aplicável para gráficos de superfície.|
||[gapWidth](/javascript/api/excel/excel.chartseriesdata#gapwidth)|Representa a largura do espaçamento de uma série de gráfico.  Válida apenas sobre gráficos de barras e colunas, bem como|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesdata#hasdatalabels)|Valor booliano representando se a série tem rótulos de dados ou não.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesdata#markerbackgroundcolor)|Representa a cor de fundo dos marcadores de uma série de gráficos.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesdata#markerforegroundcolor)|Representa cor de primeiro plano dos marcadores de uma série de gráfico.|
||[markerSize](/javascript/api/excel/excel.chartseriesdata#markersize)|Representa o tamanho do marcador de uma série de gráfico.|
||[markerStyle](/javascript/api/excel/excel.chartseriesdata#markerstyle)|Representa o estilo do marcador de uma série de gráfico. Consulte Excel. ChartMarkerStyle para obter detalhes.|
||[plotOrder](/javascript/api/excel/excel.chartseriesdata#plotorder)|Representa a ordem de plotagem de uma série de gráficos dentro do grupo de gráfico.|
||[Ocultar sombra](/javascript/api/excel/excel.chartseriesdata#showshadow)|Valor booliano que representa se a série tem uma sombra ou não.|
||[suave](/javascript/api/excel/excel.chartseriesdata#smooth)|Valor booliano representando se a série é suave ou não. Só se aplica a gráficos de linhas e de dispersão.|
||[Trendlines](/javascript/api/excel/excel.chartseriesdata#trendlines)|Representa uma coleção de todas as linha de tendência da série. Somente leitura.|
|[ChartSeriesLoadOptions](/javascript/api/excel/excel.chartseriesloadoptions)|[chartType](/javascript/api/excel/excel.chartseriesloadoptions#charttype)|Representa o tipo de gráfico de uma série. Confira Excel. ChartType para obter detalhes.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesloadoptions#doughnutholesize)|Representa o tamanho do furo de rosca de uma série de gráficos.  Válida apenas em gráficos de rosca e doughnutExploded.|
||[último](/javascript/api/excel/excel.chartseriesloadoptions#filtered)|Valor booliano representando se a série é filtrada ou não. Não aplicável para gráficos de superfície.|
||[gapWidth](/javascript/api/excel/excel.chartseriesloadoptions#gapwidth)|Representa a largura do espaçamento de uma série de gráfico.  Válida apenas sobre gráficos de barras e colunas, bem como|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesloadoptions#hasdatalabels)|Valor booliano representando se a série tem rótulos de dados ou não.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesloadoptions#markerbackgroundcolor)|Representa a cor de fundo dos marcadores de uma série de gráficos.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesloadoptions#markerforegroundcolor)|Representa cor de primeiro plano dos marcadores de uma série de gráfico.|
||[markerSize](/javascript/api/excel/excel.chartseriesloadoptions#markersize)|Representa o tamanho do marcador de uma série de gráfico.|
||[markerStyle](/javascript/api/excel/excel.chartseriesloadoptions#markerstyle)|Representa o estilo do marcador de uma série de gráfico. Consulte Excel. ChartMarkerStyle para obter detalhes.|
||[plotOrder](/javascript/api/excel/excel.chartseriesloadoptions#plotorder)|Representa a ordem de plotagem de uma série de gráficos dentro do grupo de gráfico.|
||[Ocultar sombra](/javascript/api/excel/excel.chartseriesloadoptions#showshadow)|Valor booliano que representa se a série tem uma sombra ou não.|
||[suave](/javascript/api/excel/excel.chartseriesloadoptions#smooth)|Valor booliano representando se a série é suave ou não. Só se aplica a gráficos de linhas e de dispersão.|
|[ChartSeriesUpdateData](/javascript/api/excel/excel.chartseriesupdatedata)|[chartType](/javascript/api/excel/excel.chartseriesupdatedata#charttype)|Representa o tipo de gráfico de uma série. Confira Excel. ChartType para obter detalhes.|
||[doughnutHoleSize](/javascript/api/excel/excel.chartseriesupdatedata#doughnutholesize)|Representa o tamanho do furo de rosca de uma série de gráficos.  Válida apenas em gráficos de rosca e doughnutExploded.|
||[último](/javascript/api/excel/excel.chartseriesupdatedata#filtered)|Valor booliano representando se a série é filtrada ou não. Não aplicável para gráficos de superfície.|
||[gapWidth](/javascript/api/excel/excel.chartseriesupdatedata#gapwidth)|Representa a largura do espaçamento de uma série de gráfico.  Válida apenas sobre gráficos de barras e colunas, bem como|
||[hasDataLabels](/javascript/api/excel/excel.chartseriesupdatedata#hasdatalabels)|Valor booliano representando se a série tem rótulos de dados ou não.|
||[markerBackgroundColor](/javascript/api/excel/excel.chartseriesupdatedata#markerbackgroundcolor)|Representa a cor de fundo dos marcadores de uma série de gráficos.|
||[markerForegroundColor](/javascript/api/excel/excel.chartseriesupdatedata#markerforegroundcolor)|Representa cor de primeiro plano dos marcadores de uma série de gráfico.|
||[markerSize](/javascript/api/excel/excel.chartseriesupdatedata#markersize)|Representa o tamanho do marcador de uma série de gráfico.|
||[markerStyle](/javascript/api/excel/excel.chartseriesupdatedata#markerstyle)|Representa o estilo do marcador de uma série de gráfico. Consulte Excel. ChartMarkerStyle para obter detalhes.|
||[plotOrder](/javascript/api/excel/excel.chartseriesupdatedata#plotorder)|Representa a ordem de plotagem de uma série de gráficos dentro do grupo de gráfico.|
||[Ocultar sombra](/javascript/api/excel/excel.chartseriesupdatedata#showshadow)|Valor booliano que representa se a série tem uma sombra ou não.|
||[suave](/javascript/api/excel/excel.chartseriesupdatedata#smooth)|Valor booliano representando se a série é suave ou não. Só se aplica a gráficos de linhas e de dispersão.|
|[ChartTitle](/javascript/api/excel/excel.charttitle)|[getSubstring (início: número, comprimento: número)](/javascript/api/excel/excel.charttitle#getsubstring-start--length-)|Obter a subcadeia de caracteres de um título de gráfico. A quebra de linha ' \n ' também conta um caractere.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitle#horizontalalignment)|Representa o alinhamento horizontal para título do gráfico.|
||[left](/javascript/api/excel/excel.charttitle#left)|Representa a distância, em pontos, da borda esquerda do título do gráfico até a borda esquerda da área do gráfico. NULL se o título do gráfico não estiver visível.|
||[position](/javascript/api/excel/excel.charttitle#position)|Representa a posição de título do gráfico. Consulte Excel. ChartTitlePosition para obter detalhes.|
||[height](/javascript/api/excel/excel.charttitle#height)|Representa a altura, em pontos, do título do gráfico. NULL se o título do gráfico não estiver visível. Somente leitura.|
||[width](/javascript/api/excel/excel.charttitle#width)|Retorna a largura em pontos do título do gráfico. NULL se o título do gráfico não estiver visível. Somente leitura.|
||[setformula (fórmula: cadeia de caracteres)](/javascript/api/excel/excel.charttitle#setformula-formula-)|Define um valor de cadeia de caracteres que representa a fórmula do título do eixo do gráfico usando a notação no estilo A1.|
||[Ocultar sombra](/javascript/api/excel/excel.charttitle#showshadow)|Representa um valor booliano que determina se o título do gráfico tiver uma sombra.|
||[textOrientation](/javascript/api/excel/excel.charttitle#textorientation)|Representa a orientação de texto do título do gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.charttitle#top)|Representa a distância em pontos, da borda superior do título do gráfico a parte superior da área do gráfico. NULL se o título do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.charttitle#verticalalignment)|Representa o alinhamento vertical do título do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
|[ChartTitleData](/javascript/api/excel/excel.charttitledata)|[height](/javascript/api/excel/excel.charttitledata#height)|Representa a altura, em pontos, do título do gráfico. NULL se o título do gráfico não estiver visível. Somente leitura.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitledata#horizontalalignment)|Representa o alinhamento horizontal para título do gráfico.|
||[left](/javascript/api/excel/excel.charttitledata#left)|Representa a distância, em pontos, da borda esquerda do título do gráfico até a borda esquerda da área do gráfico. NULL se o título do gráfico não estiver visível.|
||[position](/javascript/api/excel/excel.charttitledata#position)|Representa a posição de título do gráfico. Consulte Excel. ChartTitlePosition para obter detalhes.|
||[Ocultar sombra](/javascript/api/excel/excel.charttitledata#showshadow)|Representa um valor booliano que determina se o título do gráfico tiver uma sombra.|
||[textOrientation](/javascript/api/excel/excel.charttitledata#textorientation)|Representa a orientação de texto do título do gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.charttitledata#top)|Representa a distância em pontos, da borda superior do título do gráfico a parte superior da área do gráfico. NULL se o título do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.charttitledata#verticalalignment)|Representa o alinhamento vertical do título do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
||[width](/javascript/api/excel/excel.charttitledata#width)|Retorna a largura em pontos do título do gráfico. NULL se o título do gráfico não estiver visível. Somente leitura.|
|[ChartTitleFormat](/javascript/api/excel/excel.charttitleformat)|[Borderô](/javascript/api/excel/excel.charttitleformat#border)|Representa o formato da borda do título do gráfico, que inclui cores, LineStyle e Weight. Somente leitura.|
|[ChartTitleFormatData](/javascript/api/excel/excel.charttitleformatdata)|[Borderô](/javascript/api/excel/excel.charttitleformatdata#border)|Representa o formato da borda do título do gráfico, que inclui cores, LineStyle e Weight. Somente leitura.|
|[ChartTitleFormatLoadOptions](/javascript/api/excel/excel.charttitleformatloadoptions)|[Borderô](/javascript/api/excel/excel.charttitleformatloadoptions#border)|Representa o formato da borda do título do gráfico, que inclui cores, LineStyle e Weight.|
|[ChartTitleFormatUpdateData](/javascript/api/excel/excel.charttitleformatupdatedata)|[Borderô](/javascript/api/excel/excel.charttitleformatupdatedata#border)|Representa o formato da borda do título do gráfico, que inclui cores, LineStyle e Weight.|
|[ChartTitleLoadOptions](/javascript/api/excel/excel.charttitleloadoptions)|[height](/javascript/api/excel/excel.charttitleloadoptions#height)|Representa a altura, em pontos, do título do gráfico. NULL se o título do gráfico não estiver visível. Somente leitura.|
||[horizontalAlignment](/javascript/api/excel/excel.charttitleloadoptions#horizontalalignment)|Representa o alinhamento horizontal para título do gráfico.|
||[left](/javascript/api/excel/excel.charttitleloadoptions#left)|Representa a distância, em pontos, da borda esquerda do título do gráfico até a borda esquerda da área do gráfico. NULL se o título do gráfico não estiver visível.|
||[position](/javascript/api/excel/excel.charttitleloadoptions#position)|Representa a posição de título do gráfico. Consulte Excel. ChartTitlePosition para obter detalhes.|
||[Ocultar sombra](/javascript/api/excel/excel.charttitleloadoptions#showshadow)|Representa um valor booliano que determina se o título do gráfico tiver uma sombra.|
||[textOrientation](/javascript/api/excel/excel.charttitleloadoptions#textorientation)|Representa a orientação de texto do título do gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.charttitleloadoptions#top)|Representa a distância em pontos, da borda superior do título do gráfico a parte superior da área do gráfico. NULL se o título do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.charttitleloadoptions#verticalalignment)|Representa o alinhamento vertical do título do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
||[width](/javascript/api/excel/excel.charttitleloadoptions#width)|Retorna a largura em pontos do título do gráfico. NULL se o título do gráfico não estiver visível. Somente leitura.|
|[ChartTitleUpdateData](/javascript/api/excel/excel.charttitleupdatedata)|[horizontalAlignment](/javascript/api/excel/excel.charttitleupdatedata#horizontalalignment)|Representa o alinhamento horizontal para título do gráfico.|
||[left](/javascript/api/excel/excel.charttitleupdatedata#left)|Representa a distância, em pontos, da borda esquerda do título do gráfico até a borda esquerda da área do gráfico. NULL se o título do gráfico não estiver visível.|
||[position](/javascript/api/excel/excel.charttitleupdatedata#position)|Representa a posição de título do gráfico. Consulte Excel. ChartTitlePosition para obter detalhes.|
||[Ocultar sombra](/javascript/api/excel/excel.charttitleupdatedata#showshadow)|Representa um valor booliano que determina se o título do gráfico tiver uma sombra.|
||[textOrientation](/javascript/api/excel/excel.charttitleupdatedata#textorientation)|Representa a orientação de texto do título do gráfico. O valor deve ser um número inteiro de -90 a 90 ou 180 para texto orientado verticalmente.|
||[top](/javascript/api/excel/excel.charttitleupdatedata#top)|Representa a distância em pontos, da borda superior do título do gráfico a parte superior da área do gráfico. NULL se o título do gráfico não estiver visível.|
||[verticalAlignment](/javascript/api/excel/excel.charttitleupdatedata#verticalalignment)|Representa o alinhamento vertical do título do gráfico. Consulte Excel. ChartTextVerticalAlignment para obter detalhes.|
|[ChartTrendline](/javascript/api/excel/excel.charttrendline)|[delete()](/javascript/api/excel/excel.charttrendline#delete--)|Deleta o objeto Trendline.|
||[detecta](/javascript/api/excel/excel.charttrendline#intercept)|Representa o valor de intercepção da linha de tendência. Pode ser definido como um valor numérico ou uma cadeia de caracteres vazia (para valores automáticos de eixo). O valor retornado sempre é um número.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendline#movingaverageperiod)|Representa o período de uma tendência de gráfico. Aplicável somente para tendência com tipo MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendline#name)|Representa o nome da linha de tendência. Pode ser definido como um valor de sequência ou pode ser definido como valor nulo para representar valores automáticos. O valor retornado sempre é uma cadeia de caracteres.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendline#polynomialorder)|Representa a ordem de uma tendência de gráfico. Aplicável somente para tendência com tipo polinomial.|
||[format](/javascript/api/excel/excel.charttrendline#format)|Representa a formatação de uma linha de tendência do gráfico.|
||[Set (Propriedades: Excel. ChartTrendline)](/javascript/api/excel/excel.charttrendline#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartTrendlineUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.charttrendline#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[tipo](/javascript/api/excel/excel.charttrendline#type)|Representa o tipo da linha de tendência de um gráfico.|
|[ChartTrendlineCollection](/javascript/api/excel/excel.charttrendlinecollection)|[Add (tipo?: "linear" \| "exponencial \| " "logarítmica" \| "MovingAverage" \| "polinomial" \| "Power")](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Adiciona uma nova linha de tendência ao conjunto de linha de tendência.|
||[Add (tipo?: Excel. ChartTrendlineType)](/javascript/api/excel/excel.charttrendlinecollection#add-type-)|Adiciona uma nova linha de tendência ao conjunto de linha de tendência.|
||[getCount()](/javascript/api/excel/excel.charttrendlinecollection#getcount--)|Retorna o número de linha de tendência na coleção.|
||[getItem(index: number)](/javascript/api/excel/excel.charttrendlinecollection#getitem-index-)|Obtém o objeto da linha de tendência por índice, que é a ordem de inserção na matriz de itens.|
||[items](/javascript/api/excel/excel.charttrendlinecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[ChartTrendlineCollectionLoadOptions](/javascript/api/excel/excel.charttrendlinecollectionloadoptions)|[$all](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#format)|Para cada ITEM na coleção: representa a formatação de uma tendência de gráfico.|
||[detecta](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#intercept)|Para cada ITEM na coleção: representa o valor de interseção da tendência. Pode ser definido como um valor numérico ou uma cadeia de caracteres vazia (para valores automáticos de eixo). O valor retornado sempre é um número.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#movingaverageperiod)|Para cada ITEM na coleção: representa o período de uma tendência de gráfico. Aplicável somente para tendência com tipo MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#name)|Para cada ITEM na coleção: representa o nome da tendência. Pode ser definido como um valor de sequência ou pode ser definido como valor nulo para representar valores automáticos. O valor retornado sempre é uma cadeia de caracteres.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#polynomialorder)|Para cada ITEM na coleção: representa a ordem de uma tendência de gráfico. Aplicável somente para tendência com tipo polinomial.|
||[tipo](/javascript/api/excel/excel.charttrendlinecollectionloadoptions#type)|Para cada ITEM na coleção: representa o tipo de uma tendência de gráfico.|
|[ChartTrendlineData](/javascript/api/excel/excel.charttrendlinedata)|[format](/javascript/api/excel/excel.charttrendlinedata#format)|Representa a formatação de uma linha de tendência do gráfico.|
||[detecta](/javascript/api/excel/excel.charttrendlinedata#intercept)|Representa o valor de intercepção da linha de tendência. Pode ser definido como um valor numérico ou uma cadeia de caracteres vazia (para valores automáticos de eixo). O valor retornado sempre é um número.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlinedata#movingaverageperiod)|Representa o período de uma tendência de gráfico. Aplicável somente para tendência com tipo MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendlinedata#name)|Representa o nome da linha de tendência. Pode ser definido como um valor de sequência ou pode ser definido como valor nulo para representar valores automáticos. O valor retornado sempre é uma cadeia de caracteres.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlinedata#polynomialorder)|Representa a ordem de uma tendência de gráfico. Aplicável somente para tendência com tipo polinomial.|
||[tipo](/javascript/api/excel/excel.charttrendlinedata#type)|Representa o tipo da linha de tendência de um gráfico.|
|[ChartTrendlineFormat](/javascript/api/excel/excel.charttrendlineformat)|[line](/javascript/api/excel/excel.charttrendlineformat#line)|Representa a formatação de linha do gráfico. Somente leitura.|
||[Set (Propriedades: Excel. ChartTrendlineFormat)](/javascript/api/excel/excel.charttrendlineformat#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. ChartTrendlineFormatUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.charttrendlineformat#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
|[ChartTrendlineFormatData](/javascript/api/excel/excel.charttrendlineformatdata)|[line](/javascript/api/excel/excel.charttrendlineformatdata#line)|Representa a formatação de linha do gráfico. Somente leitura.|
|[ChartTrendlineFormatLoadOptions](/javascript/api/excel/excel.charttrendlineformatloadoptions)|[$all](/javascript/api/excel/excel.charttrendlineformatloadoptions#$all)||
||[line](/javascript/api/excel/excel.charttrendlineformatloadoptions#line)|Representa a formatação de linha do gráfico.|
|[ChartTrendlineFormatUpdateData](/javascript/api/excel/excel.charttrendlineformatupdatedata)|[line](/javascript/api/excel/excel.charttrendlineformatupdatedata#line)|Representa a formatação de linha do gráfico.|
|[ChartTrendlineLoadOptions](/javascript/api/excel/excel.charttrendlineloadoptions)|[$all](/javascript/api/excel/excel.charttrendlineloadoptions#$all)||
||[format](/javascript/api/excel/excel.charttrendlineloadoptions#format)|Representa a formatação de uma linha de tendência do gráfico.|
||[detecta](/javascript/api/excel/excel.charttrendlineloadoptions#intercept)|Representa o valor de intercepção da linha de tendência. Pode ser definido como um valor numérico ou uma cadeia de caracteres vazia (para valores automáticos de eixo). O valor retornado sempre é um número.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlineloadoptions#movingaverageperiod)|Representa o período de uma tendência de gráfico. Aplicável somente para tendência com tipo MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendlineloadoptions#name)|Representa o nome da linha de tendência. Pode ser definido como um valor de sequência ou pode ser definido como valor nulo para representar valores automáticos. O valor retornado sempre é uma cadeia de caracteres.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlineloadoptions#polynomialorder)|Representa a ordem de uma tendência de gráfico. Aplicável somente para tendência com tipo polinomial.|
||[tipo](/javascript/api/excel/excel.charttrendlineloadoptions#type)|Representa o tipo da linha de tendência de um gráfico.|
|[ChartTrendlineUpdateData](/javascript/api/excel/excel.charttrendlineupdatedata)|[format](/javascript/api/excel/excel.charttrendlineupdatedata#format)|Representa a formatação de uma linha de tendência do gráfico.|
||[detecta](/javascript/api/excel/excel.charttrendlineupdatedata#intercept)|Representa o valor de intercepção da linha de tendência. Pode ser definido como um valor numérico ou uma cadeia de caracteres vazia (para valores automáticos de eixo). O valor retornado sempre é um número.|
||[movingAveragePeriod](/javascript/api/excel/excel.charttrendlineupdatedata#movingaverageperiod)|Representa o período de uma tendência de gráfico. Aplicável somente para tendência com tipo MovingAverage.|
||[name](/javascript/api/excel/excel.charttrendlineupdatedata#name)|Representa o nome da linha de tendência. Pode ser definido como um valor de sequência ou pode ser definido como valor nulo para representar valores automáticos. O valor retornado sempre é uma cadeia de caracteres.|
||[polynomialOrder](/javascript/api/excel/excel.charttrendlineupdatedata#polynomialorder)|Representa a ordem de uma tendência de gráfico. Aplicável somente para tendência com tipo polinomial.|
||[tipo](/javascript/api/excel/excel.charttrendlineupdatedata#type)|Representa o tipo da linha de tendência de um gráfico.|
|[ChartUpdateData](/javascript/api/excel/excel.chartupdatedata)|[chartType](/javascript/api/excel/excel.chartupdatedata#charttype)|Representa o tipo de gráfico. Confira Excel. ChartType para obter detalhes.|
||[showAllFieldButtons](/javascript/api/excel/excel.chartupdatedata#showallfieldbuttons)|Representa se deseja exibir todos os botões de campo em um Gráfico Dinâmico.|
|[CustomProperty](/javascript/api/excel/excel.customproperty)|[delete()](/javascript/api/excel/excel.customproperty#delete--)|Exclui a propriedade personalizada.|
||[key](/javascript/api/excel/excel.customproperty#key)|Obtém a chave da propriedade personalizada. Somente leitura.|
||[tipo](/javascript/api/excel/excel.customproperty#type)|Obtém o tipo de valor da propriedade personalizada. Somente leitura.|
||[Set (Propriedades: Excel. CustomProperty)](/javascript/api/excel/excel.customproperty#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. CustomPropertyUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.customproperty#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[value](/javascript/api/excel/excel.customproperty#value)|Obtém ou define o valor da propriedade personalizada.|
|[CustomPropertyCollection](/javascript/api/excel/excel.custompropertycollection)|[Add (Key: String, value: any)](/javascript/api/excel/excel.custompropertycollection#add-key--value-)|Cria uma nova propriedade personalizada ou define uma existente.|
||[deleteAll ()](/javascript/api/excel/excel.custompropertycollection#deleteall--)|Exclui todas as propriedades personalizadas nesta coleção.|
||[getCount()](/javascript/api/excel/excel.custompropertycollection#getcount--)|Obtém a contagem das propriedades personalizadas.|
||[getItem(key: string)](/javascript/api/excel/excel.custompropertycollection#getitem-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Lança se a propriedade personalizada não existe.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.custompropertycollection#getitemornullobject-key-)|Obtém um objeto de propriedade personalizada por sua chave, que diferencia maiúsculas de minúsculas. Retorna um objeto NULL se a propriedade personalizada não existir.|
||[items](/javascript/api/excel/excel.custompropertycollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[CustomPropertyCollectionLoadOptions](/javascript/api/excel/excel.custompropertycollectionloadoptions)|[$all](/javascript/api/excel/excel.custompropertycollectionloadoptions#$all)||
||[key](/javascript/api/excel/excel.custompropertycollectionloadoptions#key)|Para cada ITEM na coleção: Obtém a chave da propriedade personalizada. Somente leitura.|
||[tipo](/javascript/api/excel/excel.custompropertycollectionloadoptions#type)|Para cada ITEM na coleção: Obtém o tipo de valor da propriedade personalizada. Somente leitura.|
||[value](/javascript/api/excel/excel.custompropertycollectionloadoptions#value)|Para cada ITEM na coleção: Obtém ou define o valor da propriedade personalizada.|
|[CustomPropertyData](/javascript/api/excel/excel.custompropertydata)|[key](/javascript/api/excel/excel.custompropertydata#key)|Obtém a chave da propriedade personalizada. Somente leitura.|
||[tipo](/javascript/api/excel/excel.custompropertydata#type)|Obtém o tipo de valor da propriedade personalizada. Somente leitura.|
||[value](/javascript/api/excel/excel.custompropertydata#value)|Obtém ou define o valor da propriedade personalizada.|
|[CustomPropertyLoadOptions](/javascript/api/excel/excel.custompropertyloadoptions)|[$all](/javascript/api/excel/excel.custompropertyloadoptions#$all)||
||[key](/javascript/api/excel/excel.custompropertyloadoptions#key)|Obtém a chave da propriedade personalizada. Somente leitura.|
||[tipo](/javascript/api/excel/excel.custompropertyloadoptions#type)|Obtém o tipo de valor da propriedade personalizada. Somente leitura.|
||[value](/javascript/api/excel/excel.custompropertyloadoptions#value)|Obtém ou define o valor da propriedade personalizada.|
|[CustomPropertyUpdateData](/javascript/api/excel/excel.custompropertyupdatedata)|[value](/javascript/api/excel/excel.custompropertyupdatedata#value)|Obtém ou define o valor da propriedade personalizada.|
|[Dataconnectioncollection](/javascript/api/excel/excel.dataconnectioncollection)|[refreshAll ()](/javascript/api/excel/excel.dataconnectioncollection#refreshall--)|Atualiza todas as conexões de dados da coleção.|
|[DocumentProperties](/javascript/api/excel/excel.documentproperties)|[autor](/javascript/api/excel/excel.documentproperties#author)|Obtém ou define o autor da pasta de trabalho.|
||[Categorias](/javascript/api/excel/excel.documentproperties#category)|Obtém ou define a categoria da pasta de trabalho.|
||[comments](/javascript/api/excel/excel.documentproperties#comments)|Obtém ou define os comentários da pasta de trabalho.|
||[company](/javascript/api/excel/excel.documentproperties#company)|Obtém ou define a empresa do documento.|
||[Palavras-chave](/javascript/api/excel/excel.documentproperties#keywords)|Obtém ou define as palavras-chave da pasta de trabalho.|
||[Gerenciador](/javascript/api/excel/excel.documentproperties#manager)|Obtém ou define o gerenciador da pasta de trabalho.|
||[creationDate](/javascript/api/excel/excel.documentproperties#creationdate)|Obtém a data de criação da pasta de trabalho. Somente leitura.|
||[cliente](/javascript/api/excel/excel.documentproperties#custom)|Obtém a coleção de propriedades personalizadas da pasta de trabalho. Somente leitura.|
||[lastAuthor](/javascript/api/excel/excel.documentproperties#lastauthor)|Obtém o último autor da pasta de trabalho. Somente leitura.|
||[revisionNumber](/javascript/api/excel/excel.documentproperties#revisionnumber)|Obtém o número de revisão da pasta de trabalho. Somente leitura.|
||[Set (Propriedades: Excel. DocumentProperties)](/javascript/api/excel/excel.documentproperties#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. DocumentPropertiesUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.documentproperties#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[subject](/javascript/api/excel/excel.documentproperties#subject)|Obtém ou define o assunto da pasta de trabalho.|
||[title](/javascript/api/excel/excel.documentproperties#title)|Obtém ou define o título da pasta de trabalho.|
|[DocumentPropertiesData](/javascript/api/excel/excel.documentpropertiesdata)|[autor](/javascript/api/excel/excel.documentpropertiesdata#author)|Obtém ou define o autor da pasta de trabalho.|
||[Categorias](/javascript/api/excel/excel.documentpropertiesdata#category)|Obtém ou define a categoria da pasta de trabalho.|
||[comments](/javascript/api/excel/excel.documentpropertiesdata#comments)|Obtém ou define os comentários da pasta de trabalho.|
||[company](/javascript/api/excel/excel.documentpropertiesdata#company)|Obtém ou define a empresa do documento.|
||[creationDate](/javascript/api/excel/excel.documentpropertiesdata#creationdate)|Obtém a data de criação da pasta de trabalho. Somente leitura.|
||[cliente](/javascript/api/excel/excel.documentpropertiesdata#custom)|Obtém a coleção de propriedades personalizadas da pasta de trabalho. Somente leitura.|
||[Palavras-chave](/javascript/api/excel/excel.documentpropertiesdata#keywords)|Obtém ou define as palavras-chave da pasta de trabalho.|
||[lastAuthor](/javascript/api/excel/excel.documentpropertiesdata#lastauthor)|Obtém o último autor da pasta de trabalho. Somente leitura.|
||[Gerenciador](/javascript/api/excel/excel.documentpropertiesdata#manager)|Obtém ou define o gerenciador da pasta de trabalho.|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesdata#revisionnumber)|Obtém o número de revisão da pasta de trabalho. Somente leitura.|
||[subject](/javascript/api/excel/excel.documentpropertiesdata#subject)|Obtém ou define o assunto da pasta de trabalho.|
||[title](/javascript/api/excel/excel.documentpropertiesdata#title)|Obtém ou define o título da pasta de trabalho.|
|[DocumentPropertiesLoadOptions](/javascript/api/excel/excel.documentpropertiesloadoptions)|[$all](/javascript/api/excel/excel.documentpropertiesloadoptions#$all)||
||[autor](/javascript/api/excel/excel.documentpropertiesloadoptions#author)|Obtém ou define o autor da pasta de trabalho.|
||[Categorias](/javascript/api/excel/excel.documentpropertiesloadoptions#category)|Obtém ou define a categoria da pasta de trabalho.|
||[comments](/javascript/api/excel/excel.documentpropertiesloadoptions#comments)|Obtém ou define os comentários da pasta de trabalho.|
||[company](/javascript/api/excel/excel.documentpropertiesloadoptions#company)|Obtém ou define a empresa do documento.|
||[creationDate](/javascript/api/excel/excel.documentpropertiesloadoptions#creationdate)|Obtém a data de criação da pasta de trabalho. Somente leitura.|
||[Palavras-chave](/javascript/api/excel/excel.documentpropertiesloadoptions#keywords)|Obtém ou define as palavras-chave da pasta de trabalho.|
||[lastAuthor](/javascript/api/excel/excel.documentpropertiesloadoptions#lastauthor)|Obtém o último autor da pasta de trabalho. Somente leitura.|
||[Gerenciador](/javascript/api/excel/excel.documentpropertiesloadoptions#manager)|Obtém ou define o gerenciador da pasta de trabalho.|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesloadoptions#revisionnumber)|Obtém o número de revisão da pasta de trabalho. Somente leitura.|
||[subject](/javascript/api/excel/excel.documentpropertiesloadoptions#subject)|Obtém ou define o assunto da pasta de trabalho.|
||[title](/javascript/api/excel/excel.documentpropertiesloadoptions#title)|Obtém ou define o título da pasta de trabalho.|
|[DocumentPropertiesUpdateData](/javascript/api/excel/excel.documentpropertiesupdatedata)|[autor](/javascript/api/excel/excel.documentpropertiesupdatedata#author)|Obtém ou define o autor da pasta de trabalho.|
||[Categorias](/javascript/api/excel/excel.documentpropertiesupdatedata#category)|Obtém ou define a categoria da pasta de trabalho.|
||[comments](/javascript/api/excel/excel.documentpropertiesupdatedata#comments)|Obtém ou define os comentários da pasta de trabalho.|
||[company](/javascript/api/excel/excel.documentpropertiesupdatedata#company)|Obtém ou define a empresa do documento.|
||[Palavras-chave](/javascript/api/excel/excel.documentpropertiesupdatedata#keywords)|Obtém ou define as palavras-chave da pasta de trabalho.|
||[Gerenciador](/javascript/api/excel/excel.documentpropertiesupdatedata#manager)|Obtém ou define o gerenciador da pasta de trabalho.|
||[revisionNumber](/javascript/api/excel/excel.documentpropertiesupdatedata#revisionnumber)|Obtém o número de revisão da pasta de trabalho. Somente leitura.|
||[subject](/javascript/api/excel/excel.documentpropertiesupdatedata#subject)|Obtém ou define o assunto da pasta de trabalho.|
||[title](/javascript/api/excel/excel.documentpropertiesupdatedata#title)|Obtém ou define o título da pasta de trabalho.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[formula](/javascript/api/excel/excel.nameditem#formula)|Obtém ou define a fórmula do item nomeado.  A fórmula sempre começa com um sinal de "=".|
||[arrayValues](/javascript/api/excel/excel.nameditem#arrayvalues)|Retorna um objeto que contém valores e tipos do item nomeado. Somente leitura.|
|[NamedItemArrayValues](/javascript/api/excel/excel.nameditemarrayvalues)|[types](/javascript/api/excel/excel.nameditemarrayvalues#types)|Representa os tipos de cada item na matriz de itens nomeados|
||[values](/javascript/api/excel/excel.nameditemarrayvalues#values)|Representa os valores de cada item na matriz de itens nomeados.|
|[NamedItemArrayValuesData](/javascript/api/excel/excel.nameditemarrayvaluesdata)|[types](/javascript/api/excel/excel.nameditemarrayvaluesdata#types)|Representa os tipos de cada item na matriz de itens nomeados|
||[values](/javascript/api/excel/excel.nameditemarrayvaluesdata#values)|Representa os valores de cada item na matriz de itens nomeados.|
|[NamedItemArrayValuesLoadOptions](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions)|[$all](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#$all)||
||[types](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#types)|Representa os tipos de cada item na matriz de itens nomeados|
||[values](/javascript/api/excel/excel.nameditemarrayvaluesloadoptions#values)|Representa os valores de cada item na matriz de itens nomeados.|
|[NamedItemCollectionLoadOptions](/javascript/api/excel/excel.nameditemcollectionloadoptions)|[arrayValues](/javascript/api/excel/excel.nameditemcollectionloadoptions#arrayvalues)|Para cada ITEM na coleção: retorna um objeto que contém valores e tipos do item nomeado.|
||[formula](/javascript/api/excel/excel.nameditemcollectionloadoptions#formula)|Para cada ITEM na coleção: Obtém ou define a fórmula do item nomeado.  A fórmula sempre começa com um sinal de "=".|
|[NamedItemData](/javascript/api/excel/excel.nameditemdata)|[arrayValues](/javascript/api/excel/excel.nameditemdata#arrayvalues)|Retorna um objeto que contém valores e tipos do item nomeado. Somente leitura.|
||[formula](/javascript/api/excel/excel.nameditemdata#formula)|Obtém ou define a fórmula do item nomeado.  A fórmula sempre começa com um sinal de "=".|
|[NamedItemLoadOptions](/javascript/api/excel/excel.nameditemloadoptions)|[arrayValues](/javascript/api/excel/excel.nameditemloadoptions#arrayvalues)|Retorna um objeto que contém valores e tipos do item nomeado.|
||[formula](/javascript/api/excel/excel.nameditemloadoptions#formula)|Obtém ou define a fórmula do item nomeado.  A fórmula sempre começa com um sinal de "=".|
|[NamedItemUpdateData](/javascript/api/excel/excel.nameditemupdatedata)|[formula](/javascript/api/excel/excel.nameditemupdatedata#formula)|Obtém ou define a fórmula do item nomeado.  A fórmula sempre começa com um sinal de "=".|
|[Range](/javascript/api/excel/excel.range)|[getAbsoluteResizedRange (numRows: Number, numColumns: Number)](/javascript/api/excel/excel.range#getabsoluteresizedrange-numrows--numcolumns-)|Obtém um objeto Range com a mesma célula superior esquerda do objeto Range atual, mas com os números especificados de linhas e colunas.|
||[GetImage ()](/javascript/api/excel/excel.range#getimage--)|Renderiza o intervalo como uma imagem png codificada em base64.|
||[getSurroundingRegion()](/javascript/api/excel/excel.range#getsurroundingregion--)|Retorna um objeto Range que representa a região circundante da célula superior esquerda nesse intervalo. Uma região ao redor é um intervalo limitado por qualquer combinação de linhas e colunas em branco em relação a esse intervalo.|
||[hiperlink](/javascript/api/excel/excel.range#hyperlink)|Representa o hiperlink do intervalo atual.|
||[numberFormatLocal](/javascript/api/excel/excel.range#numberformatlocal)|Representa o código de formato numérico do Excel para o intervalo fornecido como uma cadeia de caracteres no idioma do usuário.|
||[isEntireColumn](/javascript/api/excel/excel.range#isentirecolumn)|Representa se o intervalo atual está em uma coluna inteira. Somente leitura.|
||[isEntireRow](/javascript/api/excel/excel.range#isentirerow)|Representa se o intervalo atual está em uma linha inteira. Somente leitura.|
||[Cartão ()](/javascript/api/excel/excel.range#showcard--)|Exibe o cartão para uma célula ativa se ele tiver um conteúdo valioso.|
||[style](/javascript/api/excel/excel.range#style)|Representa o estilo de intervalo atual.|
|[RangeData](/javascript/api/excel/excel.rangedata)|[hiperlink](/javascript/api/excel/excel.rangedata#hyperlink)|Representa o hiperlink do intervalo atual.|
||[isEntireColumn](/javascript/api/excel/excel.rangedata#isentirecolumn)|Representa se o intervalo atual está em uma coluna inteira. Somente leitura.|
||[isEntireRow](/javascript/api/excel/excel.rangedata#isentirerow)|Representa se o intervalo atual está em uma linha inteira. Somente leitura.|
||[numberFormatLocal](/javascript/api/excel/excel.rangedata#numberformatlocal)|Representa o código de formato numérico do Excel para o intervalo fornecido como uma cadeia de caracteres no idioma do usuário.|
||[style](/javascript/api/excel/excel.rangedata#style)|Representa o estilo de intervalo atual.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[textOrientation](/javascript/api/excel/excel.rangeformat#textorientation)|Obtém ou define a orientação de texto de todas as células no intervalo.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformat#usestandardheight)|Determina se a altura da linha do objeto Range é igual a altura padrão da planilha.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformat#usestandardwidth)|Indica se a largura da coluna do objeto Range é igual à largura padrão da planilha.|
|[RangeFormatData](/javascript/api/excel/excel.rangeformatdata)|[textOrientation](/javascript/api/excel/excel.rangeformatdata#textorientation)|Obtém ou define a orientação de texto de todas as células no intervalo.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatdata#usestandardheight)|Determina se a altura da linha do objeto Range é igual a altura padrão da planilha.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatdata#usestandardwidth)|Indica se a largura da coluna do objeto Range é igual à largura padrão da planilha.|
|[RangeFormatLoadOptions](/javascript/api/excel/excel.rangeformatloadoptions)|[textOrientation](/javascript/api/excel/excel.rangeformatloadoptions#textorientation)|Obtém ou define a orientação de texto de todas as células no intervalo.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatloadoptions#usestandardheight)|Determina se a altura da linha do objeto Range é igual a altura padrão da planilha.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatloadoptions#usestandardwidth)|Indica se a largura da coluna do objeto Range é igual à largura padrão da planilha.|
|[RangeFormatUpdateData](/javascript/api/excel/excel.rangeformatupdatedata)|[textOrientation](/javascript/api/excel/excel.rangeformatupdatedata#textorientation)|Obtém ou define a orientação de texto de todas as células no intervalo.|
||[useStandardHeight](/javascript/api/excel/excel.rangeformatupdatedata#usestandardheight)|Determina se a altura da linha do objeto Range é igual a altura padrão da planilha.|
||[useStandardWidth](/javascript/api/excel/excel.rangeformatupdatedata#usestandardwidth)|Indica se a largura da coluna do objeto Range é igual à largura padrão da planilha.|
|[RangeHyperlink](/javascript/api/excel/excel.rangehyperlink)|[address](/javascript/api/excel/excel.rangehyperlink#address)|Representa o destino da url do hiperlink.|
||[documentReference](/javascript/api/excel/excel.rangehyperlink#documentreference)|Representa o destino de referência de documento para o hiperlink.|
||[Dica](/javascript/api/excel/excel.rangehyperlink#screentip)|Representa a cadeia exibida ao passar o mouse sobre o hiperlink.|
||[textToDisplay](/javascript/api/excel/excel.rangehyperlink#texttodisplay)|Representa a cadeia de caracteres exibida na parte superior esquerda da maioria das células no intervalo.|
|[RangeLoadOptions](/javascript/api/excel/excel.rangeloadoptions)|[hiperlink](/javascript/api/excel/excel.rangeloadoptions#hyperlink)|Representa o hiperlink do intervalo atual.|
||[isEntireColumn](/javascript/api/excel/excel.rangeloadoptions#isentirecolumn)|Representa se o intervalo atual está em uma coluna inteira. Somente leitura.|
||[isEntireRow](/javascript/api/excel/excel.rangeloadoptions#isentirerow)|Representa se o intervalo atual está em uma linha inteira. Somente leitura.|
||[numberFormatLocal](/javascript/api/excel/excel.rangeloadoptions#numberformatlocal)|Representa o código de formato numérico do Excel para o intervalo fornecido como uma cadeia de caracteres no idioma do usuário.|
||[style](/javascript/api/excel/excel.rangeloadoptions#style)|Representa o estilo de intervalo atual.|
|[RangeUpdateData](/javascript/api/excel/excel.rangeupdatedata)|[hiperlink](/javascript/api/excel/excel.rangeupdatedata#hyperlink)|Representa o hiperlink do intervalo atual.|
||[numberFormatLocal](/javascript/api/excel/excel.rangeupdatedata#numberformatlocal)|Representa o código de formato numérico do Excel para o intervalo fornecido como uma cadeia de caracteres no idioma do usuário.|
||[style](/javascript/api/excel/excel.rangeupdatedata#style)|Representa o estilo de intervalo atual.|
|[Estilo](/javascript/api/excel/excel.style)|[delete()](/javascript/api/excel/excel.style#delete--)|Exclui este estilo.|
||[formulaHidden](/javascript/api/excel/excel.style#formulahidden)|Indica se a fórmula ficará oculta quando a planilha estiver protegida.|
||[horizontalAlignment](/javascript/api/excel/excel.style#horizontalalignment)|Representa o alinhamento horizontal para o estilo. Consulte Excel. HorizontalAlignment para obter detalhes.|
||[includeAlignment](/javascript/api/excel/excel.style#includealignment)|Indica se o estilo incluem as propriedades AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, e TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.style#includeborder)|Indica se o estilo inclui as propriedades de borda Color, ColorIndex, LineStyle e Weight.|
||[includeFont](/javascript/api/excel/excel.style#includefont)|Indica se o estilo inclui as propriedades de fonte Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript e Underline.|
||[includeNumber](/javascript/api/excel/excel.style#includenumber)|Indica se o estilo inclui a propriedade NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.style#includepatterns)|Indica se o estilo inclui as propriedades internas Color, ColorIndex, InvertIfNegative, Pattern, PatternColor e PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.style#includeprotection)|Indica se o estilo incluirá as propriedades de proteção FormulaHidden e Locked.|
||[indentLevel](/javascript/api/excel/excel.style#indentlevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo do estilo.|
||[bloqueado](/javascript/api/excel/excel.style#locked)|Indica se o objeto é bloqueado quando a planilha está protegida.|
||[numberFormat](/javascript/api/excel/excel.style#numberformat)|O código de formatação de formato de número para o estilo.|
||[numberFormatLocal](/javascript/api/excel/excel.style#numberformatlocal)|O código de formato localizado do formato numérico para o estilo.|
||[readingOrder](/javascript/api/excel/excel.style#readingorder)|A ordem de leitura para o estilo.|
||[Borders](/javascript/api/excel/excel.style#borders)|Uma coleção Border de quatro objetos Border que representam o estilo das quatro bordas.|
||[Interna](/javascript/api/excel/excel.style#builtin)|Indica se o estilo é um estilo interno.|
||[fill](/javascript/api/excel/excel.style#fill)|O preenchimento do estilo.|
||[font](/javascript/api/excel/excel.style#font)|Objeto de fonte que representa a fonte do estilo.|
||[name](/javascript/api/excel/excel.style#name)|O nome do estilo.|
||[Set (Propriedades: Excel. Style)](/javascript/api/excel/excel.style#set-properties-)|Define várias propriedades no objeto ao mesmo tempo, com base em um objeto carregado existente.|
||[Set (Propriedades: interfaces. StyleUpdateData, opções?: OfficeExtension. UpdateOptions)](/javascript/api/excel/excel.style#set-properties--options-)|Define várias propriedades de um objeto ao mesmo tempo. Você pode passar um objeto simples com as propriedades apropriadas ou outro objeto API do mesmo tipo.|
||[shrinkToFit](/javascript/api/excel/excel.style#shrinktofit)|Indica se o texto é automaticamente reduzido para caber na largura da coluna disponível.|
||[verticalAlignment](/javascript/api/excel/excel.style#verticalalignment)|Representa o alinhamento vertical do estilo. Consulte Excel. VerticalAlignment para obter detalhes.|
||[wrapText](/javascript/api/excel/excel.style#wraptext)|Indica se o Microsoft Excel quebra automaticamente a linha de texto no objeto.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[add(name: string)](/javascript/api/excel/excel.stylecollection#add-name-)|Adiciona um novo estilo para o conjunto.|
||[getItem(name: string)](/javascript/api/excel/excel.stylecollection#getitem-name-)|Obtém um estilo por nome.|
||[items](/javascript/api/excel/excel.stylecollection#items)|Obtém os itens filhos carregados nesta coleção.|
|[StyleCollectionLoadOptions](/javascript/api/excel/excel.stylecollectionloadoptions)|[$all](/javascript/api/excel/excel.stylecollectionloadoptions#$all)||
||[Borders](/javascript/api/excel/excel.stylecollectionloadoptions#borders)|Para cada ITEM na coleção: uma coleção Border de quatro objetos Border que representam o estilo das quatro bordas.|
||[Interna](/javascript/api/excel/excel.stylecollectionloadoptions#builtin)|Para cada ITEM na coleção: indica se o estilo é um estilo interno.|
||[fill](/javascript/api/excel/excel.stylecollectionloadoptions#fill)|Para cada ITEM na coleção: o preenchimento do estilo.|
||[font](/javascript/api/excel/excel.stylecollectionloadoptions#font)|Para cada ITEM na coleção: um objeto Font que representa a fonte do estilo.|
||[formulaHidden](/javascript/api/excel/excel.stylecollectionloadoptions#formulahidden)|Para cada ITEM na coleção: indica se a fórmula ficará oculta quando a planilha estiver protegida.|
||[horizontalAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#horizontalalignment)|Para cada ITEM na coleção: representa o alinhamento horizontal para o estilo. Consulte Excel. HorizontalAlignment para obter detalhes.|
||[includeAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#includealignment)|Para cada ITEM na coleção: indica se o estilo inclui as propriedades autoindent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel e TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.stylecollectionloadoptions#includeborder)|Para cada ITEM na coleção: indica se o estilo inclui as propriedades de borda Color, ColorIndex, LineStyle e Weight.|
||[includeFont](/javascript/api/excel/excel.stylecollectionloadoptions#includefont)|Para cada ITEM na coleção: indica se o estilo inclui as propriedades de fonte de plano de fundo, negrito, cor, ColorIndex, FontStyle, itálico, nome, tamanho, tachado, subscrito, sobrescrito e sublinhado.|
||[includeNumber](/javascript/api/excel/excel.stylecollectionloadoptions#includenumber)|Para cada ITEM na coleção: indica se o estilo inclui a propriedade NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.stylecollectionloadoptions#includepatterns)|Para cada ITEM na coleção: indica se o estilo inclui as propriedades interiores Color, ColorIndex, InvertIfNegative, Pattern, PatternColor e PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.stylecollectionloadoptions#includeprotection)|Para cada ITEM na coleção: indica se o estilo inclui as propriedades de proteção FormulaHidden e Locked.|
||[indentLevel](/javascript/api/excel/excel.stylecollectionloadoptions#indentlevel)|Para cada ITEM na coleção: um inteiro de 0 a 250 que indica o nível de recuo para o estilo.|
||[bloqueado](/javascript/api/excel/excel.stylecollectionloadoptions#locked)|Para cada ITEM na coleção: indica se o objeto está bloqueado quando a planilha está protegida.|
||[name](/javascript/api/excel/excel.stylecollectionloadoptions#name)|Para cada ITEM na coleção: o nome do estilo.|
||[numberFormat](/javascript/api/excel/excel.stylecollectionloadoptions#numberformat)|Para cada ITEM na coleção: o código de formatação do formato de número para o estilo.|
||[numberFormatLocal](/javascript/api/excel/excel.stylecollectionloadoptions#numberformatlocal)|Para cada ITEM na coleção: o código de formato localizado do formato de número para o estilo.|
||[readingOrder](/javascript/api/excel/excel.stylecollectionloadoptions#readingorder)|Para cada ITEM na coleção: o sentido de leitura para o estilo.|
||[shrinkToFit](/javascript/api/excel/excel.stylecollectionloadoptions#shrinktofit)|Para cada ITEM na coleção: indica se o texto é automaticamente reduzido para se ajustar à largura de coluna disponível.|
||[verticalAlignment](/javascript/api/excel/excel.stylecollectionloadoptions#verticalalignment)|Para cada ITEM na coleção: representa o alinhamento vertical do estilo. Consulte Excel. VerticalAlignment para obter detalhes.|
||[wrapText](/javascript/api/excel/excel.stylecollectionloadoptions#wraptext)|Para cada ITEM na coleção: indica se o Microsoft Excel quebra o texto no objeto.|
|[StyleData](/javascript/api/excel/excel.styledata)|[Borders](/javascript/api/excel/excel.styledata#borders)|Uma coleção Border de quatro objetos Border que representam o estilo das quatro bordas.|
||[Interna](/javascript/api/excel/excel.styledata#builtin)|Indica se o estilo é um estilo interno.|
||[fill](/javascript/api/excel/excel.styledata#fill)|O preenchimento do estilo.|
||[font](/javascript/api/excel/excel.styledata#font)|Objeto de fonte que representa a fonte do estilo.|
||[formulaHidden](/javascript/api/excel/excel.styledata#formulahidden)|Indica se a fórmula ficará oculta quando a planilha estiver protegida.|
||[horizontalAlignment](/javascript/api/excel/excel.styledata#horizontalalignment)|Representa o alinhamento horizontal para o estilo. Consulte Excel. HorizontalAlignment para obter detalhes.|
||[includeAlignment](/javascript/api/excel/excel.styledata#includealignment)|Indica se o estilo incluem as propriedades AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, e TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.styledata#includeborder)|Indica se o estilo inclui as propriedades de borda Color, ColorIndex, LineStyle e Weight.|
||[includeFont](/javascript/api/excel/excel.styledata#includefont)|Indica se o estilo inclui as propriedades de fonte Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript e Underline.|
||[includeNumber](/javascript/api/excel/excel.styledata#includenumber)|Indica se o estilo inclui a propriedade NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.styledata#includepatterns)|Indica se o estilo inclui as propriedades internas Color, ColorIndex, InvertIfNegative, Pattern, PatternColor e PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.styledata#includeprotection)|Indica se o estilo incluirá as propriedades de proteção FormulaHidden e Locked.|
||[indentLevel](/javascript/api/excel/excel.styledata#indentlevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo do estilo.|
||[bloqueado](/javascript/api/excel/excel.styledata#locked)|Indica se o objeto é bloqueado quando a planilha está protegida.|
||[name](/javascript/api/excel/excel.styledata#name)|O nome do estilo.|
||[numberFormat](/javascript/api/excel/excel.styledata#numberformat)|O código de formatação de formato de número para o estilo.|
||[numberFormatLocal](/javascript/api/excel/excel.styledata#numberformatlocal)|O código de formato localizado do formato numérico para o estilo.|
||[readingOrder](/javascript/api/excel/excel.styledata#readingorder)|A ordem de leitura para o estilo.|
||[shrinkToFit](/javascript/api/excel/excel.styledata#shrinktofit)|Indica se o texto é automaticamente reduzido para caber na largura da coluna disponível.|
||[verticalAlignment](/javascript/api/excel/excel.styledata#verticalalignment)|Representa o alinhamento vertical do estilo. Consulte Excel. VerticalAlignment para obter detalhes.|
||[wrapText](/javascript/api/excel/excel.styledata#wraptext)|Indica se o Microsoft Excel quebra automaticamente a linha de texto no objeto.|
|[StyleLoadOptions](/javascript/api/excel/excel.styleloadoptions)|[$all](/javascript/api/excel/excel.styleloadoptions#$all)||
||[Borders](/javascript/api/excel/excel.styleloadoptions#borders)|Uma coleção Border de quatro objetos Border que representam o estilo das quatro bordas.|
||[Interna](/javascript/api/excel/excel.styleloadoptions#builtin)|Indica se o estilo é um estilo interno.|
||[fill](/javascript/api/excel/excel.styleloadoptions#fill)|O preenchimento do estilo.|
||[font](/javascript/api/excel/excel.styleloadoptions#font)|Objeto de fonte que representa a fonte do estilo.|
||[formulaHidden](/javascript/api/excel/excel.styleloadoptions#formulahidden)|Indica se a fórmula ficará oculta quando a planilha estiver protegida.|
||[horizontalAlignment](/javascript/api/excel/excel.styleloadoptions#horizontalalignment)|Representa o alinhamento horizontal para o estilo. Consulte Excel. HorizontalAlignment para obter detalhes.|
||[includeAlignment](/javascript/api/excel/excel.styleloadoptions#includealignment)|Indica se o estilo incluem as propriedades AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, e TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.styleloadoptions#includeborder)|Indica se o estilo inclui as propriedades de borda Color, ColorIndex, LineStyle e Weight.|
||[includeFont](/javascript/api/excel/excel.styleloadoptions#includefont)|Indica se o estilo inclui as propriedades de fonte Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript e Underline.|
||[includeNumber](/javascript/api/excel/excel.styleloadoptions#includenumber)|Indica se o estilo inclui a propriedade NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.styleloadoptions#includepatterns)|Indica se o estilo inclui as propriedades internas Color, ColorIndex, InvertIfNegative, Pattern, PatternColor e PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.styleloadoptions#includeprotection)|Indica se o estilo incluirá as propriedades de proteção FormulaHidden e Locked.|
||[indentLevel](/javascript/api/excel/excel.styleloadoptions#indentlevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo do estilo.|
||[bloqueado](/javascript/api/excel/excel.styleloadoptions#locked)|Indica se o objeto é bloqueado quando a planilha está protegida.|
||[name](/javascript/api/excel/excel.styleloadoptions#name)|O nome do estilo.|
||[numberFormat](/javascript/api/excel/excel.styleloadoptions#numberformat)|O código de formatação de formato de número para o estilo.|
||[numberFormatLocal](/javascript/api/excel/excel.styleloadoptions#numberformatlocal)|O código de formato localizado do formato numérico para o estilo.|
||[readingOrder](/javascript/api/excel/excel.styleloadoptions#readingorder)|A ordem de leitura para o estilo.|
||[shrinkToFit](/javascript/api/excel/excel.styleloadoptions#shrinktofit)|Indica se o texto é automaticamente reduzido para caber na largura da coluna disponível.|
||[verticalAlignment](/javascript/api/excel/excel.styleloadoptions#verticalalignment)|Representa o alinhamento vertical do estilo. Consulte Excel. VerticalAlignment para obter detalhes.|
||[wrapText](/javascript/api/excel/excel.styleloadoptions#wraptext)|Indica se o Microsoft Excel quebra automaticamente a linha de texto no objeto.|
|[StyleUpdateData](/javascript/api/excel/excel.styleupdatedata)|[Borders](/javascript/api/excel/excel.styleupdatedata#borders)|Uma coleção Border de quatro objetos Border que representam o estilo das quatro bordas.|
||[fill](/javascript/api/excel/excel.styleupdatedata#fill)|O preenchimento do estilo.|
||[font](/javascript/api/excel/excel.styleupdatedata#font)|Objeto de fonte que representa a fonte do estilo.|
||[formulaHidden](/javascript/api/excel/excel.styleupdatedata#formulahidden)|Indica se a fórmula ficará oculta quando a planilha estiver protegida.|
||[horizontalAlignment](/javascript/api/excel/excel.styleupdatedata#horizontalalignment)|Representa o alinhamento horizontal para o estilo. Consulte Excel. HorizontalAlignment para obter detalhes.|
||[includeAlignment](/javascript/api/excel/excel.styleupdatedata#includealignment)|Indica se o estilo incluem as propriedades AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, e TextOrientation.|
||[includeBorder](/javascript/api/excel/excel.styleupdatedata#includeborder)|Indica se o estilo inclui as propriedades de borda Color, ColorIndex, LineStyle e Weight.|
||[includeFont](/javascript/api/excel/excel.styleupdatedata#includefont)|Indica se o estilo inclui as propriedades de fonte Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript e Underline.|
||[includeNumber](/javascript/api/excel/excel.styleupdatedata#includenumber)|Indica se o estilo inclui a propriedade NumberFormat.|
||[includePatterns](/javascript/api/excel/excel.styleupdatedata#includepatterns)|Indica se o estilo inclui as propriedades internas Color, ColorIndex, InvertIfNegative, Pattern, PatternColor e PatternColorIndex.|
||[includeProtection](/javascript/api/excel/excel.styleupdatedata#includeprotection)|Indica se o estilo incluirá as propriedades de proteção FormulaHidden e Locked.|
||[indentLevel](/javascript/api/excel/excel.styleupdatedata#indentlevel)|Um número inteiro entre 0 e 250 que indica o nível de recuo do estilo.|
||[bloqueado](/javascript/api/excel/excel.styleupdatedata#locked)|Indica se o objeto é bloqueado quando a planilha está protegida.|
||[numberFormat](/javascript/api/excel/excel.styleupdatedata#numberformat)|O código de formatação de formato de número para o estilo.|
||[numberFormatLocal](/javascript/api/excel/excel.styleupdatedata#numberformatlocal)|O código de formato localizado do formato numérico para o estilo.|
||[readingOrder](/javascript/api/excel/excel.styleupdatedata#readingorder)|A ordem de leitura para o estilo.|
||[shrinkToFit](/javascript/api/excel/excel.styleupdatedata#shrinktofit)|Indica se o texto é automaticamente reduzido para caber na largura da coluna disponível.|
||[verticalAlignment](/javascript/api/excel/excel.styleupdatedata#verticalalignment)|Representa o alinhamento vertical do estilo. Consulte Excel. VerticalAlignment para obter detalhes.|
||[wrapText](/javascript/api/excel/excel.styleupdatedata#wraptext)|Indica se o Microsoft Excel quebra automaticamente a linha de texto no objeto.|
|[Table](/javascript/api/excel/excel.table)|[onChanged](/javascript/api/excel/excel.table#onchanged)|Ocorre quando os dados nas células são alterados em uma tabela específica.|
||[onSelectionChanged](/javascript/api/excel/excel.table#onselectionchanged)|Ocorre quando a seleção é alterada em uma tabela específica.|
|[TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs)|[address](/javascript/api/excel/excel.tablechangedeventargs#address)|Obtém o endereço que representa a área alterada de uma tabela em uma planilha específica.|
||[changeType](/javascript/api/excel/excel.tablechangedeventargs#changetype)|Obtém o tipo de mudança que representa como o evento Changed é acionado. Confira Excel. datachangtype para obter detalhes.|
||[source](/javascript/api/excel/excel.tablechangedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tableId](/javascript/api/excel/excel.tablechangedeventargs#tableid)|Obtém o id da tabela na qual os dados foram alterados.|
||[tipo](/javascript/api/excel/excel.tablechangedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablechangedeventargs#worksheetid)|Obtém o id da planilha na qual os dados são alterados.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onChanged](/javascript/api/excel/excel.tablecollection#onchanged)|Ocorre quando os dados são alterados em qualquer tabela em uma pasta de trabalho ou em uma planilha.|
|[TableSelectionChangedEventArgs](/javascript/api/excel/excel.tableselectionchangedeventargs)|[address](/javascript/api/excel/excel.tableselectionchangedeventargs#address)|Obtém o endereço do intervalo que representa a área selecionada da tabela em uma planilha específica.|
||[isInsideTable](/javascript/api/excel/excel.tableselectionchangedeventargs#isinsidetable)|Indica se a seleção está dentro de uma tabela, o endereço será inútil se IsInsideTable for falso.|
||[tableId](/javascript/api/excel/excel.tableselectionchangedeventargs#tableid)|Obtém o id da tabela na qual a seleção foi alterada.|
||[tipo](/javascript/api/excel/excel.tableselectionchangedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType. Somente leitura.|
||[worksheetId](/javascript/api/excel/excel.tableselectionchangedeventargs#worksheetid)|Obtém o id da planilha na qual a seleção foi alterada.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveCell()](/javascript/api/excel/excel.workbook#getactivecell--)|Obtém a célula ativa no momento da pasta de trabalho.|
||[dataConnections](/javascript/api/excel/excel.workbook#dataconnections)|Representa todas as conexões de dados na pasta de trabalho. Somente leitura.|
||[name](/javascript/api/excel/excel.workbook#name)|Obtém o nome da pasta de trabalho. Somente leitura.|
||[properties](/javascript/api/excel/excel.workbook#properties)|Obtém as propriedades da pasta de trabalho. Somente leitura.|
||[protection](/javascript/api/excel/excel.workbook#protection)|Retorna o objeto de proteção de pasta de trabalho para uma pasta de trabalho. Somente leitura.|
||[estilos](/javascript/api/excel/excel.workbook#styles)|Representa uma coleção de estilos associados à pasta de trabalho. Somente leitura.|
|[WorkbookData](/javascript/api/excel/excel.workbookdata)|[name](/javascript/api/excel/excel.workbookdata#name)|Obtém o nome da pasta de trabalho. Somente leitura.|
||[properties](/javascript/api/excel/excel.workbookdata#properties)|Obtém as propriedades da pasta de trabalho. Somente leitura.|
||[protection](/javascript/api/excel/excel.workbookdata#protection)|Retorna o objeto de proteção de pasta de trabalho para uma pasta de trabalho. Somente leitura.|
||[estilos](/javascript/api/excel/excel.workbookdata#styles)|Representa uma coleção de estilos associados à pasta de trabalho. Somente leitura.|
|[WorkbookLoadOptions](/javascript/api/excel/excel.workbookloadoptions)|[name](/javascript/api/excel/excel.workbookloadoptions#name)|Obtém o nome da pasta de trabalho. Somente leitura.|
||[properties](/javascript/api/excel/excel.workbookloadoptions#properties)|Obtém as propriedades da pasta de trabalho.|
||[protection](/javascript/api/excel/excel.workbookloadoptions#protection)|Retorna o objeto de proteção de pasta de trabalho para uma pasta de trabalho.|
|[WorkbookProtection](/javascript/api/excel/excel.workbookprotection)|[proteger (senha?: cadeia de caracteres)](/javascript/api/excel/excel.workbookprotection#protect-password-)|Protege uma pasta de trabalho. Falhará se a pasta de trabalho estiver protegida.|
||[protegido](/javascript/api/excel/excel.workbookprotection#protected)|Indica se a pasta de trabalho está protegida. Somente Leitura.|
||[desproteger (senha?: cadeia de caracteres)](/javascript/api/excel/excel.workbookprotection#unprotect-password-)|Desprotege uma pasta de trabalho.|
|[WorkbookProtectionData](/javascript/api/excel/excel.workbookprotectiondata)|[protegido](/javascript/api/excel/excel.workbookprotectiondata#protected)|Indica se a pasta de trabalho está protegida. Somente Leitura.|
|[WorkbookProtectionLoadOptions](/javascript/api/excel/excel.workbookprotectionloadoptions)|[$all](/javascript/api/excel/excel.workbookprotectionloadoptions#$all)||
||[protegido](/javascript/api/excel/excel.workbookprotectionloadoptions#protected)|Indica se a pasta de trabalho está protegida. Somente Leitura.|
|[WorkbookUpdateData](/javascript/api/excel/excel.workbookupdatedata)|[properties](/javascript/api/excel/excel.workbookupdatedata#properties)|Obtém as propriedades da pasta de trabalho.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Copy (PositionType?: "nenhum" \| "antes" \| "após" \| "início \| " "final", relativeTo?: Excel. Worksheet)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Copia uma planilha e a coloca na posição especificada. Retorna à planilha copiada.|
||[Copy (PositionType?: Excel. WorksheetPositionType, relativeTo?: Excel. Worksheet)](/javascript/api/excel/excel.worksheet#copy-positiontype--relativeto-)|Copia uma planilha e a coloca na posição especificada. Retorna à planilha copiada.|
||[getRangeByIndexes (startRow: Number, startColumn: Number, rowCount: Number, columnCount: Number)](/javascript/api/excel/excel.worksheet#getrangebyindexes-startrow--startcolumn--rowcount--columncount-)|Obtém o objeto Range que começa em um determinado índice de linha e índice de coluna e que abrange um determinado número de linhas e colunas.|
||[freezePanes](/javascript/api/excel/excel.worksheet#freezepanes)|Obtém um objeto que pode ser usado para manipular painéis congelados na planilha. Somente leitura.|
||[onActivated](/javascript/api/excel/excel.worksheet#onactivated)|Ocorre quando a planilha é ativada.|
||[onChanged](/javascript/api/excel/excel.worksheet#onchanged)|Ocorre quando os dados são alterados em uma planilha específica.|
||[onDeactivated](/javascript/api/excel/excel.worksheet#ondeactivated)|Ocorre quando a planilha é desativada.|
||[onSelectionChanged](/javascript/api/excel/excel.worksheet#onselectionchanged)|Ocorre quando a seleção é alterada em uma planilha específica.|
||[standardHeight](/javascript/api/excel/excel.worksheet#standardheight)|Retorna a altura padrão de todas as linhas na planilha, em pontos. Somente leitura.|
||[standardWidth](/javascript/api/excel/excel.worksheet#standardwidth)|Retorna ou define a largura padrão de todas as colunas na planilha.|
||[tabColor](/javascript/api/excel/excel.worksheet#tabcolor)|Obtém ou define a cor da guia de planilha.|
|[WorksheetActivatedEventArgs](/javascript/api/excel/excel.worksheetactivatedeventargs)|[tipo](/javascript/api/excel/excel.worksheetactivatedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetactivatedeventargs#worksheetid)|Obtém o id da planilha que está ativada.|
|[WorksheetAddedEventArgs](/javascript/api/excel/excel.worksheetaddedeventargs)|[source](/javascript/api/excel/excel.worksheetaddedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.worksheetaddedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetaddedeventargs#worksheetid)|Obtém o id da planilha que é adicionada à pasta de trabalho.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[address](/javascript/api/excel/excel.worksheetchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[changeType](/javascript/api/excel/excel.worksheetchangedeventargs#changetype)|Obtém o tipo de mudança que representa como o evento Changed é acionado. Confira Excel. datachangtype para obter detalhes.|
||[source](/javascript/api/excel/excel.worksheetchangedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.worksheetchangedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetchangedeventargs#worksheetid)|Obtém o id da planilha na qual os dados são alterados.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onActivated](/javascript/api/excel/excel.worksheetcollection#onactivated)|Ocorre quando qualquer planilha na pasta de trabalho é ativada.|
||[onAdded](/javascript/api/excel/excel.worksheetcollection#onadded)|Ocorre quando uma nova planilha é adicionada à pasta de trabalho.|
||[onDeactivated](/javascript/api/excel/excel.worksheetcollection#ondeactivated)|Ocorre quando qualquer planilha na pasta de trabalho é desativada.|
||[onDeleted](/javascript/api/excel/excel.worksheetcollection#ondeleted)|Ocorre quando uma planilha é excluída da pasta de trabalho.|
|[WorksheetCollectionLoadOptions](/javascript/api/excel/excel.worksheetcollectionloadoptions)|[standardHeight](/javascript/api/excel/excel.worksheetcollectionloadoptions#standardheight)|Para cada ITEM na coleção: retorna a altura padrão (padrão) de todas as linhas da planilha, em pontos. Somente leitura.|
||[standardWidth](/javascript/api/excel/excel.worksheetcollectionloadoptions#standardwidth)|Para cada ITEM na coleção: Retorna ou define a largura padrão de todas as colunas da planilha.|
||[tabColor](/javascript/api/excel/excel.worksheetcollectionloadoptions#tabcolor)|Para cada ITEM na coleção: Obtém ou define a cor da guia de planilha.|
|[WorksheetData](/javascript/api/excel/excel.worksheetdata)|[standardHeight](/javascript/api/excel/excel.worksheetdata#standardheight)|Retorna a altura padrão de todas as linhas na planilha, em pontos. Somente leitura.|
||[standardWidth](/javascript/api/excel/excel.worksheetdata#standardwidth)|Retorna ou define a largura padrão de todas as colunas na planilha.|
||[tabColor](/javascript/api/excel/excel.worksheetdata#tabcolor)|Obtém ou define a cor da guia de planilha.|
|[WorksheetDeactivatedEventArgs](/javascript/api/excel/excel.worksheetdeactivatedeventargs)|[tipo](/javascript/api/excel/excel.worksheetdeactivatedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeactivatedeventargs#worksheetid)|Obtém o id da planilha que está desativada.|
|[WorksheetDeletedEventArgs](/javascript/api/excel/excel.worksheetdeletedeventargs)|[source](/javascript/api/excel/excel.worksheetdeletedeventargs#source)|Obtém a origem do evento. Para saber detalhes, confira Excel.EventSource.|
||[tipo](/javascript/api/excel/excel.worksheetdeletedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetdeletedeventargs#worksheetid)|Obtém o id do gráfico que é excluído da pasta de trabalho.|
|[WorksheetFreezePanes](/javascript/api/excel/excel.worksheetfreezepanes)|[freezeAt (frozenRange: cadeia \| de caracteres de intervalo)](/javascript/api/excel/excel.worksheetfreezepanes#freezeat-frozenrange-)|Define as células congeladas no modo de exibição da planilha ativa.|
||[freezeColumns (contagem?: número)](/javascript/api/excel/excel.worksheetfreezepanes#freezecolumns-count-)|Congela a primeira colunas da planilha no local.|
||[freezeRows (contagem?: número)](/javascript/api/excel/excel.worksheetfreezepanes#freezerows-count-)|Congela as linhas superiores da planilha no local.|
||[getLocation()](/javascript/api/excel/excel.worksheetfreezepanes#getlocation--)|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|
||[getLocationOrNullObject()](/javascript/api/excel/excel.worksheetfreezepanes#getlocationornullobject--)|Obtém um intervalo que descreve as células congeladas no modo de exibição da planilha ativa.|
||[descongelar ()](/javascript/api/excel/excel.worksheetfreezepanes#unfreeze--)|Remove todos os painéis congelados na planilha.|
|[WorksheetLoadOptions](/javascript/api/excel/excel.worksheetloadoptions)|[standardHeight](/javascript/api/excel/excel.worksheetloadoptions#standardheight)|Retorna a altura padrão de todas as linhas na planilha, em pontos. Somente leitura.|
||[standardWidth](/javascript/api/excel/excel.worksheetloadoptions#standardwidth)|Retorna ou define a largura padrão de todas as colunas na planilha.|
||[tabColor](/javascript/api/excel/excel.worksheetloadoptions#tabcolor)|Obtém ou define a cor da guia de planilha.|
|[WorksheetProtection](/javascript/api/excel/excel.worksheetprotection)|[desproteger (senha?: cadeia de caracteres)](/javascript/api/excel/excel.worksheetprotection#unprotect-password-)|Desprotege uma planilha.|
|[WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions)|[allowEditObjects](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditobjects)|Indica a opção de proteção de planilha para permitir a edição de objetos.|
||[allowEditScenarios](/javascript/api/excel/excel.worksheetprotectionoptions#alloweditscenarios)|Indica a opção de proteção de planilha para permitir a edição de cenários.|
||[selectionMode](/javascript/api/excel/excel.worksheetprotectionoptions#selectionmode)|Representa a opção de proteção da planilha do modo de seleção.|
|[WorksheetSelectionChangedEventArgs](/javascript/api/excel/excel.worksheetselectionchangedeventargs)|[address](/javascript/api/excel/excel.worksheetselectionchangedeventargs#address)|Obtém o endereço do intervalo que representa a área selecionada de uma planilha específica.|
||[tipo](/javascript/api/excel/excel.worksheetselectionchangedeventargs#type)|Obtém o tipo do evento. Para saber detalhes, confira Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetselectionchangedeventargs#worksheetid)|Obtém o id da planilha na qual a seleção foi alterada.|
|[WorksheetUpdateData](/javascript/api/excel/excel.worksheetupdatedata)|[standardWidth](/javascript/api/excel/excel.worksheetupdatedata#standardwidth)|Retorna ou define a largura padrão de todas as colunas na planilha.|
||[tabColor](/javascript/api/excel/excel.worksheetupdatedata#tabcolor)|Obtém ou define a cor da guia de planilha.|

## <a name="see-also"></a>Confira também

- [Documentação de referência da API JavaScript do Excel](/javascript/api/excel)
- [Conjuntos de requisitos da API JavaScript do Excel](./excel-api-requirement-sets.md)
