---
title: Trabalhar com gráficos usando a API JavaScript do Excel
description: Exemplos de código que demonstram tarefas de gráfico usando Excel API JavaScript.
ms.date: 11/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 173e20977270e84c7cef39d9ea0e326cb7b5d298
ms.sourcegitcommit: 5daf91eb3be99c88b250348186189f4dc1270956
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/01/2021
ms.locfileid: "61242066"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a>Trabalhar com gráficos usando a API JavaScript do Excel

Este artigo fornece exemplos de código que mostram como executar tarefas comuns com gráficos usando a API JavaScript do Excel.
Para ver a lista completa de propriedades e métodos que os objetos e `Chart` `ChartCollection` suportam, consulte Objeto Chart [(API JavaScript para Excel)](/javascript/api/excel/excel.chart) e Objeto da coleção [Chart (API JavaScript](/javascript/api/excel/excel.chartcollection)para Excel) .

## <a name="create-a-chart"></a>Criar um gráfico

O exemplo de código a seguir cria um gráfico na planilha chamada **Amostra**. O gráfico é de **Linha** e se baseia em dados do intervalo **A1:B13**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var dataRange = sheet.getRange("A1:B13");
    var chart = sheet.charts.add("Line", dataRange, "auto");

    chart.title.text = "Sales Data";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Novo gráfico de linhas**

![Novo gráfico de linha em Excel.](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a>Adicionar uma série de dados a um gráfico

O exemplo de código a seguir adiciona uma série de dados ao primeiro gráfico na planilha. A nova série de dados corresponde à coluna chamada **2016** e baseia-se em dados do intervalo **D2:D5**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var chart = sheet.charts.getItemAt(0);
    var dataRange = sheet.getRange("D2:D5");

    var newSeries = chart.series.add("2016");
    newSeries.setValues(dataRange);

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gráfico antes da adição da série de dados de 2016**

![Gráfico em Excel antes da adoção da série de dados de 2016.](../images/excel-charts-data-series-before.png)

**Gráfico após a adição da série de dados de 2016**

![Gráfico em Excel depois que a série de dados de 2016 foi adicionada.](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a>Definir título do gráfico

O exemplo de código a seguir define o título do primeiro gráfico na planilha para **Sales Data by Year**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gráfico após definição do título**

![Gráfico com título em Excel.](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a>Definir propriedades de um eixo em um gráfico

Os gráficos que usam o [Sistema de coordenadas cartesiano](https://en.wikipedia.org/wiki/Cartesian_coordinate_system), como gráficos de colunas, gráfico de barras e gráficos de dispersão contêm um eixo de categorias e um eixo de valores. Estes exemplos mostram como definir o título e exibem a unidade de um eixo em um gráfico.

### <a name="set-axis-title"></a>Definir título do eixo

O exemplo de código a seguir define o título do eixo das categorias para o primeiro gráfico na planilha como **Product**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gráfico após definição do título do eixo das categorias**

![Gráfico com título de eixo Excel.](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a>Definir unidade de exibição do eixo

O exemplo de código a seguir define a unidade de exibição do eixo de valor para o primeiro gráfico na planilha para **centenas**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gráfico após a definição da unidade de exibição do eixo de valor**

![Gráfico com unidade de exibição de eixo Excel.](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a>Definir visibilidade de linhas de grade em um gráfico

O exemplo de código a seguir oculta as principais linhas de grade para o eixo dos valores do primeiro gráfico na planilha. Você pode mostrar as linhas de grade principais para o eixo do valor do gráfico, definindo `chart.axes.valueAxis.majorGridlines.visible` como `true` .

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gráfico com linhas de grade ocultas**

![Gráfico com linhas de grade ocultas Excel.](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a>Linhas de tendência do gráfico

### <a name="add-a-trendline"></a>Adicionar uma linha de tendência

O exemplo de código a seguir adiciona uma linha de tendência de média móvel à primeira série no primeiro gráfico da planilha chamada **Amostra**. A linha de tendência mostra uma média móvel de cinco períodos.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gráfico com linha de tendência de média móvel**

![Gráfico com linha de tendência média móvel Excel.](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a>Atualizar uma linha de tendência

O exemplo de código a seguir define a linha de tendência para digitar para a primeira série no `Linear` primeiro gráfico da planilha chamada **Sample**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    var series = seriesCollection.getItemAt(0);
    series.trendlines.getItem(0).type = "Linear";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gráfico com linha de tendência linear**

![Gráfico com linha de tendência linear Excel.](../images/excel-charts-trendline-linear.png)

## <a name="add-and-format-a-chart-data-table"></a>Adicionar e formatar uma tabela de dados de gráfico

Você pode acessar o elemento de tabela de dados de um gráfico com o [`Chart.getDataTableOrNullObject`](/javascript/api/excel/excel.chart#getDataTableOrNullObject__) método. Este método retorna o [`ChartDataTable`](/javascript/api/excel/excel.chartdatatable) objeto. O `ChartDataTable` objeto tem propriedades de formatação booliana, como , e `visible` `showLegendKey` `showHorizontalBorder` .

A propriedade retorna o objeto, que permite que `ChartDataTable.format` você [`ChartDataTableFormat`](/javascript/api/excel/excel.chartdatatableformat) formate e estilmente a tabela de dados. O `ChartDataTableFormat` objeto oferece , e `border` `fill` `font` propriedades.

O exemplo de código a seguir mostra como adicionar uma tabela de dados a um gráfico e formatar essa tabela de dados usando `ChartDataTable` os objetos `ChartDataTableFormat` e.

```js
// This code sample adds a data table to a chart that already exists on the worksheet, 
// and then adjusts the display and format of that data table.
Excel.run(function (context) {
    // Retrieve the chart on the "Sample" worksheet.
    var chart = context.workbook.worksheets.getItem("Sample").charts.getItemAt(0);

    // Get the chart data table object and load its properties.
    var chartDataTable = chart.getDataTableOrNullObject();
    chartDataTable.load();

    // Set the display properties of the chart data table.
    chartDataTable.visible = true;
    chartDataTable.showLegendKey = true;
    chartDataTable.showHorizontalBorder = false;
    chartDataTable.showVerticalBorder = true;
    chartDataTable.showOutlineBorder = true;

    // Retrieve the chart data table format object and set font and border properties. 
    var chartDataTableFormat = chartDataTable.format;
    chartDataTableFormat.font.color = "#B76E79";
    chartDataTableFormat.font.name = "Comic Sans";
    chartDataTableFormat.border.color = "blue";

    return context.sync();
}).catch(errorHandlerFunction);
```

A captura de tela a seguir mostra a tabela de dados que o exemplo de código anterior cria.

![Um gráfico com uma tabela de dados, mostrando formatação personalizada da tabela de dados.](../images/excel-charts-data-table.png)

## <a name="export-a-chart-as-an-image"></a>Exportar um gráfico como uma imagem

Os gráficos podem ser processados como imagens fora do Excel. `Chart.getImage` retorna o gráfico como uma cadeia de caracteres codificada na base 64 representando o gráfico como uma imagem JPEG. O código a seguir mostra como obter a cadeia de caracteres de imagem e registrá-la no console.

```js
Excel.run(function (ctx) {
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
    var imageAsString = chart.getImage();
    return context.sync().then(function () {
        console.log(imageAsString.value);
        // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
    });
}).catch(errorHandlerFunction);
```

`Chart.getImage` usa três parâmetros opcionais: largura, altura e o modo de ajuste.

```typescript
getImage(width?: number, height?: number, fittingMode?: Excel.ImageFittingMode): OfficeExtension.ClientResult<string>;
```

Esses parâmetros determinam o tamanho da imagem. As imagens são sempre dimensionadas proporcionalmente. Os parâmetros de largura e altura definem limites superiores ou inferiores na imagem dimensionada. `ImageFittingMode` tem três valores com os seguintes comportamentos.

- `Fill`: a altura ou largura mínima da imagem é a altura ou largura especificada (o que for atingido primeiro ao dimensionar a imagem). Esse é o comportamento padrão quando nenhum modo de ajuste é especificado.
- `Fit`: a altura ou largura máxima da imagem é a altura ou largura especificada (o que for atingido primeiro ao dimensionar a imagem).
- `FitAndCenter`: a altura ou largura máxima da imagem é a altura ou largura especificada (o que for atingido primeiro ao dimensionar a imagem). A imagem resultante é centralizada proporcionalmente à outra dimensão.

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
