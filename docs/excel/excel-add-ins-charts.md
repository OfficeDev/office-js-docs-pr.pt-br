---
title: Trabalhar com gráficos usando a API JavaScript do Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: b804e2130e30626a9caf21bca1f3955c57a3f94c
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457548"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a>Trabalhar com gráficos usando a API JavaScript do Excel

Este artigo fornece exemplos de código que mostram como executar tarefas comuns com gráficos usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos aos quais os objetos **Chart** e **ChartCollection** dão suporte, confira [Objeto Chart (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chart) e [Objeto Chart Collection (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection).

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

![Novo gráfico de linhas no Excel](../images/excel-charts-create-line.png)


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

![Gráfico no Excel antes da adição da série de dados de 2016](../images/excel-charts-data-series-before.png)

**Gráfico após a adição da série de dados de 2016**

![Gráfico no Excel após a adição da série de dados de 2016](../images/excel-charts-data-series-after.png)

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

![Gráfico com título no Excel](../images/excel-charts-title-set.png)

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

![Gráfico com título do eixo no Excel](../images/excel-charts-axis-title-set.png)

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

![Gráfico com unidade de exibição do eixo no Excel](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a>Definir visibilidade de linhas de grade em um gráfico

O exemplo de código a seguir oculta as principais linhas de grade para o eixo dos valores do primeiro gráfico na planilha. Você pode mostrar as principais linhas de grade do eixo dos valores do gráfico, definindo `chart.axes.valueAxis.majorGridlines.visible` como **true**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gráfico com linhas de grade ocultas**

![Gráfico com linhas de grade ocultas no Excel](../images/excel-charts-gridlines-removed.png)

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

![Gráfico com linha de tendência de média móvel no Excel](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a>Atualizar uma linha de tendência

O exemplo de código a seguir define a linha de tendência para o tipo **Linear** para a primeira série no primeiro gráfico da planilha chamada **Amostra**.

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

![Gráfico com linha de tendência linear no Excel](../images/excel-charts-trendline-linear.png)

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
