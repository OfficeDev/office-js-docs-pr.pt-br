---
title: Trabalhar com gr?ficos usando a API JavaScript do Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: c0f45892cb937a565a6855390344855f75e7473e
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a>Trabalhar com gr?ficos usando a API JavaScript do Excel

Este artigo fornece exemplos de c?digo que mostram como executar tarefas comuns com gr?ficos usando a API JavaScript do Excel. Para obter a lista completa de propriedades e m?todos aos quais os objetos **Chart** e **ChartCollection** d?o suporte, confira [Objeto Chart (API JavaScript para Excel)](https://dev.office.com/reference/add-ins/excel/chart) e [Objeto Chart Collection (API JavaScript para Excel)](https://dev.office.com/reference/add-ins/excel/chartcollection).

## <a name="create-a-chart"></a>Criar um gr?fico

O exemplo de c?digo a seguir cria um gr?fico na planilha chamada **Amostra**. O gr?fico ? de **Linha** e se baseia em dados do intervalo **A1:B13**.

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

**Novo gr?fico de linhas**

![Novo gr?fico de linhas no Excel](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a>Adicionar uma s?rie de dados a um gr?fico

O exemplo de c?digo a seguir adiciona uma s?rie de dados ao primeiro gr?fico na planilha. A nova s?rie de dados corresponde ? coluna chamada **2016** e baseia-se em dados do intervalo **D2:D5**.

> [!NOTE]
> Essa amostra usa APIs que s? est?o dispon?veis na vers?o pr?via p?blica (beta) no momento. Para executar essa amostra, voc? deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

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

**Gr?fico antes da adi??o da s?rie de dados de 2016**

![Gr?fico no Excel antes da adi??o da s?rie de dados de 2016](../images/excel-charts-data-series-before.png)

**Gr?fico ap?s a adi??o da s?rie de dados de 2016**

![Gr?fico no Excel ap?s a adi??o da s?rie de dados de 2016](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a>Definir t?tulo do gr?fico

O exemplo de c?digo a seguir define o t?tulo do primeiro gr?fico na planilha para **Sales Data by Year**. 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gr?fico ap?s defini??o do t?tulo**

![Gr?fico com t?tulo no Excel](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a>Definir propriedades de um eixo em um gr?fico

Os gr?ficos que usam o [Sistema de coordenadas cartesiano](https://en.wikipedia.org/wiki/Cartesian_coordinate_system), como gr?ficos de colunas, gr?fico de barras e gr?ficos de dispers?o cont?m um eixo de categorias e um eixo de valores. Estes exemplos mostram como definir o t?tulo e exibem a unidade de um eixo em um gr?fico.

### <a name="set-axis-title"></a>Definir t?tulo do eixo

O exemplo de c?digo a seguir define o t?tulo do eixo das categorias para o primeiro gr?fico na planilha como **Product**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gr?fico ap?s defini??o do t?tulo do eixo das categorias**

![Gr?fico com t?tulo do eixo no Excel](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a>Definir unidade de exibi??o do eixo

O exemplo de c?digo a seguir define a unidade de exibi??o do eixo dos valores para o primeiro gr?fico na planilha como **Hundreds**.

> [!NOTE]
> Essa amostra usa APIs que s? est?o dispon?veis na vers?o pr?via p?blica (beta) no momento. Para executar essa amostra, voc? deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gr?fico ap?s a defini??o da unidade de exibi??o do eixo dos valores**

![Gr?fico com unidade de exibi??o do eixo no Excel](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a>Definir visibilidade de linhas de grade em um gr?fico

O exemplo de c?digo a seguir oculta as principais linhas de grade para o eixo dos valores do primeiro gr?fico na planilha. Voc? pode mostrar as principais linhas de grade do eixo dos valores do gr?fico, definindo `chart.axes.valueAxis.majorGridlines.visible` como **true**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gr?fico com linhas de grade ocultas**

![Gr?fico com linhas de grade ocultas no Excel](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a>Linhas de tend?ncia do gr?fico

### <a name="add-a-trendline"></a>Adicionar uma linha de tend?ncia

O exemplo de c?digo a seguir adiciona uma linha de tend?ncia de m?dia m?vel ? primeira s?rie no primeiro gr?fico da planilha chamada **Amostra**. A linha de tend?ncia mostra uma m?dia m?vel de cinco per?odos.

> [!NOTE]
> Essa amostra usa APIs que s? est?o dispon?veis na vers?o pr?via p?blica (beta) no momento. Para executar essa amostra, voc? deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Gr?fico com linha de tend?ncia de m?dia m?vel**

![Gr?fico com linha de tend?ncia de m?dia m?vel no Excel](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a>Atualizar uma linha de tend?ncia

O exemplo de c?digo a seguir define a linha de tend?ncia para o tipo **Linear** para a primeira s?rie no primeiro gr?fico da planilha chamada **Amostra**.

> [!NOTE]
> Essa amostra usa APIs que s? est?o dispon?veis na vers?o pr?via p?blica (beta) no momento. Para executar essa amostra, voc? deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.

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

**Gr?fico com linha de tend?ncia linear**

![Gr?fico com linha de tend?ncia linear no Excel](../images/excel-charts-trendline-linear.png)

## <a name="see-also"></a>Veja tamb?m

- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Objeto Chart (API JavaScript para Excel)](https://dev.office.com/reference/add-ins/excel/chart) 
- [Objeto Chart Collection (API JavaScript para Excel)](https://dev.office.com/reference/add-ins/excel/chartcollection)