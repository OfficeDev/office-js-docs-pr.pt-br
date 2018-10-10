---
title: Trabalhar com gráficos usando a API JavaScript do Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 80b537ec66caf6e173dfe4453a257c5963156e6f
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459298"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a><span data-ttu-id="94997-102">Trabalhar com gráficos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="94997-102">Work with Charts using the Excel JavaScript API</span></span>

<span data-ttu-id="94997-p101">Este artigo fornece exemplos de código que mostram como realizar tarefas comuns com gráficos usando a API JavaScript do Excel. Para obter uma lista completa de propriedades e métodos que os objetos **Chart** e **ChartCollection** suportam, consulte [Objeto de gráfico (API JavaScript do Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chart?view=office-js) e [Objeto da coleção de gráfico (API JavaScript do Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="94997-p101">This article provides code samples that show how to perform common tasks with charts using the Excel JavaScript API. For the complete list of properties and methods that the **Chart** and **ChartCollection** objects support, see [Chart Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chart?view=office-js) and [Chart Collection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.chartcollection?view=office-js).</span></span>

## <a name="create-a-chart"></a><span data-ttu-id="94997-105">Criar um gráfico</span><span class="sxs-lookup"><span data-stu-id="94997-105">Create a chart</span></span>

<span data-ttu-id="94997-p102">O exemplo de código a seguir cria um gráfico na planilha chamada **Amostra**. O gráfico é um gráfico de **linhas** que se baseia em dados no intervalo **A1:B13**.</span><span class="sxs-lookup"><span data-stu-id="94997-p102">The following code sample creates a chart in the worksheet named **Sample**. The chart is a **Line** chart that is based upon data in the range **A1:B13**.</span></span>

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

<span data-ttu-id="94997-108">**Novo gráfico de linhas**</span><span class="sxs-lookup"><span data-stu-id="94997-108">**New line chart**</span></span>

![Novo gráfico de linhas no Excel](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a><span data-ttu-id="94997-110">Adicionar uma série de dados a um gráfico</span><span class="sxs-lookup"><span data-stu-id="94997-110">Add a data series to a chart</span></span>

<span data-ttu-id="94997-p103">O exemplo de código a seguir adiciona uma série de dados ao primeiro gráfico na planilha. A nova série de dados corresponde à coluna denominada **2016** e baseia-se em dados no intervalo **D2:D5**.</span><span class="sxs-lookup"><span data-stu-id="94997-p103">The following code sample adds a data series to the first chart in the worksheet. The new data series corresponds to the column named **2016** and is based upon data in the range **D2:D5**.</span></span>

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

<span data-ttu-id="94997-113">**Gráfico antes da adição da série de dados de 2016**</span><span class="sxs-lookup"><span data-stu-id="94997-113">**Chart before the 2016 data series is added**</span></span>

![Gráfico no Excel antes da adição da série de dados de 2016](../images/excel-charts-data-series-before.png)

<span data-ttu-id="94997-115">**Gráfico após a adição da série de dados de 2016**</span><span class="sxs-lookup"><span data-stu-id="94997-115">**Chart after the 2016 data series is added**</span></span>

![Gráfico no Excel após a adição da série de dados de 2016](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a><span data-ttu-id="94997-117">Definir título do gráfico</span><span class="sxs-lookup"><span data-stu-id="94997-117">Set chart title</span></span>

<span data-ttu-id="94997-118">O exemplo de código a seguir define o título do primeiro gráfico na planilha para **Sales Data by Year**.</span><span class="sxs-lookup"><span data-stu-id="94997-118">The following code sample sets the title of the first chart in the worksheet to **Sales Data by Year**.</span></span> 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="94997-119">**Gráfico após definição do título**</span><span class="sxs-lookup"><span data-stu-id="94997-119">**Chart after title is set**</span></span>

![Gráfico com título no Excel](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a><span data-ttu-id="94997-121">Definir propriedades de um eixo em um gráfico</span><span class="sxs-lookup"><span data-stu-id="94997-121">Set properties of an axis in a chart</span></span>

<span data-ttu-id="94997-122">Os gráficos que usam o [Sistema de coordenadas cartesiano](https://en.wikipedia.org/wiki/Cartesian_coordinate_system), como gráficos de colunas, gráficos de barras e gráficos de dispersão contêm um eixo de categorias e um eixo de valores.</span><span class="sxs-lookup"><span data-stu-id="94997-122">Charts that use the [Cartesian coordinate system](https://en.wikipedia.org/wiki/Cartesian_coordinate_system) such as column charts, bar charts, and scatter charts contain a category axis and a value axis.</span></span> <span data-ttu-id="94997-123">Estes exemplos mostram como definir o título e exibem a unidade de um eixo em um gráfico.</span><span class="sxs-lookup"><span data-stu-id="94997-123">These examples show how to set the title and display unit of an axis in a chart.</span></span>

### <a name="set-axis-title"></a><span data-ttu-id="94997-124">Definir título do eixo</span><span class="sxs-lookup"><span data-stu-id="94997-124">Set axis title</span></span>

<span data-ttu-id="94997-125">O exemplo de código a seguir define o título do eixo da categoria para o primeiro gráfico na planilha como **Product**.</span><span class="sxs-lookup"><span data-stu-id="94997-125">The following code sample sets the title of the category axis for the first chart in the worksheet to **Product**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="94997-126">**Gráfico após definição do título do eixo da categoria**</span><span class="sxs-lookup"><span data-stu-id="94997-126">**Chart after title of category axis is set**</span></span>

![Gráfico com título do eixo no Excel](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a><span data-ttu-id="94997-128">Definir unidade de exibição do eixo</span><span class="sxs-lookup"><span data-stu-id="94997-128">Set axis display unit</span></span>

<span data-ttu-id="94997-129">O exemplo de código a seguir define a unidade de exibição do eixo dos valores para o primeiro gráfico na planilha como **Hundreds**.</span><span class="sxs-lookup"><span data-stu-id="94997-129">The following code sample sets the display unit of the value axis for the first chart in the worksheet to **Hundreds**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="94997-130">**Gráfico após a definição da unidade de exibição do eixo dos valores**</span><span class="sxs-lookup"><span data-stu-id="94997-130">**Chart after display unit of value axis is set**</span></span>

![Gráfico com unidade de exibição do eixo no Excel](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a><span data-ttu-id="94997-132">Definir visibilidade de linhas de grade em um gráfico</span><span class="sxs-lookup"><span data-stu-id="94997-132">Set visibility of gridlines in a chart</span></span>

<span data-ttu-id="94997-133">O exemplo de código a seguir oculta as principais linhas de grade para o eixo de valores do primeiro gráfico na planilha.</span><span class="sxs-lookup"><span data-stu-id="94997-133">The following code sample hides the major gridlines for the value axis of the first chart in the worksheet.</span></span> <span data-ttu-id="94997-134">Você pode mostrar as principais linhas de grade do eixo de valores do gráfico, definindo `chart.axes.valueAxis.majorGridlines.visible` como **true**.</span><span class="sxs-lookup"><span data-stu-id="94997-134">You can show the major gridlines for the value axis of the chart, by setting `chart.axes.valueAxis.majorGridlines.visible` to **true**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="94997-135">**Gráfico com linhas de grade ocultas**</span><span class="sxs-lookup"><span data-stu-id="94997-135">**Chart with gridlines hidden**</span></span>

![Gráfico com linhas de grade ocultas no Excel](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a><span data-ttu-id="94997-137">Linhas de tendência do gráfico</span><span class="sxs-lookup"><span data-stu-id="94997-137">Chart trendlines</span></span>

### <a name="add-a-trendline"></a><span data-ttu-id="94997-138">Adicionar uma linha de tendência</span><span class="sxs-lookup"><span data-stu-id="94997-138">Add a trendline</span></span>

<span data-ttu-id="94997-p106">O exemplo de código a seguir adiciona uma linha de tendência de média móvel à primeira série no primeiro gráfico da planilha chamada **Amostra**. A linha de tendência mostra uma média móvel de cinco períodos.</span><span class="sxs-lookup"><span data-stu-id="94997-p106">The following code sample adds a moving average trendline to the first series in the first chart in the worksheet named **Sample**. The trendline shows a moving average over 5 periods.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="94997-141">**Gráfico com linha de tendência de média móvel**</span><span class="sxs-lookup"><span data-stu-id="94997-141">**Chart with moving average trendline**</span></span>

![Gráfico com linha de tendência de média móvel no Excel](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a><span data-ttu-id="94997-143">Atualizar uma linha de tendência</span><span class="sxs-lookup"><span data-stu-id="94997-143">Update a trendline</span></span>

<span data-ttu-id="94997-144">O exemplo de código a seguir define a linha de tendência para o tipo **Linear** para a primeira série no primeiro gráfico da planilha chamada **Amostra**.</span><span class="sxs-lookup"><span data-stu-id="94997-144">The following code sample sets the trendline to type **Linear** for the first series in the first chart in the worksheet named **Sample**.</span></span>

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

<span data-ttu-id="94997-145">**Gráfico com linha de tendência linear**</span><span class="sxs-lookup"><span data-stu-id="94997-145">**Chart with linear trendline**</span></span>

![Gráfico com linha de tendência linear no Excel](../images/excel-charts-trendline-linear.png)

## <a name="see-also"></a><span data-ttu-id="94997-147">Confira também</span><span class="sxs-lookup"><span data-stu-id="94997-147">See also</span></span>

- [<span data-ttu-id="94997-148">Conceitos de programação fundamentais com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="94997-148">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
