---
title: Trabalhar com gráficos usando a API JavaScript do Excel
description: Exemplos de código que demonstram tarefas de gráfico usando Excel API JavaScript.
ms.date: 07/17/2019
localization_priority: Normal
ms.openlocfilehash: a7199aae31e917b0609a47cc69b5e52279d43b24
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349571"
---
# <a name="work-with-charts-using-the-excel-javascript-api"></a><span data-ttu-id="ed4f3-103">Trabalhar com gráficos usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="ed4f3-103">Work with charts using the Excel JavaScript API</span></span>

<span data-ttu-id="ed4f3-104">Este artigo fornece exemplos de código que mostram como executar tarefas comuns com gráficos usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-104">This article provides code samples that show how to perform common tasks with charts using the Excel JavaScript API.</span></span>
<span data-ttu-id="ed4f3-105">Para ver a lista completa de propriedades e métodos que os objetos e `Chart` `ChartCollection` suportam, consulte Objeto Chart [(API JavaScript para Excel)](/javascript/api/excel/excel.chart) e Objeto da coleção [Chart (API JavaScript](/javascript/api/excel/excel.chartcollection)para Excel) .</span><span class="sxs-lookup"><span data-stu-id="ed4f3-105">For the complete list of properties and methods that the `Chart` and `ChartCollection` objects support, see [Chart Object (JavaScript API for Excel)](/javascript/api/excel/excel.chart) and [Chart Collection Object (JavaScript API for Excel)](/javascript/api/excel/excel.chartcollection).</span></span>

## <a name="create-a-chart"></a><span data-ttu-id="ed4f3-106">Criar um gráfico</span><span class="sxs-lookup"><span data-stu-id="ed4f3-106">Create a chart</span></span>

<span data-ttu-id="ed4f3-p102">O exemplo de código a seguir cria um gráfico na planilha chamada **Amostra**. O gráfico é de **Linha** e se baseia em dados do intervalo **A1:B13**.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-p102">The following code sample creates a chart in the worksheet named **Sample**. The chart is a **Line** chart that is based upon data in the range **A1:B13**.</span></span>

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

<span data-ttu-id="ed4f3-109">**Novo gráfico de linhas**</span><span class="sxs-lookup"><span data-stu-id="ed4f3-109">**New line chart**</span></span>

![Novo gráfico de linha em Excel.](../images/excel-charts-create-line.png)


## <a name="add-a-data-series-to-a-chart"></a><span data-ttu-id="ed4f3-111">Adicionar uma série de dados a um gráfico</span><span class="sxs-lookup"><span data-stu-id="ed4f3-111">Add a data series to a chart</span></span>

<span data-ttu-id="ed4f3-p103">O exemplo de código a seguir adiciona uma série de dados ao primeiro gráfico na planilha. A nova série de dados corresponde à coluna chamada **2016** e baseia-se em dados do intervalo **D2:D5**.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-p103">The following code sample adds a data series to the first chart in the worksheet. The new data series corresponds to the column named **2016** and is based upon data in the range **D2:D5**.</span></span>

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

<span data-ttu-id="ed4f3-114">**Gráfico antes da adição da série de dados de 2016**</span><span class="sxs-lookup"><span data-stu-id="ed4f3-114">**Chart before the 2016 data series is added**</span></span>

![Gráfico em Excel antes da adoção da série de dados de 2016.](../images/excel-charts-data-series-before.png)

<span data-ttu-id="ed4f3-116">**Gráfico após a adição da série de dados de 2016**</span><span class="sxs-lookup"><span data-stu-id="ed4f3-116">**Chart after the 2016 data series is added**</span></span>

![Gráfico em Excel depois que a série de dados de 2016 foi adicionada.](../images/excel-charts-data-series-after.png)

## <a name="set-chart-title"></a><span data-ttu-id="ed4f3-118">Definir título do gráfico</span><span class="sxs-lookup"><span data-stu-id="ed4f3-118">Set chart title</span></span>

<span data-ttu-id="ed4f3-119">O exemplo de código a seguir define o título do primeiro gráfico na planilha para **Sales Data by Year**.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-119">The following code sample sets the title of the first chart in the worksheet to **Sales Data by Year**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.title.text = "Sales Data by Year";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ed4f3-120">**Gráfico após definição do título**</span><span class="sxs-lookup"><span data-stu-id="ed4f3-120">**Chart after title is set**</span></span>

![Gráfico com título em Excel.](../images/excel-charts-title-set.png)

## <a name="set-properties-of-an-axis-in-a-chart"></a><span data-ttu-id="ed4f3-122">Definir propriedades de um eixo em um gráfico</span><span class="sxs-lookup"><span data-stu-id="ed4f3-122">Set properties of an axis in a chart</span></span>

<span data-ttu-id="ed4f3-p104">Os gráficos que usam o [Sistema de coordenadas cartesiano](https://en.wikipedia.org/wiki/Cartesian_coordinate_system), como gráficos de colunas, gráfico de barras e gráficos de dispersão contêm um eixo de categorias e um eixo de valores. Estes exemplos mostram como definir o título e exibem a unidade de um eixo em um gráfico.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-p104">Charts that use the [Cartesian coordinate system](https://en.wikipedia.org/wiki/Cartesian_coordinate_system) such as column charts, bar charts, and scatter charts contain a category axis and a value axis. These examples show how to set the title and display unit of an axis in a chart.</span></span>

### <a name="set-axis-title"></a><span data-ttu-id="ed4f3-125">Definir título do eixo</span><span class="sxs-lookup"><span data-stu-id="ed4f3-125">Set axis title</span></span>

<span data-ttu-id="ed4f3-126">O exemplo de código a seguir define o título do eixo das categorias para o primeiro gráfico na planilha como **Product**.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-126">The following code sample sets the title of the category axis for the first chart in the worksheet to **Product**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.categoryAxis.title.text = "Product";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ed4f3-127">**Gráfico após definição do título do eixo das categorias**</span><span class="sxs-lookup"><span data-stu-id="ed4f3-127">**Chart after title of category axis is set**</span></span>

![Gráfico com título de eixo Excel.](../images/excel-charts-axis-title-set.png)

### <a name="set-axis-display-unit"></a><span data-ttu-id="ed4f3-129">Definir unidade de exibição do eixo</span><span class="sxs-lookup"><span data-stu-id="ed4f3-129">Set axis display unit</span></span>

<span data-ttu-id="ed4f3-130">O exemplo de código a seguir define a unidade de exibição do eixo de valor para o primeiro gráfico na planilha para **centenas**.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-130">The following code sample sets the display unit of the value axis for the first chart in the worksheet to **Hundreds**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.displayUnit = "Hundreds";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ed4f3-131">**Gráfico após a definição da unidade de exibição do eixo de valor**</span><span class="sxs-lookup"><span data-stu-id="ed4f3-131">**Chart after display unit of value axis is set**</span></span>

![Gráfico com unidade de exibição de eixo Excel.](../images/excel-charts-axis-display-unit-set.png)

## <a name="set-visibility-of-gridlines-in-a-chart"></a><span data-ttu-id="ed4f3-133">Definir visibilidade de linhas de grade em um gráfico</span><span class="sxs-lookup"><span data-stu-id="ed4f3-133">Set visibility of gridlines in a chart</span></span>

<span data-ttu-id="ed4f3-134">O exemplo de código a seguir oculta as principais linhas de grade para o eixo dos valores do primeiro gráfico na planilha.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-134">The following code sample hides the major gridlines for the value axis of the first chart in the worksheet.</span></span> <span data-ttu-id="ed4f3-135">Você pode mostrar as linhas de grade principais para o eixo do valor do gráfico, definindo `chart.axes.valueAxis.majorGridlines.visible` como `true` .</span><span class="sxs-lookup"><span data-stu-id="ed4f3-135">You can show the major gridlines for the value axis of the chart, by setting `chart.axes.valueAxis.majorGridlines.visible` to `true`.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    chart.axes.valueAxis.majorGridlines.visible = false;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ed4f3-136">**Gráfico com linhas de grade ocultas**</span><span class="sxs-lookup"><span data-stu-id="ed4f3-136">**Chart with gridlines hidden**</span></span>

![Gráfico com linhas de grade ocultas Excel.](../images/excel-charts-gridlines-removed.png)

## <a name="chart-trendlines"></a><span data-ttu-id="ed4f3-138">Linhas de tendência do gráfico</span><span class="sxs-lookup"><span data-stu-id="ed4f3-138">Chart trendlines</span></span>

### <a name="add-a-trendline"></a><span data-ttu-id="ed4f3-139">Adicionar uma linha de tendência</span><span class="sxs-lookup"><span data-stu-id="ed4f3-139">Add a trendline</span></span>

<span data-ttu-id="ed4f3-p106">O exemplo de código a seguir adiciona uma linha de tendência de média móvel à primeira série no primeiro gráfico da planilha chamada **Amostra**. A linha de tendência mostra uma média móvel de cinco períodos.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-p106">The following code sample adds a moving average trendline to the first series in the first chart in the worksheet named **Sample**. The trendline shows a moving average over 5 periods.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var chart = sheet.charts.getItemAt(0);
    var seriesCollection = chart.series;
    seriesCollection.getItemAt(0).trendlines.add("MovingAverage").movingAveragePeriod = 5;

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ed4f3-142">**Gráfico com linha de tendência de média móvel**</span><span class="sxs-lookup"><span data-stu-id="ed4f3-142">**Chart with moving average trendline**</span></span>

![Gráfico com linha de tendência média móvel Excel.](../images/excel-charts-create-trendline.png)

### <a name="update-a-trendline"></a><span data-ttu-id="ed4f3-144">Atualizar uma linha de tendência</span><span class="sxs-lookup"><span data-stu-id="ed4f3-144">Update a trendline</span></span>

<span data-ttu-id="ed4f3-145">O exemplo de código a seguir define a linha de tendência para digitar para a primeira série no `Linear` primeiro gráfico da planilha chamada **Sample**.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-145">The following code sample sets the trendline to type `Linear` for the first series in the first chart in the worksheet named **Sample**.</span></span>

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

<span data-ttu-id="ed4f3-146">**Gráfico com linha de tendência linear**</span><span class="sxs-lookup"><span data-stu-id="ed4f3-146">**Chart with linear trendline**</span></span>

![Gráfico com linha de tendência linear Excel.](../images/excel-charts-trendline-linear.png)

## <a name="export-a-chart-as-an-image"></a><span data-ttu-id="ed4f3-148">Exportar um gráfico como uma imagem</span><span class="sxs-lookup"><span data-stu-id="ed4f3-148">Export a chart as an image</span></span>

<span data-ttu-id="ed4f3-149">Os gráficos podem ser processados como imagens fora do Excel.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-149">Charts can be rendered as images outside of Excel.</span></span> <span data-ttu-id="ed4f3-150">`Chart.getImage` retorna o gráfico como uma cadeia de caracteres codificada na base 64 representando o gráfico como uma imagem JPEG.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-150">`Chart.getImage` returns the chart as a base64-encoded string representing the chart as a JPEG image.</span></span> <span data-ttu-id="ed4f3-151">O código a seguir mostra como obter a cadeia de caracteres de imagem e registrá-la no console.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-151">The following code shows how to get the image string and log it to the console.</span></span>

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

<span data-ttu-id="ed4f3-152">`Chart.getImage` usa três parâmetros opcionais: largura, altura e o modo de ajuste.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-152">`Chart.getImage` takes three optional parameters: width, height, and the fitting mode.</span></span>

```typescript
getImage(width?: number, height?: number, fittingMode?: Excel.ImageFittingMode): OfficeExtension.ClientResult<string>;
```

<span data-ttu-id="ed4f3-153">Esses parâmetros determinam o tamanho da imagem.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-153">These parameters determine the size of the image.</span></span> <span data-ttu-id="ed4f3-154">As imagens são sempre dimensionadas proporcionalmente.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-154">Images are always proportionally scaled.</span></span> <span data-ttu-id="ed4f3-155">Os parâmetros de largura e altura definem limites superiores ou inferiores na imagem dimensionada.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-155">The width and height parameters put upper or lower bounds on the scaled image.</span></span> <span data-ttu-id="ed4f3-156">`ImageFittingMode` tem três valores com os seguintes comportamentos.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-156">`ImageFittingMode` has three values with the following behaviors.</span></span>

- <span data-ttu-id="ed4f3-157">`Fill`: a altura ou largura mínima da imagem é a altura ou largura especificada (o que for atingido primeiro ao dimensionar a imagem).</span><span class="sxs-lookup"><span data-stu-id="ed4f3-157">`Fill`: The image's minimum height or width is the specified height or width (whichever is reached first when scaling the image).</span></span> <span data-ttu-id="ed4f3-158">Esse é o comportamento padrão quando nenhum modo de ajuste é especificado.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-158">This is the default behavior when no fitting mode is specified.</span></span>
- <span data-ttu-id="ed4f3-159">`Fit`: a altura ou largura máxima da imagem é a altura ou largura especificada (o que for atingido primeiro ao dimensionar a imagem).</span><span class="sxs-lookup"><span data-stu-id="ed4f3-159">`Fit`: The image's maximum height or width is the specified height or width (whichever is reached first when scaling the image).</span></span>
- <span data-ttu-id="ed4f3-160">`FitAndCenter`: a altura ou largura máxima da imagem é a altura ou largura especificada (o que for atingido primeiro ao dimensionar a imagem).</span><span class="sxs-lookup"><span data-stu-id="ed4f3-160">`FitAndCenter`: The image's maximum height or width is the specified height or width (whichever is reached first when scaling the image).</span></span> <span data-ttu-id="ed4f3-161">A imagem resultante é centralizada proporcionalmente à outra dimensão.</span><span class="sxs-lookup"><span data-stu-id="ed4f3-161">The resulting image is centered relative to the other dimension.</span></span>

## <a name="see-also"></a><span data-ttu-id="ed4f3-162">Confira também</span><span class="sxs-lookup"><span data-stu-id="ed4f3-162">See also</span></span>

- [<span data-ttu-id="ed4f3-163">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ed4f3-163">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
