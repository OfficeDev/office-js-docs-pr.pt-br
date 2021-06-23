---
title: Definir e obter valores de intervalo, texto ou fórmulas usando Excel API JavaScript
description: Saiba como usar a EXCEL JavaScript para definir e obter valores de intervalo, texto ou fórmulas.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5d9d1bf3b248585bf27ac591754cfa4eb4dd0fbc
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075744"
---
# <a name="set-and-get-range-values-text-or-formulas-using-the-excel-javascript-api"></a><span data-ttu-id="d06fc-103">Definir e obter valores de intervalo, texto ou fórmulas usando Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="d06fc-103">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>

<span data-ttu-id="d06fc-104">Este artigo fornece exemplos de código que configuram e obteram valores de intervalo, texto ou fórmulas com a API JavaScript Excel JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d06fc-104">This article provides code samples that set and get range values, text, or formulas with the Excel JavaScript API.</span></span> <span data-ttu-id="d06fc-105">Para ver a lista completa de propriedades e métodos que o `Range` objeto oferece suporte, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="d06fc-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-values-or-formulas"></a><span data-ttu-id="d06fc-106">Definir valores ou fórmulas</span><span class="sxs-lookup"><span data-stu-id="d06fc-106">Set values or formulas</span></span>

<span data-ttu-id="d06fc-107">Os exemplos de código a seguir configuram valores e fórmulas para uma única célula ou um intervalo de células.</span><span class="sxs-lookup"><span data-stu-id="d06fc-107">The following code samples set values and formulas for a single cell or a range of cells.</span></span>

### <a name="set-value-for-a-single-cell"></a><span data-ttu-id="d06fc-108">Definir valor para uma única célula</span><span class="sxs-lookup"><span data-stu-id="d06fc-108">Set value for a single cell</span></span>

<span data-ttu-id="d06fc-109">O exemplo de código a seguir define o valor da célula **C3** como "5" e, em seguida, define a largura das colunas para melhor ajustar os dados.</span><span class="sxs-lookup"><span data-stu-id="d06fc-109">The following code sample sets the value of cell **C3** to "5" and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-value-is-updated"></a><span data-ttu-id="d06fc-110">Dados antes da atualização do valor da célula</span><span class="sxs-lookup"><span data-stu-id="d06fc-110">Data before cell value is updated</span></span>

![Dados em Excel antes que o valor da célula seja atualizado.](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-value-is-updated"></a><span data-ttu-id="d06fc-112">Dados após a atualização do valor da célula</span><span class="sxs-lookup"><span data-stu-id="d06fc-112">Data after cell value is updated</span></span>

![Dados em Excel depois que o valor da célula é atualizado.](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a><span data-ttu-id="d06fc-114">Definir valores para um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="d06fc-114">Set values for a range of cells</span></span>

<span data-ttu-id="d06fc-115">O exemplo de código a seguir define valores das células no intervalo **B5:D5** e, em seguida, define a largura das colunas para melhor ajustar os dados.</span><span class="sxs-lookup"><span data-stu-id="d06fc-115">The following code sample sets values for the cells in the range **B5:D5** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["Potato Chips", 10, 1.80],
    ];

    var range = sheet.getRange("B5:D5");
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-values-are-updated"></a><span data-ttu-id="d06fc-116">Dados antes da atualização dos valores da célula</span><span class="sxs-lookup"><span data-stu-id="d06fc-116">Data before cell values are updated</span></span>

![Dados em Excel antes que os valores da célula sejam atualizados.](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-values-are-updated"></a><span data-ttu-id="d06fc-118">Dados após a atualização dos valores da célula</span><span class="sxs-lookup"><span data-stu-id="d06fc-118">Data after cell values are updated</span></span>

![Dados em Excel depois que os valores da célula são atualizados.](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a><span data-ttu-id="d06fc-120">Definir fórmula para uma única célula</span><span class="sxs-lookup"><span data-stu-id="d06fc-120">Set formula for a single cell</span></span>

<span data-ttu-id="d06fc-121">O exemplo de código a seguir define uma fórmula para a célula **E3** e, em seguida, define a largura das colunas para melhor ajustar os dados.</span><span class="sxs-lookup"><span data-stu-id="d06fc-121">The following code sample sets a formula for cell **E3** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formula-is-set"></a><span data-ttu-id="d06fc-122">Dados antes da definição da fórmula da célula</span><span class="sxs-lookup"><span data-stu-id="d06fc-122">Data before cell formula is set</span></span>

![Dados em Excel antes que a fórmula da célula seja definida.](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formula-is-set"></a><span data-ttu-id="d06fc-124">Dados após a definição da fórmula da célula</span><span class="sxs-lookup"><span data-stu-id="d06fc-124">Data after cell formula is set</span></span>

![Dados em Excel depois que a fórmula da célula é definida.](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a><span data-ttu-id="d06fc-126">Definir fórmulas para um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="d06fc-126">Set formulas for a range of cells</span></span>

<span data-ttu-id="d06fc-127">O exemplo de código a seguir define fórmulas para células no intervalo **E2:E6** e, em seguida, define a largura das colunas para melhor ajustar os dados.</span><span class="sxs-lookup"><span data-stu-id="d06fc-127">The following code sample sets formulas for cells in the range **E2:E6** and then sets the width of the columns to best fit the data.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];

    var range = sheet.getRange("E3:E6");
    range.formulas = data;
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="data-before-cell-formulas-are-set"></a><span data-ttu-id="d06fc-128">Dados antes da definição das fórmulas da célula</span><span class="sxs-lookup"><span data-stu-id="d06fc-128">Data before cell formulas are set</span></span>

![Dados em Excel antes que as fórmulas da célula sejam definidas.](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formulas-are-set"></a><span data-ttu-id="d06fc-130">Dados após a definição das fórmulas da célula</span><span class="sxs-lookup"><span data-stu-id="d06fc-130">Data after cell formulas are set</span></span>

![Dados em Excel depois que as fórmulas da célula são definidas.](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a><span data-ttu-id="d06fc-132">Obter valores, texto ou fórmulas</span><span class="sxs-lookup"><span data-stu-id="d06fc-132">Get values, text, or formulas</span></span>

<span data-ttu-id="d06fc-133">Esses exemplos de código obterão valores, texto e fórmulas de um intervalo de células.</span><span class="sxs-lookup"><span data-stu-id="d06fc-133">These code samples get values, text, and formulas from a range of cells.</span></span>

### <a name="get-values-from-a-range-of-cells"></a><span data-ttu-id="d06fc-134">Obter valores de um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="d06fc-134">Get values from a range of cells</span></span>

<span data-ttu-id="d06fc-135">O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega sua propriedade e grava `values` os valores no console.</span><span class="sxs-lookup"><span data-stu-id="d06fc-135">The following code sample gets the range **B2:E6**, loads its `values` property, and writes the values to the console.</span></span> <span data-ttu-id="d06fc-136">A `values` propriedade de um intervalo especifica os valores brutos que as células contêm.</span><span class="sxs-lookup"><span data-stu-id="d06fc-136">The `values` property of a range specifies the raw values that the cells contain.</span></span> <span data-ttu-id="d06fc-137">Mesmo que algumas células em um intervalo contenham fórmulas, a propriedade do intervalo especifica os valores brutos dessas células, não qualquer `values` uma das fórmulas.</span><span class="sxs-lookup"><span data-stu-id="d06fc-137">Even if some cells in a range contain formulas, the `values` property of the range specifies the raw values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("values");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.values, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="d06fc-138">Dados no intervalo (valores na coluna E são um resultado de fórmulas)</span><span class="sxs-lookup"><span data-stu-id="d06fc-138">Data in range (values in column E are a result of formulas)</span></span>

![Dados em Excel depois que as fórmulas da célula são definidas.](../images/excel-ranges-set-formulas.png)

#### <a name="rangevalues-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="d06fc-140">range.values (conforme registrado em log no console pelo exemplo de código acima)</span><span class="sxs-lookup"><span data-stu-id="d06fc-140">range.values (as logged to the console by the code sample above)</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        15
    ],
    [
        "Coffee",
        1,
        34.5,
        34.5
    ],
    [
        "Chocolate",
        5,
        9.56,
        47.8
    ],
    [
        "",
        "",
        "",
        97.3
    ]
]
```

### <a name="get-text-from-a-range-of-cells"></a><span data-ttu-id="d06fc-141">Obter texto de um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="d06fc-141">Get text from a range of cells</span></span>

<span data-ttu-id="d06fc-142">O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega sua propriedade `text` e grava-a no console.</span><span class="sxs-lookup"><span data-stu-id="d06fc-142">The following code sample gets the range **B2:E6**, loads its `text` property, and writes it to the console.</span></span> <span data-ttu-id="d06fc-143">A `text` propriedade de um intervalo especifica os valores de exibição para células no intervalo.</span><span class="sxs-lookup"><span data-stu-id="d06fc-143">The `text` property of a range specifies the display values for cells in the range.</span></span> <span data-ttu-id="d06fc-144">Mesmo que algumas células em um intervalo contenham fórmulas, a propriedade do intervalo especifica os valores de exibição dessas células, não qualquer `text` uma das fórmulas.</span><span class="sxs-lookup"><span data-stu-id="d06fc-144">Even if some cells in a range contain formulas, the `text` property of the range specifies the display values for those cells, not any of the formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("text");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.text, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="d06fc-145">Dados no intervalo (valores na coluna E são um resultado de fórmulas)</span><span class="sxs-lookup"><span data-stu-id="d06fc-145">Data in range (values in column E are a result of formulas)</span></span>

![Dados em Excel depois que as fórmulas da célula são definidas.](../images/excel-ranges-set-formulas.png)

#### <a name="rangetext-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="d06fc-147">range.text (conforme registrado em log no console pelo exemplo de código acima)</span><span class="sxs-lookup"><span data-stu-id="d06fc-147">range.text (as logged to the console by the code sample above)</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        "2",
        "7.5",
        "15"
    ],
    [
        "Coffee",
        "1",
        "34.5",
        "34.5"
    ],
    [
        "Chocolate",
        "5",
        "9.56",
        "47.8"
    ],
    [
        "",
        "",
        "",
        "97.3"
    ]
]
```

### <a name="get-formulas-from-a-range-of-cells"></a><span data-ttu-id="d06fc-148">Obter fórmulas de um intervalo de células</span><span class="sxs-lookup"><span data-stu-id="d06fc-148">Get formulas from a range of cells</span></span>

<span data-ttu-id="d06fc-149">O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega sua propriedade `formulas` e grava-a no console.</span><span class="sxs-lookup"><span data-stu-id="d06fc-149">The following code sample gets the range **B2:E6**, loads its `formulas` property, and writes it to the console.</span></span> <span data-ttu-id="d06fc-150">A propriedade de um intervalo especifica as fórmulas para células no intervalo que contêm fórmulas e os valores brutos para células no intervalo que não `formulas` contêm fórmulas.</span><span class="sxs-lookup"><span data-stu-id="d06fc-150">The `formulas` property of a range specifies the formulas for cells in the range that contain formulas and the raw values for cells in the range that do not contain formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E6");
    range.load("formulas");

    return context.sync()
        .then(function () {
            console.log(JSON.stringify(range.formulas, null, 4));
        });
}).catch(errorHandlerFunction);
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a><span data-ttu-id="d06fc-151">Dados no intervalo (valores na coluna E são um resultado de fórmulas)</span><span class="sxs-lookup"><span data-stu-id="d06fc-151">Data in range (values in column E are a result of formulas)</span></span>

![Dados em Excel depois que as fórmulas da célula são definidas.](../images/excel-ranges-set-formulas.png)

#### <a name="rangeformulas-as-logged-to-the-console-by-the-code-sample-above"></a><span data-ttu-id="d06fc-153">range.formulas (conforme registrado em log no console pelo exemplo de código acima)</span><span class="sxs-lookup"><span data-stu-id="d06fc-153">range.formulas (as logged to the console by the code sample above)</span></span>

```json
[
    [
        "Product",
        "Qty",
        "Unit Price",
        "Total Price"
    ],
    [
        "Almonds",
        2,
        7.5,
        "=C3 * D3"
    ],
    [
        "Coffee",
        1,
        34.5,
        "=C4 * D4"
    ],
    [
        "Chocolate",
        5,
        9.56,
        "=C5 * D5"
    ],
    [
        "",
        "",
        "",
        "=SUM(E3:E5)"
    ]
]
```

## <a name="see-also"></a><span data-ttu-id="d06fc-154">Confira também</span><span class="sxs-lookup"><span data-stu-id="d06fc-154">See also</span></span>

- [<span data-ttu-id="d06fc-155">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="d06fc-155">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d06fc-156">Trabalhar com células usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="d06fc-156">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="d06fc-157">Definir e obter intervalos usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="d06fc-157">Set and get ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get.md)
- [<span data-ttu-id="d06fc-158">Definir o formato de intervalo usando a EXCEL JavaScript</span><span class="sxs-lookup"><span data-stu-id="d06fc-158">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
