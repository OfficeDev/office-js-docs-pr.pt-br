---
title: Definir e obter valores de intervalo, texto ou fórmulas usando a EXCEL JavaScript
description: Saiba como usar a EXCEL JavaScript para definir e obter valores de intervalo, texto ou fórmulas.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 3b27a1596259c5950cdd41999f00c30c3bd0f4e0
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745376"
---
# <a name="set-and-get-range-values-text-or-formulas-using-the-excel-javascript-api"></a>Definir e obter valores de intervalo, texto ou fórmulas usando a EXCEL JavaScript

Este artigo fornece exemplos de código que configuram e obteram valores de intervalo, texto ou fórmulas com a API JavaScript Excel JavaScript. Para ver a lista completa de propriedades e métodos compatíveis `Range` com o objeto, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-values-or-formulas"></a>Definir valores ou fórmulas

Os exemplos de código a seguir configuram valores e fórmulas para uma única célula ou um intervalo de células.

### <a name="set-value-for-a-single-cell"></a>Definir valor para uma única célula

O exemplo de código a seguir define o valor da célula **C3** como "5" e, em seguida, define a largura das colunas para melhor ajustar os dados.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    await context.sync();
});
```

#### <a name="data-before-cell-value-is-updated"></a>Dados antes da atualização do valor da célula

![Dados em Excel antes que o valor da célula seja atualizado.](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-value-is-updated"></a>Dados após a atualização do valor da célula

![Dados em Excel depois que o valor da célula é atualizado.](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a>Definir valores para um intervalo de células

O exemplo de código a seguir define valores das células no intervalo **B5:D5** e, em seguida, define a largura das colunas para melhor ajustar os dados.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let data = [
        ["Potato Chips", 10, 1.80],
    ];

    let range = sheet.getRange("B5:D5");
    range.values = data;
    range.format.autofitColumns();

    await context.sync();
});
```

#### <a name="data-before-cell-values-are-updated"></a>Dados antes da atualização dos valores da célula

![Dados em Excel antes que os valores da célula sejam atualizados.](../images/excel-ranges-set-start.png)

#### <a name="data-after-cell-values-are-updated"></a>Dados após a atualização dos valores da célula

![Dados em Excel depois que os valores da célula são atualizados.](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a>Definir fórmula para uma única célula

O exemplo de código a seguir define uma fórmula para a célula **E3** e, em seguida, define a largura das colunas para melhor ajustar os dados.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    await context.sync();
});
```

#### <a name="data-before-cell-formula-is-set"></a>Dados antes da definição da fórmula da célula

![Dados em Excel antes que a fórmula da célula seja definida.](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formula-is-set"></a>Dados após a definição da fórmula da célula

![Dados na Excel depois que a fórmula da célula é definida.](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a>Definir fórmulas para um intervalo de células

O exemplo de código a seguir define fórmulas para células no intervalo **E2:E6** e, em seguida, define a largura das colunas para melhor ajustar os dados.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let data = [
        ["=C3 * D3"],
        ["=C4 * D4"],
        ["=C5 * D5"],
        ["=SUM(E3:E5)"]
    ];

    let range = sheet.getRange("E3:E6");
    range.formulas = data;
    range.format.autofitColumns();

    await context.sync();
});
```

#### <a name="data-before-cell-formulas-are-set"></a>Dados antes da definição das fórmulas da célula

![Dados em Excel antes que as fórmulas da célula sejam definidas.](../images/excel-ranges-start-set-formula.png)

#### <a name="data-after-cell-formulas-are-set"></a>Dados após a definição das fórmulas da célula

![Dados em Excel depois que as fórmulas da célula são definidas.](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a>Obter valores, texto ou fórmulas

Esses exemplos de código obterão valores, texto e fórmulas de um intervalo de células.

### <a name="get-values-from-a-range-of-cells"></a>Obter valores de um intervalo de células

O exemplo de código a seguir obtém o intervalo **B2:E6**, `values` carrega sua propriedade e grava os valores no console. A `values` propriedade de um intervalo especifica os valores brutos que as células contêm. Mesmo que algumas células em um intervalo contenham fórmulas, `values` a propriedade do intervalo especifica os valores brutos dessas células, não qualquer uma das fórmulas.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E6");
    range.load("values");
    await context.sync();

    console.log(JSON.stringify(range.values, null, 4));
});
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a>Dados no intervalo (valores na coluna E são um resultado de fórmulas)

![Dados em Excel depois que as fórmulas da célula são definidas.](../images/excel-ranges-set-formulas.png)

#### <a name="rangevalues-as-logged-to-the-console-by-the-code-sample-above"></a>range.values (conforme registrado em log no console pelo exemplo de código acima)

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

### <a name="get-text-from-a-range-of-cells"></a>Obter texto de um intervalo de células

O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega `text` sua propriedade e grava-a no console. A `text` propriedade de um intervalo especifica os valores de exibição para células no intervalo. Mesmo que algumas células em um intervalo contenham fórmulas, `text` a propriedade do intervalo especifica os valores de exibição dessas células, não qualquer uma das fórmulas.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E6");
    range.load("text");
    await context.sync();

    console.log(JSON.stringify(range.text, null, 4));
});
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a>Dados no intervalo (valores na coluna E são um resultado de fórmulas)

![Dados em Excel depois que as fórmulas da célula são definidas.](../images/excel-ranges-set-formulas.png)

#### <a name="rangetext-as-logged-to-the-console-by-the-code-sample-above"></a>range.text (conforme registrado em log no console pelo exemplo de código acima)

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

### <a name="get-formulas-from-a-range-of-cells"></a>Obter fórmulas de um intervalo de células

O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega `formulas` sua propriedade e grava-a no console. A `formulas` propriedade de um intervalo especifica as fórmulas para células no intervalo que contêm fórmulas e os valores brutos para células no intervalo que não contêm fórmulas.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E6");
    range.load("formulas");
    await context.sync();

    console.log(JSON.stringify(range.formulas, null, 4));
});
```

#### <a name="data-in-range-values-in-column-e-are-a-result-of-formulas"></a>Dados no intervalo (valores na coluna E são um resultado de fórmulas)

![Dados em Excel depois que as fórmulas da célula são definidas.](../images/excel-ranges-set-formulas.png)

#### <a name="rangeformulas-as-logged-to-the-console-by-the-code-sample-above"></a>range.formulas (conforme registrado em log no console pelo exemplo de código acima)

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

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Definir e obter intervalos usando a EXCEL JavaScript](excel-add-ins-ranges-set-get.md)
- [Definir o formato de intervalo usando a EXCEL JavaScript](excel-add-ins-ranges-set-format.md)
