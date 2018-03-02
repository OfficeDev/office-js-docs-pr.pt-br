---
title: Trabalhar com intervalos usando a API JavaScript do Excel
description: ''
ms.date: 12/04/2017
---


# <a name="work-with-ranges-using-the-excel-javascript-api"></a>Trabalhar com intervalos usando a API JavaScript do Excel

Este artigo fornece exemplos de código que mostram como executar tarefas comuns com intervalos usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos que o objeto **Range** suporta, confira [Objeto Range (API JavaScript para Excel)](https://dev.office.com/reference/add-ins/excel/range).

## <a name="get-a-range"></a>Obter um intervalo

Os exemplos a seguir mostram diferentes maneiras de obter uma referência a um intervalo em uma planilha.

### <a name="get-range-by-address"></a>Obter intervalo por endereço

O exemplo de código a seguir obtém o intervalo com o endereço **B2:B5** da planilha chamada **Amostra**, carrega sua propriedade **address** e grava uma mensagem no console.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-range-by-name"></a>Obter intervalo por nome

O exemplo de código a seguir obtém o intervalo chamado **MyRange** da planilha chamada **Amostra**, carrega sua propriedade **address** e grava uma mensagem no console.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-used-range"></a>Obter intervalo usado

O exemplo de código a seguir obtém o intervalo usado da planilha chamada **Amostra**, carrega sua propriedade **address** e grava uma mensagem no console. O intervalo usado é o menor intervalo que abrange todas as células na planilha que têm um valor ou uma formatação atribuída a elas. Se a planilha inteira estiver em branco, o método **getUsedRange()** retornará um intervalo que consiste apenas na célula superior esquerda da planilha.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-entire-range"></a>Obter intervalo inteiro

O exemplo de código a seguir obtém todo o intervalo da planilha chamada **Amostra**, carrega sua propriedade **address** e grava uma mensagem no console.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="insert-a-range-of-cells"></a>Inserir um intervalo de células

O exemplo de código a seguir insere um intervalo de células no local **B4:E4** e desloca outras células para baixo a fim de fornecer espaço para as novas células.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);
    
    return context.sync();
}).catch(errorHandlerFunction);
```

**Dados antes da inserção do intervalo**

![Dados no Excel antes da inserção do intervalo](../images/excel-ranges-start.png)

**Dados após a inserção do intervalo**

![Dados no Excel após a inserção do intervalo](../images/excel-ranges-after-insert.png)

## <a name="clear-a-range-of-cells"></a>Limpar um intervalo de células

O exemplo de código a seguir limpa todo o conteúdo e a formatação das células no intervalo **E2:E5**.  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Dados antes da limpeza do intervalo**

![Dados no Excel antes da limpeza do intervalo](../images/excel-ranges-start.png)

**Dados após a limpeza do intervalo**

![Dados no Excel após a limpeza do intervalo](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a>Excluir um intervalo de células

O exemplo de código a seguir exclui as células no intervalo **B4:E4** e desloca outras células para cima a fim de preencher o espaço deixado pelas células excluídas.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

**Dados antes da exclusão do intervalo**

![Dados no Excel antes da exclusão do intervalo](../images/excel-ranges-start.png)

**Dados após a exclusão do intervalo**

![Dados no Excel após a exclusão do intervalo](../images/excel-ranges-after-delete.png)

## <a name="set-the-selected-range"></a>Definir o intervalo selecionado

O exemplo de código a seguir seleciona o intervalo **B2:E6** na planilha ativa.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Intervalo selecionado B2:E6**

![Intervalo selecionado no Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a>Obter o intervalo selecionado

O exemplo de código a seguir obtém o intervalo selecionado, carrega sua propriedade **address** e grava uma mensagem no console. 

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-values-or-formulas"></a>Definir valores ou fórmulas

Os exemplos a seguir mostram como definir valores e fórmulas para uma única célula ou um intervalo de células.

### <a name="set-value-for-a-single-cell"></a>Definir valor para uma única célula

O exemplo de código a seguir define o valor da célula **C3** como "5" e, em seguida, define a largura das colunas para melhor ajustar os dados.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("C3");
    range.values = [[ 5 ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Dados antes da atualização do valor da célula**

![Dados no Excel antes da atualização do valor da célula](../images/excel-ranges-set-start.png)

**Dados após a atualização do valor da célula**

![Dados no Excel após a atualização do valor da célula](../images/excel-ranges-set-cell-value.png)

### <a name="set-values-for-a-range-of-cells"></a>Definir valores para um intervalo de células

O exemplo de código a seguir define valores das células no intervalo **B5:D5** e, em seguida, define a largura das colunas para melhor ajustar os dados.

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

**Dados antes da atualização dos valores da célula**

![Dados no Excel antes da atualização dos valores da célula](../images/excel-ranges-set-start.png)

**Dados após a atualização dos valores da célula**

![Dados no Excel após a atualização dos valores da célula](../images/excel-ranges-set-cell-values.png)

### <a name="set-formula-for-a-single-cell"></a>Definir fórmula para uma única célula

O exemplo de código a seguir define uma fórmula para a célula **E3** e, em seguida, define a largura das colunas para melhor ajustar os dados.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("E3");
    range.formulas = [[ "=C3 * D3" ]];
    range.format.autofitColumns();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Dados antes da definição da fórmula da célula**

![Dados no Excel antes da definição da fórmula da célula](../images/excel-ranges-start-set-formula.png)

**Dados após a definição da fórmula da célula**

![Dados no Excel após a definição da fórmula da célula](../images/excel-ranges-set-formula.png)

### <a name="set-formulas-for-a-range-of-cells"></a>Definir fórmulas para um intervalo de células

O exemplo de código a seguir define fórmulas para células no intervalo **E2:E6** e, em seguida, define a largura das colunas para melhor ajustar os dados.

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

**Dados antes da definição das fórmulas da célula**

![Dados no Excel antes da definição das fórmulas da célula](../images/excel-ranges-start-set-formula.png)

**Dados após a definição das fórmulas da célula**

![Dados no Excel após a definição das fórmulas da célula](../images/excel-ranges-set-formulas.png)

## <a name="get-values-text-or-formulas"></a>Obter valores, texto ou fórmulas

Estes exemplos mostram como obter valores, texto e fórmulas de um intervalo de células.

### <a name="get-values-from-a-range-of-cells"></a>Obter valores de um intervalo de células

O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega sua propriedade **values** e grava os valores no console. A propriedade **values** de um intervalo especifica os novos valores brutos que as células contêm. Mesmo que algumas células em um intervalo contenham fórmulas, a propriedade **values** do intervalo especifica os valores brutos para essas células, não alguma das fórmulas.

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

**Dados no intervalo (valores na coluna E são um resultado de fórmulas)**

![Dados no Excel após a definição das fórmulas da célula](../images/excel-ranges-set-formulas.png)

**range.values (conforme registrado em log no console pelo exemplo de código acima)**

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

O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega sua propriedade **text** e o grava no console.  A propriedade **text** de um intervalo especifica os valores de exibição para as células no intervalo. Mesmo que algumas células em um intervalo contenham fórmulas, a propriedade **text** do intervalo especifica os valores de exibição para essas células, não alguma das fórmulas.

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

**Dados no intervalo (valores na coluna E são um resultado de fórmulas)**

![Dados no Excel após a definição das fórmulas da célula](../images/excel-ranges-set-formulas.png)

**range.text (conforme registrado em log no console pelo exemplo de código acima)**

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

O exemplo de código a seguir obtém o intervalo **B2:E6**, carrega sua propriedade **formulas** e o grava no console.  A propriedade **formulas** de um intervalo especifica as fórmulas para células no intervalo que contêm fórmulas e os valores brutos para células no intervalo que não contêm fórmulas.

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

**Dados no intervalo (valores na coluna E são um resultado de fórmulas)**

![Dados no Excel após a definição das fórmulas da célula](../images/excel-ranges-set-formulas.png)

**range.formulas (conforme registrado em log no console pelo exemplo de código acima)**

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

## <a name="set-range-format"></a>Definir formato do intervalo

Os exemplos a seguir mostram como definir a cor da fonte, a cor de preenchimento e o formato de número para células em um intervalo.

### <a name="set-font-color-and-fill-color"></a>Definir cor da fonte e cor de preenchimento

O exemplo de código a seguir define a cor da fonte e a cor de preenchimento para células no intervalo **B2:E2**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";;
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

**Dados no intervalo antes da definição da cor da fonte e da cor de preenchimento**

![Dados no Excel antes da definição do formato](../images/excel-ranges-format-before.png)

**Dados no intervalo após a definição da cor da fonte e da cor de preenchimento**

![Dados no Excel após a definição do formato](../images/excel-ranges-format-font-and-fill.png)

### <a name="set-number-format"></a>Definir formato de número

O exemplo de código a seguir define o formato de número para as células no intervalo **D3:E5**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

**Dados no intervalo antes da definição do formato de número**

![Dados no Excel antes da definição do formato](../images/excel-ranges-format-font-and-fill.png)

**Dados no intervalo após a definição do formato de número**

![Dados no Excel após a definição do formato](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a>Veja também

- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Objeto Range (API JavaScript para Excel)](https://dev.office.com/reference/add-ins/excel/range)
