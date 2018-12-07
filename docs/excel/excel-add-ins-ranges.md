---
title: Trabalhar com intervalos usando a API JavaScript do Excel
description: ''
ms.date: 12/04/2018
ms.openlocfilehash: 4a6e0014da82956b15e11e2739f6f58fb82d5030
ms.sourcegitcommit: e2ba9d7210c921d068f40d9f689314c73ad5ab4a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/05/2018
ms.locfileid: "27156604"
---
# <a name="work-with-ranges-using-the-excel-javascript-api"></a>Trabalhar com intervalos usando a API JavaScript do Excel

Este artigo fornece exemplos de código que mostram como executar tarefas comuns com intervalos usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos que o objeto **Range** suporta, confira [Objeto Range (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).

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

### <a name="conditional-formatting-of-ranges"></a>Formatação condicional de intervalos

Os intervalos podem ter formatos aplicados a células individuais baseadas em condições. Para saber mais sobre isso, confira [Aplicar a formatação condicional a intervalos do Excel](excel-add-ins-conditional-formatting.md).

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a>Trabalhar com datas usando o plug-in Moment-MSDate

A [biblioteca Moment do JavaScript](https://momentjs.com/) fornece uma maneira conveniente de usar datas e carimbos de data e hora. O [plug-in Moment-MSDate](https://www.npmjs.com/package/moment-msdate) converte o formato de momentos em um formato mais apropriado para o Excel. Este é o mesmo formato que a [função NOW](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) retorna.

O código a seguir mostra como definir o intervalo em ** B4 ** para o carimbo de data/hora de um momento:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    
    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();
    
    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];
    
    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];
    
    return context.sync();
}).catch(errorHandlerFunction);
```

É uma técnica semelhante para retirar a data da célula e convertê-la em um momento ou outro formato, conforme demonstrado no código a seguir:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");
        
    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp 
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

Seu suplemento terá que formatar os intervalos para exibir as datas em um formato mais legível. O exemplo de `"[$-409]m/d/yy h:mm AM/PM;@"` exibe a hora como "3/12/18 15:57". Para obter mais informações sobre formatos de números de data e hora, confira as "Diretrizes para formatos de data e hora" no artigo [Diretrizes de revisão para personalizar um formato de número](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).

## <a name="copy-and-paste"></a>Copiar e colar

> [!NOTE]
> A função copyFrom no momento só está disponível na versão prévia pública (beta). Para usar esse recurso, você deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> Se você estiver usando o TypeScript ou se seu editor de código usar arquivos de definição de tipo do TypeScript do IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

A função de copyFrom do intervalo replica o comportamento de copiar e colar da IU do Excel. O objeto de intervalo para o qual a função copyFrom é chamada é o destino. A fonte a ser copiada é passada como um intervalo ou um endereço de cadeia de caracteres que representa um intervalo. O exemplo a seguir copia dados de **A1:E1** para o intervalo que começa em **G1** (que acaba sendo colado em **G1:K1**).

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

Range.copyFrom tem três parâmetros opcionais.

```ts
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
``` 

`copyType` especifica quais dados são copiados da origem para o destino. 
`“Formulas”` transfere as fórmulas nas células de origem e preserva o posicionamento relativo dos intervalos dessas fórmulas. As entradas que não sejam uma fórmula são copiadas no seu estado original. 
`“Values”` copia os valores dos dados e, no caso de fórmulas, o resultado da fórmula. 
`“Formats”` copia a formatação do intervalo incluindo a fonte, cor e outras configurações de formato, mas nenhum valor. 
`”All”` (a opção padrão) copia ambos os dados e formatação, preservando as fórmulas das células, caso elas sejam encontradas.

`skipBlanks` define se as células em branco são copiadas para o destino. Quando for verdadeiro, `copyFrom` ignora células em branco no intervalo de origem. As células ignoradas não substituem os dados existentes de suas células correspondentes no intervalo de destino. O padrão é false.

O exemplo de código e as imagens a seguir demonstram esse comportamento em um cenário simples. 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

*Antes da função precedente ter sido executada.*

![Os dados no Excel antes do método de copiar do intervalo foram executados.](../images/excel-range-copyfrom-skipblanks-before.png)

*Após a função precedente ter sido executada.*

![Os dados no Excel após o método de copiar do intervalo foram executados.](../images/excel-range-copyfrom-skipblanks-after.png)

`transpose` determina se os dados são transpostos, ou seja, suas linhas e colunas são alternadas para o local de origem. Um intervalo transposto invertido na diagonal principal, portanto as linhas **1**, **2** e **3** se tornarão as colunas **A**, **B** e **C**. 


## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)

