---
title: Otimização de desempenho da API JavaScript do Excel
description: Otimize o desempenho usando a API JavaScript do Excel
ms.date: 03/28/2018
ms.openlocfilehash: 83150e01a691379f244ce1ce43c190ea32dd170f
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459123"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Otimização de desempenho usando a API JavaScript do Excel

Há várias maneiras de realizar tarefas comuns com a API JavaScript do Excel. Você encontrará diferenças significativas de desempenho entre várias abordagens. Este artigo oferece orientação e códigos de exemplo para mostrar como executar tarefas comuns com eficiência usando a API JavaScript do Excel.

## <a name="minimize-the-number-of-sync-calls"></a>Minimize o número de chamadas sync()

Na API JavaScript do Excel, ```sync()``` é a única operação assíncrona, e ela pode ser lenta em algumas circunstâncias, especialmente no Excel Online. Para otimizar o desempenho, minimize o número de chamadas para ```sync()``` enfileirando o máximo possível de alterações antes de chamá-la.

Confira [Conceitos Básicos - sync()](excel-add-ins-core-concepts.md#sync) para obter exemplos de código que seguem essa prática.

## <a name="minimize-the-number-of-proxy-objects-created"></a>Minimize o número de objetos de proxy criados

Evite criar repetidamente o mesmo objeto de proxy. Em vez disso, se precisar usar um mesmo objeto de proxy em mais de uma operação, crie-o uma única vez, atribua-o a uma variável e use essa variável no código.

```javascript
// BAD: repeated calls to .getRange() to create the same proxy object
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: create the range proxy object once and assign to a variable
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: use a "set" method to immediately set all the properties without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

## <a name="load-necessary-properties-only"></a>Carregue apenas as propriedades necessárias

Na API JavaScript do Excel, você precisa carregar explicitamente as propriedades de um objeto de proxy. Embora você possa carregar todas as propriedades de uma só vez com uma chamada de  ```load()``` vazio, essa abordagem pode afetar o desempenho de maneira significativa. Em vez disso, sugerimos que você carregue apenas as propriedades necessárias, especialmente para os objetos que têm um grande número de propriedades.

Por exemplo, se sua intenção é apenas ler a propriedade **address** de um objeto range, especifique apenas essa propriedade quando chamar o método **load()**:
 
```js
range.load('address');
```
 
É possível chamar o método **load()** de duas maneiras:
 
_Sintaxe:_
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
_Onde:_
 
* `properties` é a lista de propriedades que devem ser carregadas, especificada como sequências de caracteres delimitadas por vírgula ou como uma matriz de nomes. Para obter mais informações, consulte os métodos **load ()** definidos para os objetos na [referência de API JavaScript do Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview).
* `loadOption` especifica um objeto que descreve as opções de seleção, expansão, topo e ignorar. Confira as [opções](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) de carregamento do objeto para saber mais.

Esteja ciente de que algumas das "propriedades" em um objeto podem ter o mesmo nome de outro objeto. Por exemplo, `format` é uma propriedade em um objeto range, mas `format` também é um objeto. Portanto, se você faz uma chamada como `range.load("format")`, isto é equivalente a `range.format.load()`, que é uma chamada load () vazia que pode causar problemas de desempenho, conforme descrito anteriormente. Para evitar isso, seu código deve carregar apenas  "nós folha" em uma árvore de objeto. 

## <a name="suspend-calculation-temporarily"></a>Suspenda os cálculos temporariamente

Se você estiver tentando executar uma operação em um grande número de células (por exemplo, configurar o valor de um objeto range enorme) e não se importar em suspender temporariamente os cálculos no Excel até que a operação seja concluída, recomendamos suspender os cálculos até a chamada da próxima ```context.sync()```.

Confira a documentação de referência do [Objeto Application](https://docs.microsoft.com/javascript/api/excel/excel.application) para obter informações sobre como usar a API ```suspendApiCalculationUntilNextSync()``` para suspender e reativar os cálculos de uma maneira muito conveniente. O código a seguir demonstra como suspender temporariamente o cálculo:

```js
Excel.run(async function(ctx) {
    var app = ctx.workbook.application;
    var sheet = ctx.workbook.worksheets.getItem("sheet1");
    var rangeToSet: Excel.Range;
    var rangeToGet: Excel.Range;
    app.load("calculationMode");
    await ctx.sync();
    // Calculation mode should be "Automatic" by default
    console.log(app.calculationMode);
    
    rangeToSet = sheet.getRange("A1:C1");
    rangeToSet.values = [[1, 2, "=SUM(A1:B1)"]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [1, 2, 3] now
    console.log(rangeToGet.values);

    // Suspending recalc
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with supend recalc
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

## <a name="update-all-cells-in-a-range"></a>Atualize todas as células em um intervalo 

Quando você precisar atualizar todas as células em um intervalo com o mesmo valor ou propriedade, pode ser lento fazer isso por meio de uma matriz bidimensional que atribui o mesmo valor repetidamente, pois essa abordagem exige que o Excel itere todas as células no intervalo para definir cada uma em separado. O Excel tem uma maneira mais eficiente para atualizar todas as células em um intervalo com o mesmo valor ou propriedade.

Se você precisa aplicar o mesmo valor, a mesma formatação de número ou a mesma fórmula para um intervalo de células, é mais eficiente especificar um valor único, em vez de uma matriz de valores. Isso melhorará significativamente o desempenho. Para um exemplo de código que mostra essa abordagem em ação, confira [Conceitos fundamentais - atualizar todas as células em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

Um cenário comum onde você pode aplicar essa abordagem é quando aplica formatos diferentes de números em colunas diferentes da planilha. Nesse caso, você pode simplesmente percorrer as colunas e definir o formato de número em cada coluna com um único valor. Trate cada coluna como um intervalo, conforme mostrado no exemplo de código de [atualização de todas as células em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) .

> [!NOTE]
> Se você estiver usando TypeScript, perceberá um erro de compilação dizendo que um único valor não pode ser atribuído a uma matriz bidimensional. Isso é inevitável, pois os valores *formam* uma matriz bidimensional quando as propriedades são recuperadas e o TypeScript não permite tipos diferentes tipos setter vs getter.  No entanto, uma solução alternativa simples é definir os valores com um sufixo`as any`, por exemplo, `range.values = "hello world" as any`.

## <a name="importing-data-into-tables"></a>Importar dados em tabelas

Ao tentar importar uma grande quantidade de dados em um objeto [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) diretamente (por exemplo, usando `TableRowCollection.add()`), você pode sofrer lentidão. Se você estiver tentando adicionar uma nova tabela, você deve preencher os dados primeiro, atribuindo `range.values` e então chamar `worksheet.tables.add()` para criar a tabela com o intervalo. Se você estiver tentando gravar dados em uma tabela existente, grave os dados em um objeto range via `table.getDataBodyRange()`, e a tabela se expandirá automaticamente. 

Veja um exemplo dessa abordagem:

```js
Excel.run(async (ctx) => {
    var sheet = ctx.workbook.worksheets.getItem("Sheet1");
    // Write the data into the range first 
    var range = sheet.getRange("A1:B3");
    range.values = [["Key", "Value"], ["A", 1], ["B", 2]];

    // Create the table over the range
    var table = sheet.tables.add('A1:B3', true);
    table.name = "Example";
    await ctx.sync();


    // Insert a new row to the table
    table.getDataBodyRange().getRowsBelow(1).values = [["C", 3]];
    // Change a existing row value
    table.getDataBodyRange().getRow(1).values = [["D", 4]];
    await ctx.sync();
})
```

> [!NOTE]
> É possível converter convenientemente um objeto Table em um objeto Range usando o método [Table.convertToRange()](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--).

## <a name="enable-and-disable-events"></a>Ative e desative eventos

O desempenho de um suplemento pode ser melhorado com a desativação de eventos. Confira um exemplo de código mostrando como ativar e desativar eventos no artigo [Trabalhando com eventos](excel-add-ins-events.md#enable-and-disable-events) .

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Conceitos avançados de programação com a API JavaScript do Excel](excel-add-ins-advanced-concepts.md)
- [Especificação aberta da API JavaScript do Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Objeto Worksheet Functions (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.functions)
