---
title: Otimização de desempenho da API JavaScript do Excel
description: Otimize o desempenho usando a API JavaScript do Excel
ms.date: 03/28/2018
ms.openlocfilehash: dabbb69f8dee0df782a265edcfdfb1c89894e915
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437406"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Otimização de desempenho usando a API JavaScript do Excel

Existem várias maneiras de executar tarefas comuns com a API JavaScript do Excel. Você encontrará diferenças de desempenho significativas entre as diferentes abordagens. Este artigo fornece diretrizes e exemplos de código para mostrar como executar tarefas comuns com eficiência usando a API JavaScript do Excel.

## <a name="minimize-the-number-of-sync-calls"></a>Minimizar o número de chamadas sync()

Na API JavaScript do Excel, ```sync()``` é a única operação assíncrona e pode ser lenta em determinadas circunstâncias, especialmente no Excel Online. Para otimizar o desempenho, minimize o número de chamadas para ```sync()```, colocando em fila o maior número possível de alterações antes de chamá-la.

Veja [Conceitos Básicos - sync()](excel-add-ins-core-concepts.md#sync) para obter exemplos de código que seguem essa prática.

## <a name="minimize-the-number-of-proxy-objects-created"></a>Minimizar o número de objetos de proxy criados

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

## <a name="load-necessary-properties-only"></a>Carregar apenas as propriedades necessárias

Na API JavaScript do Excel, é preciso carregar explicitamente as propriedades de um objeto de proxy. Embora seja possível carregar todas as propriedades de uma só vez com uma chamada vazia de ```load()```, essa abordagem pode sobrecarregar significativamente o desempenho. Em vez disso, sugerimos carregar apenas as propriedades necessárias, especialmente para aqueles objetos que possuem um grande número de propriedades.

Por exemplo, se sua intenção é apenas ler a propriedade **address** de um objeto de intervalo, especifique somente essa propriedade ao chamar o método **load()**:
 
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
 
* `properties` é a lista de propriedades a serem carregadas, especificadas como cadeias de caracteres delimitadas por vírgula ou como uma matriz de nomes. Para saber mais, veja os métodos **load()** definidos para objetos na [referência da API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).
* `loadOption` especifica um objeto que descreve as opções de seleção, expansão, topo e ignorar. Confira as [opções](https://dev.office.com/reference/add-ins/excel/loadoption) de carregamento de objetos para saber mais.

Observe que algumas das “propriedades” sob um objeto podem ter o mesmo nome que outro objeto. Por exemplo, `format` é uma propriedade do objeto de intervalo, mas `format` em si também é um objeto. Assim, se você fizer uma chamada como `range.load("format")`, isso será equivalente a `range.format.load()`, que é uma chamada vazia de load() que pode causar problemas de desempenho, conforme descrito anteriormente. Para evitar isso, o código deve carregar apenas os "nós folha" em uma árvore de objetos. 

## <a name="suspend-calculation-temporarily"></a>Suspender os cálculos temporariamente

Se estiver tentando executar uma operação em um grande número de células (por exemplo, configurando o valor de um objeto de intervalo enorme) e não se importar em suspender temporariamente os cálculos no Excel enquanto a operação é concluída, recomendamos suspender os cálculos até a chamada da próxima ```context.sync()```.

Veja a documentação de referência do [Objeto de Aplicativo](https://dev.office.com/reference/add-ins/excel/application) para obter informações sobre como usar a ```suspendApiCalculationUntilNextSync()``` API para suspender e reativar os cálculos de forma muito conveniente. O seguinte código demonstra como suspender os cálculos temporariamente:

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

## <a name="update-all-cells-in-a-range"></a>Atualizar todas as células em um intervalo 

Quando você precisa atualizar todas as células de um intervalo com um mesmo valor ou propriedade, poderá ser lento fazer isso por meio de uma matriz bidimensional que especifica repetidamente o mesmo valor, já que essa abordagem exige que o Excel execute uma iteração em todas as células do intervalo para definir cada uma separadamente. O Excel tem uma maneira mais eficiente para atualizar todas as células de um intervalo com um mesmo valor ou propriedade.

Se você precisar aplicar o mesmo valor, formato de número ou fórmula a um intervalo de células, será mais eficiente especificar um único valor em vez de uma matriz de valores. Isso aumentará significativamente o desempenho. Para obter um exemplo de código que mostre essa abordagem em ação, veja [Conceitos principais - Atualizar todas as células em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

Um cenário comum em que você pode aplicar essa abordagem é ao definir formatos numéricos diferentes em colunas diferentes em uma planilha. Nesse caso, você pode simplesmente iterar pelas colunas e definir o formato numérico em cada coluna com um valor único. Trate cada coluna como um intervalo, conforme mostrado no exemplo de código em [Atualizar todas as células em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

> [!NOTE]
> Se estiver usando o TypeScript, você observará um erro de compilação informando que um valor único não pode ser definido como uma matriz 2D.  Isso é inevitável, já que os valores *são* uma matriz 2D ao recuperar as propriedades e o TypeScript não permite tipos diferentes de setter versus getter.  No entanto, uma solução simples é definir os valores com um sufixo `as any`, por exemplo, `range.values = "hello world" as any`.

## <a name="importing-data-into-tables"></a>Importação de dados em tabelas

Ao tentar importar uma enorme quantidade de dados diretamente para um objeto [Table](https://dev.office.com/reference/add-ins/excel/table) (por exemplo, usando `TableRowCollection.add()`), o desempenho poderá ser mais lento. Se você estiver tentando adicionar uma nova tabela, deverá preencher os dados primeiro, definindo `range.values` e chamando `worksheet.tables.add()` para criar uma tabela no intervalo. Se você estiver tentando gravar dados em uma tabela existente, grave-os em um objeto de intervalo por meio de `table.getDataBodyRange()` e a tabela será expandida automaticamente. 

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
> É possível converter convenientemente um objeto Table em um objeto Range usando o método [Table.convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange).

## <a name="see-also"></a>Confira também

- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Conceitos avançados da API JavaScript do Excel](excel-add-ins-advanced-concepts.md)
- [Especificação para abrir a API JavaScript do Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Objeto de funções de planilha (API JavaScript para Excel)](https://dev.office.com/reference/add-ins/excel/functions)
