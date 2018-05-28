---
title: Otimiza??o de desempenho da API JavaScript do Excel
description: Otimize o desempenho usando a API JavaScript do Excel
ms.date: 03/28/2018
ms.openlocfilehash: dabbb69f8dee0df782a265edcfdfb1c89894e915
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Otimiza??o de desempenho usando a API JavaScript do Excel

Existem v?rias maneiras de executar tarefas comuns com a API JavaScript do Excel. Voc? encontrar? diferen?as de desempenho significativas entre as diferentes abordagens. Este artigo fornece diretrizes e exemplos de c?digo para mostrar como executar tarefas comuns com efici?ncia usando a API JavaScript do Excel.

## <a name="minimize-the-number-of-sync-calls"></a>Minimizar o n?mero de chamadas sync()

Na API JavaScript do Excel, ```sync()``` ? a ?nica opera??o ass?ncrona e pode ser lenta em determinadas circunst?ncias, especialmente no Excel Online. Para otimizar o desempenho, minimize o n?mero de chamadas para ```sync()```, colocando em fila o maior n?mero poss?vel de altera??es antes de cham?-la.

Veja [Conceitos B?sicos - sync()](excel-add-ins-core-concepts.md#sync) para obter exemplos de c?digo que seguem essa pr?tica.

## <a name="minimize-the-number-of-proxy-objects-created"></a>Minimizar o n?mero de objetos de proxy criados

Evite criar repetidamente o mesmo objeto de proxy. Em vez disso, se precisar usar um mesmo objeto de proxy em mais de uma opera??o, crie-o uma ?nica vez, atribua-o a uma vari?vel e use essa vari?vel no c?digo.

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

## <a name="load-necessary-properties-only"></a>Carregar apenas as propriedades necess?rias

Na API JavaScript do Excel, ? preciso carregar explicitamente as propriedades de um objeto de proxy. Embora seja poss?vel carregar todas as propriedades de uma s? vez com uma chamada vazia de ```load()```, essa abordagem pode sobrecarregar significativamente o desempenho. Em vez disso, sugerimos carregar apenas as propriedades necess?rias, especialmente para aqueles objetos que possuem um grande n?mero de propriedades.

Por exemplo, se sua inten??o ? apenas ler a propriedade **address** de um objeto de intervalo, especifique somente essa propriedade ao chamar o m?todo **load()**:
 
```js
range.load('address');
```
 
? poss?vel chamar o m?todo **load()** de duas maneiras:
 
_Sintaxe:_
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
_Onde:_
 
* `properties` ? a lista de propriedades a serem carregadas, especificadas como cadeias de caracteres delimitadas por v?rgula ou como uma matriz de nomes. Para saber mais, veja os m?todos **load()** definidos para objetos na [refer?ncia da API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).
* `loadOption` especifica um objeto que descreve as op??es de sele??o, expans?o, topo e ignorar. Confira as [op??es](https://dev.office.com/reference/add-ins/excel/loadoption) de carregamento de objetos para saber mais.

Observe que algumas das ?propriedades? sob um objeto podem ter o mesmo nome que outro objeto. Por exemplo, `format` ? uma propriedade do objeto de intervalo, mas `format` em si tamb?m ? um objeto. Assim, se voc? fizer uma chamada como `range.load("format")`, isso ser? equivalente a `range.format.load()`, que ? uma chamada vazia de load() que pode causar problemas de desempenho, conforme descrito anteriormente. Para evitar isso, o c?digo deve carregar apenas os "n?s folha" em uma ?rvore de objetos. 

## <a name="suspend-calculation-temporarily"></a>Suspender os c?lculos temporariamente

Se estiver tentando executar uma opera??o em um grande n?mero de c?lulas (por exemplo, configurando o valor de um objeto de intervalo enorme) e n?o se importar em suspender temporariamente os c?lculos no Excel enquanto a opera??o ? conclu?da, recomendamos suspender os c?lculos at? a chamada da pr?xima ```context.sync()```.

Veja a documenta??o de refer?ncia do [Objeto de Aplicativo](https://dev.office.com/reference/add-ins/excel/application) para obter informa??es sobre como usar a ```suspendApiCalculationUntilNextSync()``` API para suspender e reativar os c?lculos de forma muito conveniente. O seguinte c?digo demonstra como suspender os c?lculos temporariamente:

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

## <a name="update-all-cells-in-a-range"></a>Atualizar todas as c?lulas em um intervalo 

Quando voc? precisa atualizar todas as c?lulas de um intervalo com um mesmo valor ou propriedade, poder? ser lento fazer isso por meio de uma matriz bidimensional que especifica repetidamente o mesmo valor, j? que essa abordagem exige que o Excel execute uma itera??o em todas as c?lulas do intervalo para definir cada uma separadamente. O Excel tem uma maneira mais eficiente para atualizar todas as c?lulas de um intervalo com um mesmo valor ou propriedade.

Se voc? precisar aplicar o mesmo valor, formato de n?mero ou f?rmula a um intervalo de c?lulas, ser? mais eficiente especificar um ?nico valor em vez de uma matriz de valores. Isso aumentar? significativamente o desempenho. Para obter um exemplo de c?digo que mostre essa abordagem em a??o, veja [Conceitos principais - Atualizar todas as c?lulas em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

Um cen?rio comum em que voc? pode aplicar essa abordagem ? ao definir formatos num?ricos diferentes em colunas diferentes em uma planilha. Nesse caso, voc? pode simplesmente iterar pelas colunas e definir o formato num?rico em cada coluna com um valor ?nico. Trate cada coluna como um intervalo, conforme mostrado no exemplo de c?digo em [Atualizar todas as c?lulas em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

> [!NOTE]
> Se estiver usando o TypeScript, voc? observar? um erro de compila??o informando que um valor ?nico n?o pode ser definido como uma matriz 2D.  Isso ? inevit?vel, j? que os valores *s?o* uma matriz 2D ao recuperar as propriedades e o TypeScript n?o permite tipos diferentes de setter versus getter.  No entanto, uma solu??o simples ? definir os valores com um sufixo `as any`, por exemplo, `range.values = "hello world" as any`.

## <a name="importing-data-into-tables"></a>Importa??o de dados em tabelas

Ao tentar importar uma enorme quantidade de dados diretamente para um objeto [Table](https://dev.office.com/reference/add-ins/excel/table) (por exemplo, usando `TableRowCollection.add()`), o desempenho poder? ser mais lento. Se voc? estiver tentando adicionar uma nova tabela, dever? preencher os dados primeiro, definindo `range.values` e chamando `worksheet.tables.add()` para criar uma tabela no intervalo. Se voc? estiver tentando gravar dados em uma tabela existente, grave-os em um objeto de intervalo por meio de `table.getDataBodyRange()` e a tabela ser? expandida automaticamente. 

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
> ? poss?vel converter convenientemente um objeto Table em um objeto Range usando o m?todo [Table.convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange).

## <a name="see-also"></a>Confira tamb?m

- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Conceitos avan?ados da API JavaScript do Excel](excel-add-ins-advanced-concepts.md)
- [Especifica??o para abrir a API JavaScript do Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [Objeto de fun??es de planilha (API JavaScript para Excel)](https://dev.office.com/reference/add-ins/excel/functions)
