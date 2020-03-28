---
title: Otimização de desempenho do da API JavaScript do Excel
description: Otimizar o desempenho usando as API JavaScript do Excel
ms.date: 03/27/2020
localization_priority: Normal
ms.openlocfilehash: a202776569cdfc31a1221e3de1a356f0dafa2bfb
ms.sourcegitcommit: 559a7e178e84947e830cc00dfa01c5c6e398ddc2
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2020
ms.locfileid: "43030828"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a>Otimização de desempenho usando a API JavaScript do Excel

Existem várias maneiras de executar tarefas comuns com a API JavaScript do Excel. Você encontrará diferenças significativas de desempenho entre várias abordagens. Este artigo fornece orientações e amostras de código para mostrar como realizar tarefas comuns com eficiência usando as API JavaScript do Excel.

## <a name="minimize-the-number-of-sync-calls"></a>Minimizar o número de chamadas sync()

Na API do JavaScript do Excel, ```sync()``` é a única operação assíncrona e pode ser lenta em algumas circunstâncias, especialmente no Excel Online na Web. Para otimizar o desempenho, minimize o número de chamadas para ```sync()```, enfileirando o maior número possível de alterações antes de chamá-lo.

Ver [Principais conceitos - sync()](excel-add-ins-core-concepts.md#sync) para as amostras de código que seguem esta prática.

## <a name="minimize-the-number-of-proxy-objects-created"></a>Minimizar o número de objetos proxy criados

Evite criar repetidamente o mesmo objeto proxy. Em vez disso, se você precisar do mesmo objeto proxy para mais de uma operação, crie-o uma vez e o atribua a uma variável, em seguida, use essa variável no seu código.

```js
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

## <a name="load-necessary-properties-only"></a>Carregar propriedades necessárias

Na API JavaScript do Excel, você precisa explicitamente carregar as propriedades de um objeto de proxy. Embora você seja capaz de carregar todas as propriedades de uma vez com uma ```load()``` chamada vazia, essa abordagem pode ter uma sobrecarga de desempenho significativa. Em vez disso, é recomendável apenas carregar as propriedades necessárias, especialmente para os objetos que têm um grande número de propriedades.

Por exemplo, se você pretende apenas ler a `address` propriedade de um objeto Range, especifique somente essa propriedade quando chamar o `load()` método:

```js
range.load('address');
```

Você pode chamar `load()` método de qualquer uma das seguintes maneiras:

_Sintaxe:_

```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```

_Onde:_

* `properties` é a lista de propriedades para carregar, especificadas como cadeias de caracteres delimitadas por vírgula ou como uma matriz de nomes. Para obter mais informações, consulte `load()` os métodos definidos para objetos na [referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md).
* `loadOption` especifica um objeto que descreve as opções de seleção, expansão, topo e ignorar. Confira as [opções](/javascript/api/office/officeextension.loadoption) de carregamento do objeto para saber mais.

Observe que algumas das "Propriedades" em um objeto podem ter o mesmo nome de outro objeto. Por exemplo, `format` é uma propriedade dentro do objeto de intervalo, mas `format` também é um objeto. Portanto, se você fizer uma chamada, como `range.load("format")`, isso equivale a`range.format.load()`, que é uma chamada load vazia () que pode causar problemas de desempenho, conforme descrito anteriormente. Para evitar isso, o código só deve carregar os "nós folha" em uma árvore de objetos.

## <a name="suspend-excel-processes-temporarily"></a>Suspender temporariamente os processos do Excel

O Excel tem várias tarefas em segundo plano reagindo à entrada de usuários e seu suplemento. Alguns desses processos do Excel podem ser controlado para obter o benefício de desempenho. Isso é útil principalmente quando o suplemento lida com grandes conjuntos de dados.

### <a name="suspend-calculation-temporarily"></a>Suspender os cálculos temporariamente

Se você estiver tentando executar uma operação em um grande número de células (por exemplo, definindo o valor do objeto de um grande intervalo) e não se importar em suspender o cálculo no Excel temporariamente enquanto a operação for concluída, é recomendável que você suspenda o cálculo até o próximo `context.sync()` ser chamado.

Ver a documentação de referência [objeto de aplicativo](/javascript/api/excel/excel.application) para saber mais sobre como usar a API`suspendApiCalculationUntilNextSync()`para suspender e reativar cálculos de maneira muito fácil. O código a seguir demonstra como suspender temporariamente um cálculo:

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

    // Suspending recalculation
    app.suspendApiCalculationUntilNextSync();
    rangeToSet = sheet.getRange("A1:B1");
    rangeToSet.values = [[10, 20]];
    rangeToGet = sheet.getRange("A1:C1");
    rangeToGet.load("values");
    app.load("calculationMode");
    await ctx.sync();
    // Range value should be [10, 20, 3] when we load the property, because calculation is suspended at that point
    console.log(rangeToGet.values);
    // Calculation mode should still be "Automatic" even with suspend recalculation
    console.log(app.calculationMode);

    rangeToGet.load("values");
    await ctx.sync();
    // Range value should be [10, 20, 30] when we load the property, because calculation is resumed after last sync
    console.log(rangeToGet.values);
})
```

### <a name="suspend-screen-updating"></a>Suspender a atualização da tela

O Excel exibe as alterações que seu suplemento faz aproximadamente conforme elas acontecem no código. Para conjuntos de dados grandes e interativos, talvez não seja necessário não esse andamento na tela em tempo real. `Application.suspendScreenUpdatingUntilNextSync()` pausa atualizações visuais no Excel até as chamadas do suplemento `context.sync()`, ou até o`Excel.run` terminar (chamadas implícitas `context.sync`). Lembre-se, o Excel não mostrará os sinais de atividade até a próxima sincronização. Seu suplemento deve fornecer orientação aos usuários para prepará-los para esse atraso ou fornecer uma barra de status para demonstrar atividade.

> [!NOTE]
> Não chame `suspendScreenUpdatingUntilNextSync` repetidamente (como em um loop). As chamadas repetidas farão com que a janela do Excel fique de piscar.

### <a name="enable-and-disable-events"></a>Habilitar e desabilitar eventos

O desempenho de um suplemento pode ser melhorado desabilitando eventos. Um exemplo de código mostrando como habilitar e desabilitar os eventos está no artigo [trabalhar com eventos](excel-add-ins-events.md#enable-and-disable-events).

## <a name="update-all-cells-in-a-range"></a>Atualizar todas as células em um intervalo

Quando você precisa atualizar todas as células em um intervalo com o mesmo valor ou propriedade, pode ser lento fazer isso por meio de uma matriz bidimensional que especifica repetidamente o mesmo valor, já que essa abordagem requer que o Excel faça uma iteração em todas as células do intervalo para definir cada uma delas separadamente. O Excel tem uma forma mais eficiente para atualizar todas as células em um intervalo com o mesmo valor ou propriedade.

Se desejar aplicar o mesmo valor, o mesmo formato de número ou a mesma fórmula para um intervalo de células, é mais eficiente especificar um valor único em vez de uma matriz de valores. Isso melhorará consideravelmente o desempenho. Para ver uma amostra de código que mostra essa abordagem em ação, confira [conceitos fundamentais: atualizar todas as células em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

Um cenário comum em que você pode aplicar essa abordagem é ao configurar formatos numéricos diferentes em colunas diferentes em uma planilha. Nesse caso, simplesmente percorra as colunas e defina o formato de número em cada coluna com um único valor. Lidar com cada coluna como um intervalo, como é mostrado  na amostra de código [atualizar todas as células em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).

> [!NOTE]
> Se você estiver usando o TypeScript, vai notar um erro de compilação dizendo que um único valor não pode ser  definido como uma matriz 2D.  Isso é inevitável, pois os valores *são* uma matriz 2D ao recuperar as propriedades e o TypeScript não permite diferentes tipos de setter vs getter.  No entanto, uma solução simples é definir valores com um sufixo`as any`, por exemplo, `range.values = "hello world" as any`.

## <a name="importing-data-into-tables"></a>Importar dados em tabelas

Ao tentar importar um grande volume de dados diretamente em um objeto[tabela](/javascript/api/excel/excel.table) diretamente (por exemplo, usando `TableRowCollection.add()`), você poderá observar um desempenho lento. Se você estiver tentando adicionar uma nova tabela, você deve preencher os dados primeiro definindo `range.values`e em seguida, ligue `worksheet.tables.add()` para criar uma tabela de intervalo. Se você está tentando gravar dados em uma tabela existente, grave os dados em um intervalo de objeto via`table.getDataBodyRange()`, e a tabela será expandida automaticamente. 

Aqui está um exemplo dessa abordagem:

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
> Você pode converter convenientemente um objeto de tabela em um objeto de intervalo usando o método[Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).

## <a name="untrack-unneeded-ranges"></a>Desviar intervalos desnecessários

A camada JavaScript cria objetos de proxy para o seu suplemento interagir com a pasta de trabalho do Excel e os intervalos subjacentes. Esses objetos são mantidos na memória até `context.sync()` ser acionado. Grandes operações em lote podem gerar muitos objetos de proxy que são necessários apenas uma vez pelo suplemento e podem ser liberados da memória antes da execução do lote.

O método [Range.untrack()](/javascript/api/excel/excel.range#untrack--) libera um Objeto Range do Excel da memória. Chamar esse método depois que o suplemento for feito com o intervalo deve render um benefício de desempenho perceptível ao usar um grande número de objetos Range.

> [!NOTE]
> `Range.untrack()` é um atalho para [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-). Qualquer objeto de proxy pode ser não-rastreado, removendo-o da lista de objetos rastreados no contexto. Normalmente, os objetos Range são os únicos objetos do Excel usados ​​em quantidade suficiente para justificar o não-rastreamento.

O exemplo de código a seguir preenche um intervalo selecionado com dados, uma célula por vez. Depois que o valor é adicionado à célula, o intervalo que representa a célula é não-rastreado. Execute esse código em um intervalo selecionado de 20.000 de 10.000 células, primeiro, com a linha `cell.untrack()` e, em seguida, sem ela. Você deve observar que o código é executado mais rapidamente com a linha `cell.untrack()` do que sem ela. Você também poderá observar um tempo de resposta mais rápido posteriormente, porque a etapa de limpeza leva menos tempo.

```js
Excel.run(async (context) => {
    var largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    for (var i = 0; i < largeRange.rowCount; i++) {
        for (var j = 0; j < largeRange.columnCount; j++) {
            var cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // call untrack() to release the range from memory
            cell.untrack();
        }
    }

    await context.sync();
});
```

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Conceitos avançados de programação com a API JavaScript do Excel](excel-add-ins-advanced-concepts.md)
- [Limites de recurso e otimização de desempenho para Suplementos do Office](../concepts/resource-limits-and-performance-optimization.md)
- [Objeto de funções de planilha (API JavaScript para Excel)](/javascript/api/excel/excel.functions)
