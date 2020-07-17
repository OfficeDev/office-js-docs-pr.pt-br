---
title: Otimização de desempenho do da API JavaScript do Excel
description: Otimizar o desempenho usando as API JavaScript do Excel
ms.date: 07/14/2020
localization_priority: Normal
ms.openlocfilehash: 193cbe8c8cd1a432c6567401ed645990cb93e5e9
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159091"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="caa98-103">Otimização de desempenho usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="caa98-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="caa98-104">Existem várias maneiras de executar tarefas comuns com a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="caa98-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="caa98-105">Você encontrará diferenças significativas de desempenho entre várias abordagens.</span><span class="sxs-lookup"><span data-stu-id="caa98-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="caa98-106">Este artigo fornece orientações e amostras de código para mostrar como realizar tarefas comuns com eficiência usando as API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="caa98-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="caa98-107">Minimizar o número de chamadas sync()</span><span class="sxs-lookup"><span data-stu-id="caa98-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="caa98-108">Na API do JavaScript do Excel, `sync()` é a única operação assíncrona e pode ser lenta em algumas circunstâncias, especialmente no Excel Online na Web.</span><span class="sxs-lookup"><span data-stu-id="caa98-108">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="caa98-109">Para otimizar o desempenho, minimize o número de chamadas para `sync()`, enfileirando o maior número possível de alterações antes de chamá-lo.</span><span class="sxs-lookup"><span data-stu-id="caa98-109">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="caa98-110">Ver [Principais conceitos - sync()](excel-add-ins-core-concepts.md#sync) para as amostras de código que seguem esta prática.</span><span class="sxs-lookup"><span data-stu-id="caa98-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="caa98-111">Minimizar o número de objetos proxy criados</span><span class="sxs-lookup"><span data-stu-id="caa98-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="caa98-112">Evite criar repetidamente o mesmo objeto proxy.</span><span class="sxs-lookup"><span data-stu-id="caa98-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="caa98-113">Em vez disso, se você precisar do mesmo objeto proxy para mais de uma operação, crie-o uma vez e o atribua a uma variável, em seguida, use essa variável no seu código.</span><span class="sxs-lookup"><span data-stu-id="caa98-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="caa98-114">Carregar propriedades necessárias</span><span class="sxs-lookup"><span data-stu-id="caa98-114">Load necessary properties only</span></span>

<span data-ttu-id="caa98-115">Na API JavaScript do Excel, você precisa explicitamente carregar as propriedades de um objeto de proxy.</span><span class="sxs-lookup"><span data-stu-id="caa98-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="caa98-116">Embora você seja capaz de carregar todas as propriedades de uma vez com uma `load()` chamada vazia, essa abordagem pode ter uma sobrecarga de desempenho significativa.</span><span class="sxs-lookup"><span data-stu-id="caa98-116">Although you're able to load all the properties at once with an empty `load()` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="caa98-117">Em vez disso, é recomendável apenas carregar as propriedades necessárias, especialmente para os objetos que têm um grande número de propriedades.</span><span class="sxs-lookup"><span data-stu-id="caa98-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="caa98-118">Por exemplo, se você pretende apenas ler a `address` propriedade de um objeto Range, especifique somente essa propriedade quando chamar o `load()` método:</span><span class="sxs-lookup"><span data-stu-id="caa98-118">For example, if you only intend to read the `address` property of a range object, specify only that property when you call the `load()` method:</span></span>

```js
range.load('address');
```

<span data-ttu-id="caa98-119">Você pode chamar `load()` método de qualquer uma das seguintes maneiras:</span><span class="sxs-lookup"><span data-stu-id="caa98-119">You can call `load()` method in any of the following ways:</span></span>

<span data-ttu-id="caa98-120">_Sintaxe:_</span><span class="sxs-lookup"><span data-stu-id="caa98-120">_Syntax:_</span></span>

```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```

<span data-ttu-id="caa98-121">_Onde:_</span><span class="sxs-lookup"><span data-stu-id="caa98-121">_Where:_</span></span>

* <span data-ttu-id="caa98-122">`properties` é a lista de propriedades para carregar, especificadas como cadeias de caracteres delimitadas por vírgula ou como uma matriz de nomes.</span><span class="sxs-lookup"><span data-stu-id="caa98-122">`properties` is the list of properties to load, specified as comma-delimited strings or as an array of names.</span></span> <span data-ttu-id="caa98-123">Para obter mais informações, consulte os `load()` métodos definidos para objetos na [referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md).</span><span class="sxs-lookup"><span data-stu-id="caa98-123">For more information, see the `load()` methods defined for objects in [Excel JavaScript API reference](../reference/overview/excel-add-ins-reference-overview.md).</span></span>
* <span data-ttu-id="caa98-p106">`loadOption` especifica um objeto que descreve as opções de seleção, expansão, topo e ignorar. Confira as [opções](/javascript/api/office/officeextension.loadoption) de carregamento do objeto para saber mais.</span><span class="sxs-lookup"><span data-stu-id="caa98-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="caa98-126">Observe que algumas das "Propriedades" em um objeto podem ter o mesmo nome de outro objeto.</span><span class="sxs-lookup"><span data-stu-id="caa98-126">Please be aware that some of the "properties" under an object may have the same name as another object.</span></span> <span data-ttu-id="caa98-127">Por exemplo, `format` é uma propriedade dentro do objeto de intervalo, mas `format` também é um objeto.</span><span class="sxs-lookup"><span data-stu-id="caa98-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="caa98-128">Portanto, se você fizer uma chamada, como `range.load("format")`, isso equivale a`range.format.load()`, que é uma chamada load vazia () que pode causar problemas de desempenho, conforme descrito anteriormente.</span><span class="sxs-lookup"><span data-stu-id="caa98-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="caa98-129">Para evitar isso, o código só deve carregar os "nós folha" em uma árvore de objetos.</span><span class="sxs-lookup"><span data-stu-id="caa98-129">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="caa98-130">Suspender temporariamente os processos do Excel</span><span class="sxs-lookup"><span data-stu-id="caa98-130">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="caa98-131">O Excel tem várias tarefas em segundo plano reagindo à entrada de usuários e seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="caa98-131">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="caa98-132">Alguns desses processos do Excel podem ser controlado para obter o benefício de desempenho.</span><span class="sxs-lookup"><span data-stu-id="caa98-132">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="caa98-133">Isso é útil principalmente quando o suplemento lida com grandes conjuntos de dados.</span><span class="sxs-lookup"><span data-stu-id="caa98-133">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="caa98-134">Suspender os cálculos temporariamente</span><span class="sxs-lookup"><span data-stu-id="caa98-134">Suspend calculation temporarily</span></span>

<span data-ttu-id="caa98-135">Se você estiver tentando executar uma operação em um grande número de células (por exemplo, definindo o valor do objeto de um grande intervalo) e não se importar em suspender o cálculo no Excel temporariamente enquanto a operação for concluída, é recomendável que você suspenda o cálculo até o próximo `context.sync()` ser chamado.</span><span class="sxs-lookup"><span data-stu-id="caa98-135">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="caa98-136">Ver a documentação de referência [objeto de aplicativo](/javascript/api/excel/excel.application) para saber mais sobre como usar a API`suspendApiCalculationUntilNextSync()`para suspender e reativar cálculos de maneira muito fácil.</span><span class="sxs-lookup"><span data-stu-id="caa98-136">See the [Application Object](/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="caa98-137">O código a seguir demonstra como suspender temporariamente um cálculo:</span><span class="sxs-lookup"><span data-stu-id="caa98-137">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

<span data-ttu-id="caa98-138">Observe que somente os cálculos de fórmula são suspensos.</span><span class="sxs-lookup"><span data-stu-id="caa98-138">Please note that only formula calculations are suspended.</span></span> <span data-ttu-id="caa98-139">Todas as referências alteradas ainda serão recriadas.</span><span class="sxs-lookup"><span data-stu-id="caa98-139">Any altered references are still rebuilt.</span></span> <span data-ttu-id="caa98-140">Por exemplo, renomear uma planilha ainda atualiza quaisquer referências em fórmulas para essa planilha.</span><span class="sxs-lookup"><span data-stu-id="caa98-140">For example, renaming a worksheet still updates any references in formulas to that worksheet.</span></span>

### <a name="suspend-screen-updating"></a><span data-ttu-id="caa98-141">Suspender a atualização da tela</span><span class="sxs-lookup"><span data-stu-id="caa98-141">Suspend screen updating</span></span>

<span data-ttu-id="caa98-142">O Excel exibe as alterações que seu suplemento faz aproximadamente conforme elas acontecem no código.</span><span class="sxs-lookup"><span data-stu-id="caa98-142">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="caa98-143">Para conjuntos de dados grandes e interativos, talvez não seja necessário não esse andamento na tela em tempo real.</span><span class="sxs-lookup"><span data-stu-id="caa98-143">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="caa98-144">`Application.suspendScreenUpdatingUntilNextSync()` pausa atualizações visuais no Excel até as chamadas do suplemento `context.sync()`, ou até o`Excel.run` terminar (chamadas implícitas `context.sync`).</span><span class="sxs-lookup"><span data-stu-id="caa98-144">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="caa98-145">Lembre-se, o Excel não mostrará os sinais de atividade até a próxima sincronização. Seu suplemento deve fornecer orientação aos usuários para prepará-los para esse atraso ou fornecer uma barra de status para demonstrar atividade.</span><span class="sxs-lookup"><span data-stu-id="caa98-145">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

> [!NOTE]
> <span data-ttu-id="caa98-146">Não chame `suspendScreenUpdatingUntilNextSync` repetidamente (como em um loop).</span><span class="sxs-lookup"><span data-stu-id="caa98-146">Don't call `suspendScreenUpdatingUntilNextSync` repeatedly (such as in a loop).</span></span> <span data-ttu-id="caa98-147">As chamadas repetidas farão com que a janela do Excel fique de piscar.</span><span class="sxs-lookup"><span data-stu-id="caa98-147">Repeated calls will cause the Excel window to flicker.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="caa98-148">Habilitar e desabilitar eventos</span><span class="sxs-lookup"><span data-stu-id="caa98-148">Enable and disable events</span></span>

<span data-ttu-id="caa98-149">O desempenho de um suplemento pode ser melhorado desabilitando eventos.</span><span class="sxs-lookup"><span data-stu-id="caa98-149">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="caa98-150">Um exemplo de código mostrando como habilitar e desabilitar os eventos está no artigo [trabalhar com eventos](excel-add-ins-events.md#enable-and-disable-events).</span><span class="sxs-lookup"><span data-stu-id="caa98-150">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="caa98-151">Importar dados em tabelas</span><span class="sxs-lookup"><span data-stu-id="caa98-151">Importing data into tables</span></span>

<span data-ttu-id="caa98-152">Ao tentar importar um grande volume de dados diretamente em um objeto[tabela](/javascript/api/excel/excel.table) diretamente (por exemplo, usando `TableRowCollection.add()`), você poderá observar um desempenho lento.</span><span class="sxs-lookup"><span data-stu-id="caa98-152">When trying to import a huge amount of data directly into a [Table](/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="caa98-153">Se você estiver tentando adicionar uma nova tabela, você deve preencher os dados primeiro definindo `range.values`e em seguida, ligue `worksheet.tables.add()` para criar uma tabela de intervalo.</span><span class="sxs-lookup"><span data-stu-id="caa98-153">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="caa98-154">Se você está tentando gravar dados em uma tabela existente, grave os dados em um intervalo de objeto via`table.getDataBodyRange()`, e a tabela será expandida automaticamente.</span><span class="sxs-lookup"><span data-stu-id="caa98-154">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="caa98-155">Aqui está um exemplo dessa abordagem:</span><span class="sxs-lookup"><span data-stu-id="caa98-155">Here is an example of this approach:</span></span>

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
> <span data-ttu-id="caa98-156">Você pode converter convenientemente um objeto de tabela em um objeto de intervalo usando o método[Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="caa98-156">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="untrack-unneeded-ranges"></a><span data-ttu-id="caa98-157">Desviar intervalos desnecessários</span><span class="sxs-lookup"><span data-stu-id="caa98-157">Untrack unneeded ranges</span></span>

<span data-ttu-id="caa98-158">A camada JavaScript cria objetos de proxy para o seu suplemento interagir com a pasta de trabalho do Excel e os intervalos subjacentes.</span><span class="sxs-lookup"><span data-stu-id="caa98-158">The JavaScript layer creates proxy objects for your add-in to interact with the Excel workbook and underlying ranges.</span></span> <span data-ttu-id="caa98-159">Esses objetos são mantidos na memória até `context.sync()` ser acionado.</span><span class="sxs-lookup"><span data-stu-id="caa98-159">These objects persist in memory until `context.sync()` is called.</span></span> <span data-ttu-id="caa98-160">Grandes operações em lote podem gerar muitos objetos de proxy que são necessários apenas uma vez pelo suplemento e podem ser liberados da memória antes da execução do lote.</span><span class="sxs-lookup"><span data-stu-id="caa98-160">Large batch operations may generate a lot of proxy objects that are only needed once by the add-in and can be released from memory before the batch executes.</span></span>

<span data-ttu-id="caa98-161">O método [Range.untrack()](/javascript/api/excel/excel.range#untrack--) libera um Objeto Range do Excel da memória.</span><span class="sxs-lookup"><span data-stu-id="caa98-161">The [Range.untrack()](/javascript/api/excel/excel.range#untrack--) method releases an Excel Range object from memory.</span></span> <span data-ttu-id="caa98-162">Chamar esse método depois que o suplemento for feito com o intervalo deve render um benefício de desempenho perceptível ao usar um grande número de objetos Range.</span><span class="sxs-lookup"><span data-stu-id="caa98-162">Calling this method after your add-in is done with the range should yield a noticeable performance benefit when using large numbers of Range objects.</span></span>

> [!NOTE]
> <span data-ttu-id="caa98-163">`Range.untrack()` é um atalho para [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span><span class="sxs-lookup"><span data-stu-id="caa98-163">`Range.untrack()` is a shortcut for [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span></span> <span data-ttu-id="caa98-164">Qualquer objeto de proxy pode ser não-rastreado, removendo-o da lista de objetos rastreados no contexto.</span><span class="sxs-lookup"><span data-stu-id="caa98-164">Any proxy object can be untracked by removing it from the tracked objects list in the context.</span></span> <span data-ttu-id="caa98-165">Normalmente, os objetos Range são os únicos objetos do Excel usados ​​em quantidade suficiente para justificar o não-rastreamento.</span><span class="sxs-lookup"><span data-stu-id="caa98-165">Typically, Range objects are the only Excel objects used in sufficient quantity to justify untracking.</span></span>

<span data-ttu-id="caa98-166">O exemplo de código a seguir preenche um intervalo selecionado com dados, uma célula por vez.</span><span class="sxs-lookup"><span data-stu-id="caa98-166">The following code sample fills a selected range with data, one cell at a time.</span></span> <span data-ttu-id="caa98-167">Depois que o valor é adicionado à célula, o intervalo que representa a célula é não-rastreado.</span><span class="sxs-lookup"><span data-stu-id="caa98-167">After the value is added to the cell, the range representing that cell is untracked.</span></span> <span data-ttu-id="caa98-168">Execute esse código em um intervalo selecionado de 20.000 de 10.000 células, primeiro, com a linha `cell.untrack()` e, em seguida, sem ela.</span><span class="sxs-lookup"><span data-stu-id="caa98-168">Run this code with a selected range of 10,000 to 20,000 cells, first with the `cell.untrack()` line, and then without it.</span></span> <span data-ttu-id="caa98-169">Você deve observar que o código é executado mais rapidamente com a linha `cell.untrack()` do que sem ela.</span><span class="sxs-lookup"><span data-stu-id="caa98-169">You should notice the code runs faster with the `cell.untrack()` line than without it.</span></span> <span data-ttu-id="caa98-170">Você também poderá observar um tempo de resposta mais rápido posteriormente, porque a etapa de limpeza leva menos tempo.</span><span class="sxs-lookup"><span data-stu-id="caa98-170">You may also notice a quicker response time afterwards, since the cleanup step takes less time.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="caa98-171">Confira também</span><span class="sxs-lookup"><span data-stu-id="caa98-171">See also</span></span>

- [<span data-ttu-id="caa98-172">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="caa98-172">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="caa98-173">Conceitos avançados de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="caa98-173">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="caa98-174">Limites de recurso e otimização de desempenho para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="caa98-174">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
- [<span data-ttu-id="caa98-175">Objeto de funções de planilha (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="caa98-175">Worksheet Functions Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.functions)
