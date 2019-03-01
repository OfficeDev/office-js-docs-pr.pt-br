---
title: Otimização de desempenho do da API JavaScript do Excel
description: Otimizar o desempenho usando as API JavaScript do Excel
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: d15a4b3ad4ae44399572282889855b1cdc32bc39
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/28/2019
ms.locfileid: "30199575"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="5e001-103">Otimização de desempenho usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="5e001-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="5e001-104">Existem várias maneiras de executar tarefas comuns com a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="5e001-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="5e001-105">Você encontrará diferenças significativas de desempenho entre várias abordagens.</span><span class="sxs-lookup"><span data-stu-id="5e001-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="5e001-106">Este artigo fornece orientações e amostras de código para mostrar como realizar tarefas comuns com eficiência usando as API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="5e001-106">This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="5e001-107">Minimizar o número de chamadas sync()</span><span class="sxs-lookup"><span data-stu-id="5e001-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="5e001-108">Na API do JavaScript do Excel, ```sync()``` é a única operação assíncrona e pode ser lenta em algumas circunstâncias, especialmente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="5e001-108">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online.</span></span> <span data-ttu-id="5e001-109">Para otimizar o desempenho, minimize o número de chamadas para ```sync()```, enfileirando o maior número possível de alterações antes de chamá-lo.</span><span class="sxs-lookup"><span data-stu-id="5e001-109">To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="5e001-110">Ver [Principais conceitos - sync()](excel-add-ins-core-concepts.md#sync) para as amostras de código que seguem esta prática.</span><span class="sxs-lookup"><span data-stu-id="5e001-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="5e001-111">Minimizar o número de objetos proxy criados</span><span class="sxs-lookup"><span data-stu-id="5e001-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="5e001-112">Evite criar repetidamente o mesmo objeto proxy.</span><span class="sxs-lookup"><span data-stu-id="5e001-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="5e001-113">Em vez disso, se você precisar do mesmo objeto proxy para mais de uma operação, crie-o uma vez e o atribua a uma variável, em seguida, use essa variável no seu código.</span><span class="sxs-lookup"><span data-stu-id="5e001-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="5e001-114">Carregar propriedades necessárias</span><span class="sxs-lookup"><span data-stu-id="5e001-114">Load necessary properties only</span></span>

<span data-ttu-id="5e001-115">Na API JavaScript do Excel, você precisa explicitamente carregar as propriedades de um objeto de proxy.</span><span class="sxs-lookup"><span data-stu-id="5e001-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="5e001-116">Embora você seja capaz de carregar todas as propriedades de uma vez com uma ```load()``` chamada vazia, essa abordagem pode ter uma sobrecarga de desempenho significativa.</span><span class="sxs-lookup"><span data-stu-id="5e001-116">Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="5e001-117">Em vez disso, é recomendável apenas carregar as propriedades necessárias, especialmente para os objetos que têm um grande número de propriedades.</span><span class="sxs-lookup"><span data-stu-id="5e001-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="5e001-118">Por exemplo, se sua intenção é apenas ler a propriedade **address** de um objeto do intervalo, especifique somente essa propriedade quando chamar o método **load()**:</span><span class="sxs-lookup"><span data-stu-id="5e001-118">For example, if you only intend to read the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="5e001-119">Você pode chamar o método **load()** de duas maneiras:</span><span class="sxs-lookup"><span data-stu-id="5e001-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="5e001-120">_Sintaxe:_</span><span class="sxs-lookup"><span data-stu-id="5e001-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="5e001-121">_Onde:_</span><span class="sxs-lookup"><span data-stu-id="5e001-121">_Where:_</span></span>
 
* <span data-ttu-id="5e001-122">`properties` é a lista de propriedades para carregar, especificadas como cadeias de caracteres delimitadas por vírgula ou como uma matriz de nomes.</span><span class="sxs-lookup"><span data-stu-id="5e001-122">`properties` is the list of properties to load, specified as comma-delimited strings or as an array of names.</span></span> <span data-ttu-id="5e001-123">Para saber mais, veja os métodos **load()** definidos para objetos na [referência da API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="5e001-123">For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="5e001-p106">`loadOption` especifica um objeto que descreve as opções de seleção, expansão, topo e ignorar. Confira as [opções](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) de carregamento do objeto para saber mais.</span><span class="sxs-lookup"><span data-stu-id="5e001-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="5e001-126">Por favor, esteja ciente de que algumas das "propriedades" sob um objeto podem ter o mesmo nome que outro objeto.</span><span class="sxs-lookup"><span data-stu-id="5e001-126">Please be aware that some of the “properties” under an object may have the same name as another object.</span></span> <span data-ttu-id="5e001-127">Por exemplo, `format` é uma propriedade dentro do objeto de intervalo, mas `format` também é um objeto.</span><span class="sxs-lookup"><span data-stu-id="5e001-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="5e001-128">Portanto, se você fizer uma chamada, como `range.load("format")`, isso equivale a`range.format.load()`, que é uma chamada load vazia () que pode causar problemas de desempenho, conforme descrito anteriormente.</span><span class="sxs-lookup"><span data-stu-id="5e001-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="5e001-129">Para evitar isso, o código deve carregar apenas "nós folha" na árvore de objetos.</span><span class="sxs-lookup"><span data-stu-id="5e001-129">To avoid this, your code should only load the “leaf nodes” in an object tree.</span></span> 

## <a name="suspend-excel-processes-temporarily"></a><span data-ttu-id="5e001-130">Suspender temporariamente os processos do Excel</span><span class="sxs-lookup"><span data-stu-id="5e001-130">Suspend Excel processes temporarily</span></span>

<span data-ttu-id="5e001-131">O Excel tem várias tarefas em segundo plano reagindo à entrada de usuários e seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="5e001-131">Excel has a number of background tasks reacting to input from both users and your add-in.</span></span> <span data-ttu-id="5e001-132">Alguns desses processos do Excel podem ser controlado para obter o benefício de desempenho.</span><span class="sxs-lookup"><span data-stu-id="5e001-132">Some of these Excel processes can be controlled to yield a performance benefit.</span></span> <span data-ttu-id="5e001-133">Isso é útil principalmente quando o suplemento lida com grandes conjuntos de dados.</span><span class="sxs-lookup"><span data-stu-id="5e001-133">This is especially helpful when your add-in deals with large data sets.</span></span>

### <a name="suspend-calculation-temporarily"></a><span data-ttu-id="5e001-134">Suspender os cálculos temporariamente</span><span class="sxs-lookup"><span data-stu-id="5e001-134">Suspend calculation temporarily</span></span>

<span data-ttu-id="5e001-135">Se você estiver tentando executar uma operação em um grande número de células (por exemplo, definindo o valor do objeto de um grande intervalo) e não se importar em suspender o cálculo no Excel temporariamente enquanto a operação for concluída, é recomendável que você suspenda o cálculo até o próximo `context.sync()` ser chamado.</span><span class="sxs-lookup"><span data-stu-id="5e001-135">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next `context.sync()` is called.</span></span>

<span data-ttu-id="5e001-136">Ver a documentação de referência [objeto de aplicativo](https://docs.microsoft.com/javascript/api/excel/excel.application) para saber mais sobre como usar a API`suspendApiCalculationUntilNextSync()`para suspender e reativar cálculos de maneira muito fácil.</span><span class="sxs-lookup"><span data-stu-id="5e001-136">See the [Application Object](https://docs.microsoft.com/javascript/api/excel/excel.application) reference documentation for information about how to use the `suspendApiCalculationUntilNextSync()` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="5e001-137">O código a seguir demonstra como suspender temporariamente um cálculo:</span><span class="sxs-lookup"><span data-stu-id="5e001-137">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

### <a name="suspend-screen-updating"></a><span data-ttu-id="5e001-138">Suspender a atualização da tela</span><span class="sxs-lookup"><span data-stu-id="5e001-138">Suspend screen updating</span></span>

> [!NOTE]
> <span data-ttu-id="5e001-139">O método `suspendScreenUpdatingUntilNextSync` descrito neste artigo só está disponível atualmente na versão prévia pública.</span><span class="sxs-lookup"><span data-stu-id="5e001-139">The `suspendScreenUpdatingUntilNextSync` method described in this article is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

<span data-ttu-id="5e001-140">O Excel exibe as alterações que seu suplemento faz aproximadamente conforme elas acontecem no código.</span><span class="sxs-lookup"><span data-stu-id="5e001-140">Excel displays changes your add-in makes approximately as they happen in the code.</span></span> <span data-ttu-id="5e001-141">Para conjuntos de dados grandes e interativos, talvez não seja necessário não esse andamento na tela em tempo real.</span><span class="sxs-lookup"><span data-stu-id="5e001-141">For large, iterative data sets, you may not need to see this progress on the screen in real-time.</span></span> <span data-ttu-id="5e001-142">`Application.suspendScreenUpdatingUntilNextSync()` pausa atualizações visuais no Excel até as chamadas do suplemento `context.sync()`, ou até o`Excel.run` terminar (chamadas implícitas `context.sync`).</span><span class="sxs-lookup"><span data-stu-id="5e001-142">`Application.suspendScreenUpdatingUntilNextSync()` pauses visual updates to Excel until the add-in calls `context.sync()`, or until `Excel.run` ends (implicitly calling `context.sync`).</span></span> <span data-ttu-id="5e001-143">Lembre-se, o Excel não mostrará os sinais de atividade até a próxima sincronização. Seu suplemento deve fornecer orientação aos usuários para prepará-los para esse atraso ou fornecer uma barra de status para demonstrar atividade.</span><span class="sxs-lookup"><span data-stu-id="5e001-143">Be aware, Excel will not show any signs of activity until the next sync. Your add-in should either give users guidance to prepare them for this delay or provide a status bar to demonstrate activity.</span></span>

### <a name="enable-and-disable-events"></a><span data-ttu-id="5e001-144">Habilitar e desabilitar eventos</span><span class="sxs-lookup"><span data-stu-id="5e001-144">Enable and disable events</span></span>

<span data-ttu-id="5e001-145">O desempenho de um suplemento pode ser melhorado desabilitando eventos.</span><span class="sxs-lookup"><span data-stu-id="5e001-145">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="5e001-146">Um exemplo de código mostrando como habilitar e desabilitar os eventos está no artigo [trabalhar com eventos](excel-add-ins-events.md#enable-and-disable-events).</span><span class="sxs-lookup"><span data-stu-id="5e001-146">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="5e001-147">Atualizar todas as células em um intervalo</span><span class="sxs-lookup"><span data-stu-id="5e001-147">Update all cells in a range</span></span>

<span data-ttu-id="5e001-148">Quando você precisa atualizar todas as células em um intervalo com o mesmo valor ou propriedade, pode ser lento fazer isso por meio de uma matriz bidimensional que especifica repetidamente o mesmo valor, já que essa abordagem requer que o Excel faça uma iteração em todas as células do intervalo para definir cada uma delas separadamente.</span><span class="sxs-lookup"><span data-stu-id="5e001-148">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately.</span></span> <span data-ttu-id="5e001-149">O Excel tem uma forma mais eficiente para atualizar todas as células em um intervalo com o mesmo valor ou propriedade.</span><span class="sxs-lookup"><span data-stu-id="5e001-149">Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="5e001-150">Se desejar aplicar o mesmo valor, o mesmo formato de número ou a mesma fórmula para um intervalo de células, é mais eficiente especificar um valor único em vez de uma matriz de valores.</span><span class="sxs-lookup"><span data-stu-id="5e001-150">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values.</span></span> <span data-ttu-id="5e001-151">Isso melhorará consideravelmente o desempenho.</span><span class="sxs-lookup"><span data-stu-id="5e001-151">Doing so will significantly improve performance.</span></span> <span data-ttu-id="5e001-152">Para ver uma amostra de código que mostra essa abordagem em ação, confira [conceitos fundamentais: atualizar todas as células em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="5e001-152">For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="5e001-153">Um cenário comum em que você pode aplicar essa abordagem é ao configurar formatos numéricos diferentes em colunas diferentes em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="5e001-153">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet.</span></span> <span data-ttu-id="5e001-154">Nesse caso, simplesmente percorra as colunas e defina o formato de número em cada coluna com um único valor.</span><span class="sxs-lookup"><span data-stu-id="5e001-154">In this case, you can simply iterate through the columns and set the number format on each column with a single value.</span></span> <span data-ttu-id="5e001-155">Lidar com cada coluna como um intervalo, como é mostrado  na amostra de código [atualizar todas as células em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="5e001-155">Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="5e001-156">Se você estiver usando o TypeScript, vai notar um erro de compilação dizendo que um único valor não pode ser  definido como uma matriz 2D.</span><span class="sxs-lookup"><span data-stu-id="5e001-156">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.</span></span>  <span data-ttu-id="5e001-157">Isso é inevitável, pois os valores *são* uma matriz 2D ao recuperar as propriedades e o TypeScript não permite diferentes tipos de setter vs getter.</span><span class="sxs-lookup"><span data-stu-id="5e001-157">This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.</span></span>  <span data-ttu-id="5e001-158">No entanto, uma solução simples é definir valores com um sufixo`as any`, por exemplo, `range.values = "hello world" as any`.</span><span class="sxs-lookup"><span data-stu-id="5e001-158">However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="5e001-159">Importar dados em tabelas</span><span class="sxs-lookup"><span data-stu-id="5e001-159">Importing data into tables</span></span>

<span data-ttu-id="5e001-160">Ao tentar importar um grande volume de dados diretamente em um objeto[tabela](https://docs.microsoft.com/javascript/api/excel/excel.table) diretamente (por exemplo, usando `TableRowCollection.add()`), você poderá observar um desempenho lento.</span><span class="sxs-lookup"><span data-stu-id="5e001-160">When trying to import a huge amount of data directly into a [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="5e001-161">Se você estiver tentando adicionar uma nova tabela, você deve preencher os dados primeiro definindo `range.values`e em seguida, ligue `worksheet.tables.add()` para criar uma tabela de intervalo.</span><span class="sxs-lookup"><span data-stu-id="5e001-161">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="5e001-162">Se você está tentando gravar dados em uma tabela existente, grave os dados em um intervalo de objeto via`table.getDataBodyRange()`, e a tabela será expandida automaticamente.</span><span class="sxs-lookup"><span data-stu-id="5e001-162">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="5e001-163">Aqui está um exemplo dessa abordagem:</span><span class="sxs-lookup"><span data-stu-id="5e001-163">Here is an example of this approach:</span></span>

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
> <span data-ttu-id="5e001-164">Você pode converter convenientemente um objeto de tabela em um objeto de intervalo usando o método[Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="5e001-164">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="untrack-unneeded-ranges"></a><span data-ttu-id="5e001-165">Desviar intervalos desnecessários</span><span class="sxs-lookup"><span data-stu-id="5e001-165">Untrack unneeded ranges</span></span>

<span data-ttu-id="5e001-166">A camada JavaScript cria objetos de proxy para o seu suplemento interagir com a pasta de trabalho do Excel e os intervalos subjacentes.</span><span class="sxs-lookup"><span data-stu-id="5e001-166">The JavaScript layer creates proxy objects for your add-in to interact with the Excel workbook and underlying ranges.</span></span> <span data-ttu-id="5e001-167">Esses objetos são mantidos na memória até `context.sync()` ser acionado.</span><span class="sxs-lookup"><span data-stu-id="5e001-167">These objects persist in memory until `context.sync()` is called.</span></span> <span data-ttu-id="5e001-168">Grandes operações em lote podem gerar muitos objetos de proxy que são necessários apenas uma vez pelo suplemento e podem ser liberados da memória antes da execução do lote.</span><span class="sxs-lookup"><span data-stu-id="5e001-168">Large batch operations may generate a lot of proxy objects that are only needed once by the add-in and can be released from memory before the batch executes.</span></span>

<span data-ttu-id="5e001-169">O método [Range.untrack()](/javascript/api/excel/excel.range#untrack--) libera um Objeto Range do Excel da memória.</span><span class="sxs-lookup"><span data-stu-id="5e001-169">The [Range.untrack()](/javascript/api/excel/excel.range#untrack--) method releases an Excel Range object from memory.</span></span> <span data-ttu-id="5e001-170">Chamar esse método depois que o suplemento for feito com o intervalo deve render um benefício de desempenho perceptível ao usar um grande número de objetos Range.</span><span class="sxs-lookup"><span data-stu-id="5e001-170">Calling this method after your add-in is done with the range should yield a noticeable performance benefit when using large numbers of Range objects.</span></span>

> [!NOTE]
> <span data-ttu-id="5e001-171">`Range.untrack()` é um atalho para [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span><span class="sxs-lookup"><span data-stu-id="5e001-171">`Range.untrack()` is a shortcut for [ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-).</span></span> <span data-ttu-id="5e001-172">Qualquer objeto de proxy pode ser não-rastreado, removendo-o da lista de objetos rastreados no contexto.</span><span class="sxs-lookup"><span data-stu-id="5e001-172">Any proxy object can be untracked by removing it from the tracked objects list in the context.</span></span> <span data-ttu-id="5e001-173">Normalmente, os objetos Range são os únicos objetos do Excel usados ​​em quantidade suficiente para justificar o não-rastreamento.</span><span class="sxs-lookup"><span data-stu-id="5e001-173">Typically, Range objects are the only Excel objects used in sufficient quantity to justify untracking.</span></span>

<span data-ttu-id="5e001-174">O exemplo de código a seguir preenche um intervalo selecionado com dados, uma célula por vez.</span><span class="sxs-lookup"><span data-stu-id="5e001-174">The following code sample fills a selected range with data, one cell at a time.</span></span> <span data-ttu-id="5e001-175">Depois que o valor é adicionado à célula, o intervalo que representa a célula é não-rastreado.</span><span class="sxs-lookup"><span data-stu-id="5e001-175">After the value is added to the cell, the range representing that cell is untracked.</span></span> <span data-ttu-id="5e001-176">Execute esse código em um intervalo selecionado de 20.000 de 10.000 células, primeiro, com a linha `cell.untrack()` e, em seguida, sem ela.</span><span class="sxs-lookup"><span data-stu-id="5e001-176">Run this code with a selected range of 10,000 to 20,000 cells, first with the `cell.untrack()` line, and then without it.</span></span> <span data-ttu-id="5e001-177">Você deve observar que o código é executado mais rapidamente com a linha `cell.untrack()` do que sem ela.</span><span class="sxs-lookup"><span data-stu-id="5e001-177">You should notice the code runs faster with the `cell.untrack()` line than without it.</span></span> <span data-ttu-id="5e001-178">Você também poderá observar um tempo de resposta mais rápido posteriormente, porque a etapa de limpeza leva menos tempo.</span><span class="sxs-lookup"><span data-stu-id="5e001-178">You may also notice a quicker response time afterwards, since the cleanup step takes less time.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="5e001-179">Confira também</span><span class="sxs-lookup"><span data-stu-id="5e001-179">See also</span></span>

- [<span data-ttu-id="5e001-180">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="5e001-180">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="5e001-181">Conceitos avançados de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="5e001-181">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="5e001-182">Limites de recurso e otimização de desempenho para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5e001-182">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
- [<span data-ttu-id="5e001-183">Especificação abrir API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="5e001-183">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="5e001-184">Objeto de funções de planilha (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="5e001-184">Worksheet Functions Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.functions)
