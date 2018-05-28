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
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="c3637-103">Otimiza??o de desempenho usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="c3637-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="c3637-104">Existem v?rias maneiras de executar tarefas comuns com a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="c3637-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="c3637-105">Voc? encontrar? diferen?as de desempenho significativas entre as diferentes abordagens.</span><span class="sxs-lookup"><span data-stu-id="c3637-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="c3637-106">Este artigo fornece diretrizes e exemplos de c?digo para mostrar como executar tarefas comuns com efici?ncia usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="c3637-106">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="c3637-107">Minimizar o n?mero de chamadas sync()</span><span class="sxs-lookup"><span data-stu-id="c3637-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="c3637-108">Na API JavaScript do Excel, ```sync()``` ? a ?nica opera??o ass?ncrona e pode ser lenta em determinadas circunst?ncias, especialmente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="c3637-108">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online.</span></span> <span data-ttu-id="c3637-109">Para otimizar o desempenho, minimize o n?mero de chamadas para ```sync()```, colocando em fila o maior n?mero poss?vel de altera??es antes de cham?-la.</span><span class="sxs-lookup"><span data-stu-id="c3637-109">To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="c3637-110">Veja [Conceitos B?sicos - sync()](excel-add-ins-core-concepts.md#sync) para obter exemplos de c?digo que seguem essa pr?tica.</span><span class="sxs-lookup"><span data-stu-id="c3637-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="c3637-111">Minimizar o n?mero de objetos de proxy criados</span><span class="sxs-lookup"><span data-stu-id="c3637-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="c3637-112">Evite criar repetidamente o mesmo objeto de proxy.</span><span class="sxs-lookup"><span data-stu-id="c3637-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="c3637-113">Em vez disso, se precisar usar um mesmo objeto de proxy em mais de uma opera??o, crie-o uma ?nica vez, atribua-o a uma vari?vel e use essa vari?vel no c?digo.</span><span class="sxs-lookup"><span data-stu-id="c3637-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="c3637-114">Carregar apenas as propriedades necess?rias</span><span class="sxs-lookup"><span data-stu-id="c3637-114">Load necessary properties only</span></span>

<span data-ttu-id="c3637-115">Na API JavaScript do Excel, ? preciso carregar explicitamente as propriedades de um objeto de proxy.</span><span class="sxs-lookup"><span data-stu-id="c3637-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="c3637-116">Embora seja poss?vel carregar todas as propriedades de uma s? vez com uma chamada vazia de ```load()```, essa abordagem pode sobrecarregar significativamente o desempenho.</span><span class="sxs-lookup"><span data-stu-id="c3637-116">Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="c3637-117">Em vez disso, sugerimos carregar apenas as propriedades necess?rias, especialmente para aqueles objetos que possuem um grande n?mero de propriedades.</span><span class="sxs-lookup"><span data-stu-id="c3637-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="c3637-118">Por exemplo, se sua inten??o ? apenas ler a propriedade **address** de um objeto de intervalo, especifique somente essa propriedade ao chamar o m?todo **load()**:</span><span class="sxs-lookup"><span data-stu-id="c3637-118">For example, if you only intend to read back the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="c3637-119">? poss?vel chamar o m?todo **load()** de duas maneiras:</span><span class="sxs-lookup"><span data-stu-id="c3637-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="c3637-120">_Sintaxe:_</span><span class="sxs-lookup"><span data-stu-id="c3637-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="c3637-121">_Onde:_</span><span class="sxs-lookup"><span data-stu-id="c3637-121">_Where:_</span></span>
 
* <span data-ttu-id="c3637-122">`properties` ? a lista de propriedades a serem carregadas, especificadas como cadeias de caracteres delimitadas por v?rgula ou como uma matriz de nomes.</span><span class="sxs-lookup"><span data-stu-id="c3637-122">`properties` is the list of properties and/or relationship names to be loaded specified as comma-delimited strings, or an array of names.</span></span> <span data-ttu-id="c3637-123">Para saber mais, veja os m?todos **load()** definidos para objetos na [refer?ncia da API JavaScript do Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="c3637-123">For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="c3637-p106">`loadOption` especifica um objeto que descreve as op??es de sele??o, expans?o, topo e ignorar. Confira as [op??es](https://dev.office.com/reference/add-ins/excel/loadoption) de carregamento de objetos para saber mais.</span><span class="sxs-lookup"><span data-stu-id="c3637-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://dev.office.com/reference/add-ins/excel/loadoption) for details.</span></span>

<span data-ttu-id="c3637-126">Observe que algumas das ?propriedades? sob um objeto podem ter o mesmo nome que outro objeto.</span><span class="sxs-lookup"><span data-stu-id="c3637-126">Please be aware that some of the ?properties? under an object may have the same name as another object.</span></span> <span data-ttu-id="c3637-127">Por exemplo, `format` ? uma propriedade do objeto de intervalo, mas `format` em si tamb?m ? um objeto.</span><span class="sxs-lookup"><span data-stu-id="c3637-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="c3637-128">Assim, se voc? fizer uma chamada como `range.load("format")`, isso ser? equivalente a `range.format.load()`, que ? uma chamada vazia de load() que pode causar problemas de desempenho, conforme descrito anteriormente.</span><span class="sxs-lookup"><span data-stu-id="c3637-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="c3637-129">Para evitar isso, o c?digo deve carregar apenas os "n?s folha" em uma ?rvore de objetos.</span><span class="sxs-lookup"><span data-stu-id="c3637-129">To avoid this, your code should only load the ?leaf nodes? in an object tree.</span></span> 

## <a name="suspend-calculation-temporarily"></a><span data-ttu-id="c3637-130">Suspender os c?lculos temporariamente</span><span class="sxs-lookup"><span data-stu-id="c3637-130">Suspend calculation temporarily</span></span>

<span data-ttu-id="c3637-131">Se estiver tentando executar uma opera??o em um grande n?mero de c?lulas (por exemplo, configurando o valor de um objeto de intervalo enorme) e n?o se importar em suspender temporariamente os c?lculos no Excel enquanto a opera??o ? conclu?da, recomendamos suspender os c?lculos at? a chamada da pr?xima ```context.sync()```.</span><span class="sxs-lookup"><span data-stu-id="c3637-131">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.</span></span>

<span data-ttu-id="c3637-132">Veja a documenta??o de refer?ncia do [Objeto de Aplicativo](https://dev.office.com/reference/add-ins/excel/application) para obter informa??es sobre como usar a ```suspendApiCalculationUntilNextSync()``` API para suspender e reativar os c?lculos de forma muito conveniente.</span><span class="sxs-lookup"><span data-stu-id="c3637-132">See [Application Object](https://dev.office.com/reference/add-ins/excel/application) reference documentation for information about how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="c3637-133">O seguinte c?digo demonstra como suspender os c?lculos temporariamente:</span><span class="sxs-lookup"><span data-stu-id="c3637-133">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="c3637-134">Atualizar todas as c?lulas em um intervalo</span><span class="sxs-lookup"><span data-stu-id="c3637-134">Update all cells in a range</span></span> 

<span data-ttu-id="c3637-135">Quando voc? precisa atualizar todas as c?lulas de um intervalo com um mesmo valor ou propriedade, poder? ser lento fazer isso por meio de uma matriz bidimensional que especifica repetidamente o mesmo valor, j? que essa abordagem exige que o Excel execute uma itera??o em todas as c?lulas do intervalo para definir cada uma separadamente.</span><span class="sxs-lookup"><span data-stu-id="c3637-135">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately.</span></span> <span data-ttu-id="c3637-136">O Excel tem uma maneira mais eficiente para atualizar todas as c?lulas de um intervalo com um mesmo valor ou propriedade.</span><span class="sxs-lookup"><span data-stu-id="c3637-136">Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="c3637-137">Se voc? precisar aplicar o mesmo valor, formato de n?mero ou f?rmula a um intervalo de c?lulas, ser? mais eficiente especificar um ?nico valor em vez de uma matriz de valores.</span><span class="sxs-lookup"><span data-stu-id="c3637-137">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values.</span></span> <span data-ttu-id="c3637-138">Isso aumentar? significativamente o desempenho.</span><span class="sxs-lookup"><span data-stu-id="c3637-138">Doing so will significantly improve performance.</span></span> <span data-ttu-id="c3637-139">Para obter um exemplo de c?digo que mostre essa abordagem em a??o, veja [Conceitos principais - Atualizar todas as c?lulas em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="c3637-139">For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="c3637-140">Um cen?rio comum em que voc? pode aplicar essa abordagem ? ao definir formatos num?ricos diferentes em colunas diferentes em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="c3637-140">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet.</span></span> <span data-ttu-id="c3637-141">Nesse caso, voc? pode simplesmente iterar pelas colunas e definir o formato num?rico em cada coluna com um valor ?nico.</span><span class="sxs-lookup"><span data-stu-id="c3637-141">In this case, you can simply iterate through the columns and set the number format on each column with a single value.</span></span> <span data-ttu-id="c3637-142">Trate cada coluna como um intervalo, conforme mostrado no exemplo de c?digo em [Atualizar todas as c?lulas em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="c3637-142">Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="c3637-143">Se estiver usando o TypeScript, voc? observar? um erro de compila??o informando que um valor ?nico n?o pode ser definido como uma matriz 2D.</span><span class="sxs-lookup"><span data-stu-id="c3637-143">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.</span></span>  <span data-ttu-id="c3637-144">Isso ? inevit?vel, j? que os valores *s?o* uma matriz 2D ao recuperar as propriedades e o TypeScript n?o permite tipos diferentes de setter versus getter.</span><span class="sxs-lookup"><span data-stu-id="c3637-144">This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.</span></span>  <span data-ttu-id="c3637-145">No entanto, uma solu??o simples ? definir os valores com um sufixo `as any`, por exemplo, `range.values = "hello world" as any`.</span><span class="sxs-lookup"><span data-stu-id="c3637-145">However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="c3637-146">Importa??o de dados em tabelas</span><span class="sxs-lookup"><span data-stu-id="c3637-146">Importing data into tables</span></span>

<span data-ttu-id="c3637-147">Ao tentar importar uma enorme quantidade de dados diretamente para um objeto [Table](https://dev.office.com/reference/add-ins/excel/table) (por exemplo, usando `TableRowCollection.add()`), o desempenho poder? ser mais lento.</span><span class="sxs-lookup"><span data-stu-id="c3637-147">When trying to import a huge amount of data directly into a [Table](https://dev.office.com/reference/add-ins/excel/table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="c3637-148">Se voc? estiver tentando adicionar uma nova tabela, dever? preencher os dados primeiro, definindo `range.values` e chamando `worksheet.tables.add()` para criar uma tabela no intervalo.</span><span class="sxs-lookup"><span data-stu-id="c3637-148">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="c3637-149">Se voc? estiver tentando gravar dados em uma tabela existente, grave-os em um objeto de intervalo por meio de `table.getDataBodyRange()` e a tabela ser? expandida automaticamente.</span><span class="sxs-lookup"><span data-stu-id="c3637-149">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="c3637-150">Veja um exemplo dessa abordagem:</span><span class="sxs-lookup"><span data-stu-id="c3637-150">Here is an example in JavaScript of this operation.</span></span>

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
> <span data-ttu-id="c3637-151">? poss?vel converter convenientemente um objeto Table em um objeto Range usando o m?todo [Table.convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange).</span><span class="sxs-lookup"><span data-stu-id="c3637-151">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](https://dev.office.com/reference/add-ins/excel/table#converttorange) method.</span></span>

## <a name="see-also"></a><span data-ttu-id="c3637-152">Confira tamb?m</span><span class="sxs-lookup"><span data-stu-id="c3637-152">See also</span></span>

- [<span data-ttu-id="c3637-153">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="c3637-153">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c3637-154">Conceitos avan?ados da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="c3637-154">Excel JavaScript API advanced concepts</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="c3637-155">Especifica??o para abrir a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="c3637-155">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="c3637-156">Objeto de fun??es de planilha (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="c3637-156">Worksheet Functions Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/functions)
