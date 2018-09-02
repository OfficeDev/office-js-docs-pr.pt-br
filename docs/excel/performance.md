---
title: Otimização de desempenho da API JavaScript do Excel
description: Otimize o desempenho usando a API JavaScript do Excel
ms.date: 03/28/2018
ms.openlocfilehash: 50fac999093abb3fbfe1bd5be1cd6a77dc930399
ms.sourcegitcommit: 78b28ae88d53bfef3134c09cc4336a5a8722c70b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/01/2018
ms.locfileid: "23797312"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="57b39-103">Otimização de desempenho usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="57b39-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="57b39-104">Existem várias maneiras de executar tarefas comuns com a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="57b39-104">There are multiple ways that you can perform common tasks with the Excel JavaScript API.</span></span> <span data-ttu-id="57b39-105">Você encontrará diferenças de desempenho significativas entre as diferentes abordagens.</span><span class="sxs-lookup"><span data-stu-id="57b39-105">You'll find significant performance differences between various approaches.</span></span> <span data-ttu-id="57b39-106">Este artigo fornece diretrizes e exemplos de código para mostrar como executar tarefas comuns com eficiência usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="57b39-106">This article provides code samples that show how to perform common tasks with ranges using the Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="57b39-107">Minimizar o número de chamadas sync()</span><span class="sxs-lookup"><span data-stu-id="57b39-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="57b39-108">Na API JavaScript do Excel, ```sync()``` é a única operação assíncrona e pode ser lenta em determinadas circunstâncias, especialmente no Excel Online.</span><span class="sxs-lookup"><span data-stu-id="57b39-108">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online.</span></span> <span data-ttu-id="57b39-109">Para otimizar o desempenho, minimize o número de chamadas para ```sync()```, colocando em fila o maior número possível de alterações antes de chamá-la.</span><span class="sxs-lookup"><span data-stu-id="57b39-109">To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="57b39-110">Veja [Conceitos Básicos - sync()](excel-add-ins-core-concepts.md#sync) para obter exemplos de código que seguem essa prática.</span><span class="sxs-lookup"><span data-stu-id="57b39-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="57b39-111">Minimizar o número de objetos de proxy criados</span><span class="sxs-lookup"><span data-stu-id="57b39-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="57b39-112">Evite criar repetidamente o mesmo objeto de proxy.</span><span class="sxs-lookup"><span data-stu-id="57b39-112">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="57b39-113">Em vez disso, se precisar usar um mesmo objeto de proxy em mais de uma operação, crie-o uma única vez, atribua-o a uma variável e use essa variável no código.</span><span class="sxs-lookup"><span data-stu-id="57b39-113">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="57b39-114">Carregar apenas as propriedades necessárias</span><span class="sxs-lookup"><span data-stu-id="57b39-114">Load necessary properties only</span></span>

<span data-ttu-id="57b39-115">Na API JavaScript do Excel, é preciso carregar explicitamente as propriedades de um objeto de proxy.</span><span class="sxs-lookup"><span data-stu-id="57b39-115">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object.</span></span> <span data-ttu-id="57b39-116">Embora seja possível carregar todas as propriedades de uma só vez com uma chamada vazia de ```load()```, essa abordagem pode sobrecarregar significativamente o desempenho.</span><span class="sxs-lookup"><span data-stu-id="57b39-116">Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead.</span></span> <span data-ttu-id="57b39-117">Em vez disso, sugerimos carregar apenas as propriedades necessárias, especialmente para aqueles objetos que possuem um grande número de propriedades.</span><span class="sxs-lookup"><span data-stu-id="57b39-117">Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="57b39-118">Por exemplo, se sua intenção é apenas ler a propriedade **address** de um objeto de intervalo, especifique somente essa propriedade quando chamar o método **load()**:</span><span class="sxs-lookup"><span data-stu-id="57b39-118">For example, if you only intend to read back the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="57b39-119">É possível chamar o método **load()** de duas maneiras:</span><span class="sxs-lookup"><span data-stu-id="57b39-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="57b39-120">_Sintaxe:_</span><span class="sxs-lookup"><span data-stu-id="57b39-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="57b39-121">_Onde:_</span><span class="sxs-lookup"><span data-stu-id="57b39-121">_Where:_</span></span>
 
* <span data-ttu-id="57b39-122">`properties` é a lista de propriedades a serem carregadas especificadas como sequências de caracteres delimitadas por vírgula ou como uma matriz de nomes.</span><span class="sxs-lookup"><span data-stu-id="57b39-122">`properties` is the list of properties and/or relationship names to be loaded specified as comma-delimited strings, or an array of names.</span></span> <span data-ttu-id="57b39-123">Para saber mais, veja os métodos **load()** definidos para objetos na [referência da API JavaScript do Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="57b39-123">For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="57b39-p106">`loadOption` especifica um objeto que descreve as opções de seleção, expansão, topo e ignorar. Confira as [opções](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) de carregamento do objeto obter detalhes.</span><span class="sxs-lookup"><span data-stu-id="57b39-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="57b39-126">Observe que algumas das “propriedades” sob um objeto podem ter o mesmo nome que outro objeto.</span><span class="sxs-lookup"><span data-stu-id="57b39-126">Please be aware that some of the “properties” under an object may have the same name as another object.</span></span> <span data-ttu-id="57b39-127">Por exemplo, `format` é uma propriedade do objeto de intervalo, mas `format` em si também é um objeto.</span><span class="sxs-lookup"><span data-stu-id="57b39-127">For example, `format` is a property under range object, but `format` itself is an object as well.</span></span> <span data-ttu-id="57b39-128">Assim, se você fizer uma chamada como `range.load("format")`, isso será equivalente a `range.format.load()`, que é uma chamada vazia de load() que pode causar problemas de desempenho, conforme descrito anteriormente.</span><span class="sxs-lookup"><span data-stu-id="57b39-128">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously.</span></span> <span data-ttu-id="57b39-129">Para evitar isso, o código deve carregar apenas os "nós folha" em uma árvore de objetos.</span><span class="sxs-lookup"><span data-stu-id="57b39-129">To avoid this, your code should only load the “leaf nodes” in an object tree.</span></span> 

## <a name="suspend-calculation-temporarily"></a><span data-ttu-id="57b39-130">Suspender os cálculos temporariamente</span><span class="sxs-lookup"><span data-stu-id="57b39-130">Suspend calculation temporarily</span></span>

<span data-ttu-id="57b39-131">Se estiver tentando executar uma operação em um grande número de células (por exemplo, configurando o valor de um objeto de intervalo enorme) e não se importar em suspender temporariamente os cálculos no Excel enquanto a operação é concluída, recomendamos suspender os cálculos até a chamada da próxima ```context.sync()```.</span><span class="sxs-lookup"><span data-stu-id="57b39-131">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.</span></span>

<span data-ttu-id="57b39-132">Veja a documentação de referência do [Objeto de Aplicativo](https://docs.microsoft.com/javascript/api/excel/excel.application) para obter informações sobre como usar a ```suspendApiCalculationUntilNextSync()``` API para suspender e reativar os cálculos de forma muito conveniente.</span><span class="sxs-lookup"><span data-stu-id="57b39-132">See [Application Object](https://docs.microsoft.com/javascript/api/excel/excel.application) reference documentation for information about how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way.</span></span> <span data-ttu-id="57b39-133">O seguinte código demonstra como suspender os cálculos temporariamente:</span><span class="sxs-lookup"><span data-stu-id="57b39-133">The following code demonstrates how to suspend calculation temporarily:</span></span>

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

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="57b39-134">Atualizar todas as células em um intervalo</span><span class="sxs-lookup"><span data-stu-id="57b39-134">Update all cells in a range</span></span> 

<span data-ttu-id="57b39-135">Quando você precisa atualizar todas as células de um intervalo com um mesmo valor ou propriedade, poderá ser lento fazer isso por meio de uma matriz bidimensional que especifica repetidamente o mesmo valor, já que essa abordagem exige que o Excel execute uma iteração em todas as células do intervalo para definir cada uma separadamente.</span><span class="sxs-lookup"><span data-stu-id="57b39-135">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately.</span></span> <span data-ttu-id="57b39-136">O Excel tem uma maneira mais eficiente para atualizar todas as células de um intervalo com um mesmo valor ou propriedade.</span><span class="sxs-lookup"><span data-stu-id="57b39-136">Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="57b39-137">Se você precisar aplicar o mesmo valor, formato de número ou fórmula a um intervalo de células, será mais eficiente especificar um único valor em vez de uma matriz de valores.</span><span class="sxs-lookup"><span data-stu-id="57b39-137">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values.</span></span> <span data-ttu-id="57b39-138">Isso aumentará significativamente o desempenho.</span><span class="sxs-lookup"><span data-stu-id="57b39-138">Doing so will significantly improve performance.</span></span> <span data-ttu-id="57b39-139">Para obter um exemplo de código que mostre essa abordagem em ação, veja [Conceitos principais - Atualizar todas as células em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="57b39-139">For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="57b39-140">Um cenário comum em que você pode aplicar essa abordagem é ao definir formatos numéricos diferentes em colunas diferentes em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="57b39-140">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet.</span></span> <span data-ttu-id="57b39-141">Nesse caso, você pode simplesmente iterar pelas colunas e definir o formato numérico em cada coluna com um valor único.</span><span class="sxs-lookup"><span data-stu-id="57b39-141">In this case, you can simply iterate through the columns and set the number format on each column with a single value.</span></span> <span data-ttu-id="57b39-142">Trate cada coluna como um intervalo, conforme mostrado no exemplo de código em [Atualizar todas as células em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="57b39-142">Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="57b39-143">Se estiver usando o TypeScript, você observará um erro de compilação informando que um valor único não pode ser definido como uma matriz 2D.</span><span class="sxs-lookup"><span data-stu-id="57b39-143">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.</span></span>  <span data-ttu-id="57b39-144">Isso é inevitável, já que os valores *são* uma matriz 2D ao recuperar as propriedades e o TypeScript não permite tipos diferentes de setter versus getter.</span><span class="sxs-lookup"><span data-stu-id="57b39-144">This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.</span></span>  <span data-ttu-id="57b39-145">No entanto, uma solução simples é definir os valores com um sufixo `as any`, por exemplo, `range.values = "hello world" as any`.</span><span class="sxs-lookup"><span data-stu-id="57b39-145">However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="57b39-146">Importação de dados em tabelas</span><span class="sxs-lookup"><span data-stu-id="57b39-146">Importing data into tables</span></span>

<span data-ttu-id="57b39-147">Ao tentar importar uma enorme quantidade de dados diretamente para um objeto [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) (por exemplo, usando `TableRowCollection.add()`), o desempenho poderá ser mais lento.</span><span class="sxs-lookup"><span data-stu-id="57b39-147">When trying to import a huge amount of data directly into a [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance.</span></span> <span data-ttu-id="57b39-148">Se você estiver tentando adicionar uma nova tabela, deverá preencher os dados primeiro, definindo `range.values` e chamando `worksheet.tables.add()` para criar uma tabela no intervalo.</span><span class="sxs-lookup"><span data-stu-id="57b39-148">If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range.</span></span> <span data-ttu-id="57b39-149">Se você estiver tentando gravar dados em uma tabela existente, grave-os em um objeto de intervalo por meio de `table.getDataBodyRange()` e a tabela será expandida automaticamente.</span><span class="sxs-lookup"><span data-stu-id="57b39-149">If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="57b39-150">Veja um exemplo dessa abordagem:</span><span class="sxs-lookup"><span data-stu-id="57b39-150">Here is an example in JavaScript of this operation.</span></span>

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
> <span data-ttu-id="57b39-151">É possível converter convenientemente um objeto Table em um objeto Range usando o método [Table.convertToRange()](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="57b39-151">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="enable-and-disable-events"></a><span data-ttu-id="57b39-152">Ativar e desativar eventos</span><span class="sxs-lookup"><span data-stu-id="57b39-152">Enable and disable agents</span></span>

<span data-ttu-id="57b39-153">O desempenho de um suplemento pode ser melhorado por meio da desabilitação de eventos.</span><span class="sxs-lookup"><span data-stu-id="57b39-153">Performance of an add-in may be improved by disabling events.</span></span> <span data-ttu-id="57b39-154">Um exemplo de código mostrando como habilitar e desabilitar eventos está no artigo [Trabalho com eventos](excel-add-ins-events.md#enable-and-disable-events) .</span><span class="sxs-lookup"><span data-stu-id="57b39-154">A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="57b39-155">Confira também</span><span class="sxs-lookup"><span data-stu-id="57b39-155">See also</span></span>

- [<span data-ttu-id="57b39-156">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="57b39-156">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="57b39-157">Conceitos avançados da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="57b39-157">Excel JavaScript API advanced concepts</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="57b39-158">Especificação aberta da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="57b39-158">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="57b39-159">Objeto de funções de planilha (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="57b39-159">Worksheet Functions Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.functions)
