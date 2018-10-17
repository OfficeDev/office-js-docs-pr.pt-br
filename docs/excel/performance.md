---
title: Otimização de desempenho da API JavaScript do Excel
description: Otimize o desempenho usando a API JavaScript do Excel
ms.date: 03/28/2018
ms.openlocfilehash: ee1687fcb1a5db74e65f5e73994653df235b4823
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505374"
---
# <a name="performance-optimization-using-the-excel-javascript-api"></a><span data-ttu-id="64e17-103">Otimização de desempenho usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="64e17-103">Performance optimization using the Excel JavaScript API</span></span>

<span data-ttu-id="64e17-p101">Há várias maneiras de realizar tarefas comuns com a API JavaScript do Excel. Você encontrará diferenças significativas de desempenho entre várias abordagens. Este artigo oferece orientação e códigos de exemplo para mostrar como executar tarefas comuns com eficiência usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="64e17-p101">There are multiple ways that you can perform common tasks with the Excel JavaScript API. You'll find significant performance differences between various approaches. This article provides guidance and code samples to show you how to perform common tasks efficiently using Excel JavaScript API.</span></span>

## <a name="minimize-the-number-of-sync-calls"></a><span data-ttu-id="64e17-107">Minimize o número de chamadas sync()</span><span class="sxs-lookup"><span data-stu-id="64e17-107">Minimize the number of sync() calls</span></span>

<span data-ttu-id="64e17-p102">Na API JavaScript do Excel, ```sync()``` é a única operação assíncrona, e ela pode ser lenta em algumas circunstâncias, especialmente no Excel Online. Para otimizar o desempenho, minimize o número de chamadas para ```sync()``` enfileirando o máximo possível de alterações antes de chamá-la.</span><span class="sxs-lookup"><span data-stu-id="64e17-p102">In the Excel JavaScript API, ```sync()``` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel Online. To optimize performance, minimize the number of calls to ```sync()``` by queueing up as many changes as possible before calling it.</span></span>

<span data-ttu-id="64e17-110">Confira [Conceitos Básicos - sync()](excel-add-ins-core-concepts.md#sync) para obter exemplos de código que seguem essa prática.</span><span class="sxs-lookup"><span data-stu-id="64e17-110">See [Core Concepts - sync()](excel-add-ins-core-concepts.md#sync) for code samples that follow this practice.</span></span>

## <a name="minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="64e17-111">Minimize o número de objetos de proxy criados</span><span class="sxs-lookup"><span data-stu-id="64e17-111">Minimize the number of proxy objects created</span></span>

<span data-ttu-id="64e17-p103">Evite criar repetidamente o mesmo objeto de proxy. Em vez disso, se precisar usar um mesmo objeto de proxy em mais de uma operação, crie-o uma única vez, atribua-o a uma variável e use essa variável no código.</span><span class="sxs-lookup"><span data-stu-id="64e17-p103">Avoid repeatedly creating the same proxy object. Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

## <a name="load-necessary-properties-only"></a><span data-ttu-id="64e17-114">Carregue apenas as propriedades necessárias</span><span class="sxs-lookup"><span data-stu-id="64e17-114">Load necessary properties only</span></span>

<span data-ttu-id="64e17-p104">Na API JavaScript do Excel, você precisa carregar explicitamente as propriedades de um objeto de proxy. Embora você possa carregar todas as propriedades de uma só vez com uma chamada de  ```load()``` vazio, essa abordagem pode afetar o desempenho de maneira significativa. Em vez disso, sugerimos que você carregue apenas as propriedades necessárias, especialmente para os objetos que têm um grande número de propriedades.</span><span class="sxs-lookup"><span data-stu-id="64e17-p104">In the Excel JavaScript API, you need to explicitly load the properties of a proxy object. Although you're able to load all the properties at once with an empty ```load()``` call, that approach can have significant performance overhead. Instead, we suggest that you only load the necessary properties, especially for those objects which have a large number of properties.</span></span>

<span data-ttu-id="64e17-118">Por exemplo, se sua intenção é apenas ler a propriedade **address** de um objeto range, especifique apenas essa propriedade quando chamar o método **load()**:</span><span class="sxs-lookup"><span data-stu-id="64e17-118">For example, if you only intend to read back the **address** property of a range object, specify only that property when you call the **load()** method:</span></span>
 
```js
range.load('address');
```
 
<span data-ttu-id="64e17-119">É possível chamar o método **load()** de duas maneiras:</span><span class="sxs-lookup"><span data-stu-id="64e17-119">You can call **load()** method in any of the following ways:</span></span>
 
<span data-ttu-id="64e17-120">_Sintaxe:_</span><span class="sxs-lookup"><span data-stu-id="64e17-120">_Syntax:_</span></span>
 
```js
object.load(string: properties);
// or
object.load(array: properties);
// or
object.load({ loadOption });
```
 
<span data-ttu-id="64e17-121">_Onde:_</span><span class="sxs-lookup"><span data-stu-id="64e17-121">_Where:_</span></span>
 
* <span data-ttu-id="64e17-p105">`properties` é a lista de propriedades que devem ser carregadas, especificada como sequências de caracteres delimitadas por vírgula ou como uma matriz de nomes. Para obter mais informações, consulte os métodos **load ()** definidos para os objetos na [referência de API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span><span class="sxs-lookup"><span data-stu-id="64e17-p105">`properties` is the list of properties to load, specified as comma-delimited strings or as an array of names. For more information, see the **load()** methods defined for objects in [Excel JavaScript API reference](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview).</span></span>
* <span data-ttu-id="64e17-p106">`loadOption` especifica um objeto que descreve as opções de seleção, expansão, topo e ignorar. Confira as [opções](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) de carregamento do objeto para saber mais.</span><span class="sxs-lookup"><span data-stu-id="64e17-p106">`loadOption` specifies an object that describes the selection, expansion, top, and skip options. See object load [options](https://docs.microsoft.com/javascript/api/office/officeextension.loadoption) for details.</span></span>

<span data-ttu-id="64e17-p107">Esteja ciente de que algumas das "propriedades" em um objeto podem ter o mesmo nome de outro objeto. Por exemplo, `format` é uma propriedade em um objeto range, mas `format` também é um objeto. Portanto, se você faz uma chamada como `range.load("format")`, isto é equivalente a `range.format.load()`, que é uma chamada load () vazia que pode causar problemas de desempenho, conforme descrito anteriormente. Para evitar isso, seu código deve carregar apenas  "nós folha" em uma árvore de objeto.</span><span class="sxs-lookup"><span data-stu-id="64e17-p107">Please be aware that some of the “properties” under an object may have the same name as another object. For example, `format` is a property under range object, but `format` itself is an object as well. So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()`, which is an empty load() call that can cause performance problems as outlined previously. To avoid this, your code should only load the “leaf nodes” in an object tree.</span></span> 

## <a name="suspend-calculation-temporarily"></a><span data-ttu-id="64e17-130">Suspenda os cálculos temporariamente</span><span class="sxs-lookup"><span data-stu-id="64e17-130">Suspend calculation temporarily</span></span>

<span data-ttu-id="64e17-131">Se você estiver tentando executar uma operação em um grande número de células (por exemplo, configurar o valor de um objeto range enorme) e não se importar em suspender temporariamente os cálculos no Excel até que a operação seja concluída, recomendamos suspender os cálculos até a chamada da próxima ```context.sync()```.</span><span class="sxs-lookup"><span data-stu-id="64e17-131">If you are trying to perform an operation on a large number of cells (for example, setting the value of a huge range object) and you don't mind suspending the calculation in Excel temporarily while your operation finishes, we recommend that you suspend calculation until the next ```context.sync()``` is called.</span></span>

<span data-ttu-id="64e17-p108">Confira a documentação de referência do [Objeto Application](https://docs.microsoft.com/javascript/api/excel/excel.application) para obter informações sobre como usar a API ```suspendApiCalculationUntilNextSync()``` para suspender e reativar os cálculos de uma maneira muito conveniente. O código a seguir demonstra como suspender temporariamente o cálculo:</span><span class="sxs-lookup"><span data-stu-id="64e17-p108">See [Application Object](https://docs.microsoft.com/javascript/api/excel/excel.application) reference documentation for information about how to use the ```suspendApiCalculationUntilNextSync()``` API to suspend and reactivate calculations in a very convenient way. The following code demonstrates how to suspend calculation temporarily:</span></span>

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

## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="64e17-134">Atualize todas as células em um intervalo</span><span class="sxs-lookup"><span data-stu-id="64e17-134">Update all cells in a range</span></span> 

<span data-ttu-id="64e17-p109">Quando você precisar atualizar todas as células em um intervalo com o mesmo valor ou propriedade, pode ser lento fazer isso por meio de uma matriz bidimensional que atribui o mesmo valor repetidamente, pois essa abordagem exige que o Excel itere todas as células no intervalo para definir cada uma em separado. O Excel tem uma maneira mais eficiente para atualizar todas as células em um intervalo com o mesmo valor ou propriedade.</span><span class="sxs-lookup"><span data-stu-id="64e17-p109">When you need to update all cells in a range with the same value or property, it can be slow to do this via a 2-dimensional array that repeatedly specifies the same value, since that approach requires Excel to iterate over all of the cells in the range to set each one separately. Excel has a more efficient way to update all the cells in a range with the same value or property.</span></span>

<span data-ttu-id="64e17-p110">Se você precisa aplicar o mesmo valor, a mesma formatação de número ou a mesma fórmula para um intervalo de células, é mais eficiente especificar um valor único, em vez de uma matriz de valores. Isso melhorará significativamente o desempenho. Para um exemplo de código que mostra essa abordagem em ação, confira [Conceitos fundamentais - atualizar todas as células em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span><span class="sxs-lookup"><span data-stu-id="64e17-p110">If you need to apply the same value, the same number format, or the same formula to a range of cells, it's more efficient to specify a single value instead of an array of values. Doing so will significantly improve performance. For a code sample that shows this approach in action, see [Core concepts - Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range).</span></span>

<span data-ttu-id="64e17-p111">Um cenário comum onde você pode aplicar essa abordagem é quando aplica formatos diferentes de números em colunas diferentes da planilha. Nesse caso, você pode simplesmente percorrer as colunas e definir o formato de número em cada coluna com um único valor. Trate cada coluna como um intervalo, conforme mostrado no exemplo de código de [atualização de todas as células em um intervalo](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) .</span><span class="sxs-lookup"><span data-stu-id="64e17-p111">A common scenario where you can apply this approach is when setting different number formats on different columns in a worksheet. In this case, you can simply iterate through the columns and set the number format on each column with a single value. Handle each column as a range, as shown in the [Update all cells in a range](excel-add-ins-core-concepts.md#update-all-cells-in-a-range) code sample.</span></span>

> [!NOTE]
> <span data-ttu-id="64e17-p112">Se você estiver usando TypeScript, perceberá um erro de compilação dizendo que um único valor não pode ser atribuído a uma matriz bidimensional. Isso é inevitável, pois os valores *formam* uma matriz bidimensional quando as propriedades são recuperadas e o TypeScript não permite tipos diferentes tipos setter vs getter.  No entanto, uma solução alternativa simples é definir os valores com um sufixo`as any`, por exemplo, `range.values = "hello world" as any`.</span><span class="sxs-lookup"><span data-stu-id="64e17-p112">If you're using TypeScript, you will notice a compile error saying that a single value cannot be set to a 2D array.  This is unavoidable since the values *are* a 2D array when retrieving the properties, and TypeScript does not allow different setter vs getter types.  However, a simple workaround is to set the values with a `as any` suffix, e.g., `range.values = "hello world" as any`.</span></span>

## <a name="importing-data-into-tables"></a><span data-ttu-id="64e17-146">Importar dados em tabelas</span><span class="sxs-lookup"><span data-stu-id="64e17-146">Importing data into tables</span></span>

<span data-ttu-id="64e17-p113">Ao tentar importar uma grande quantidade de dados em um objeto [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) diretamente (por exemplo, usando `TableRowCollection.add()`), você pode sofrer lentidão. Se você estiver tentando adicionar uma nova tabela, você deve preencher os dados primeiro, atribuindo `range.values` e então chamar `worksheet.tables.add()` para criar a tabela com o intervalo. Se você estiver tentando gravar dados em uma tabela existente, grave os dados em um objeto range via `table.getDataBodyRange()`, e a tabela se expandirá automaticamente.</span><span class="sxs-lookup"><span data-stu-id="64e17-p113">When trying to import a huge amount of data directly into a [Table](https://docs.microsoft.com/javascript/api/excel/excel.table) object directly (for example, by using `TableRowCollection.add()`), you might experience slow performance. If you are trying to add a new table, you should fill in the data first by setting `range.values`, and then call `worksheet.tables.add()` to create a table over the range. If you are trying to write data into an existing table, write the data into a range object via `table.getDataBodyRange()`, and the table will expand automatically.</span></span> 

<span data-ttu-id="64e17-150">Veja um exemplo dessa abordagem:</span><span class="sxs-lookup"><span data-stu-id="64e17-150">Here is an example in JavaScript of this operation.</span></span>

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
> <span data-ttu-id="64e17-151">É possível converter convenientemente um objeto Table em um objeto Range usando o método [Table.convertToRange()](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--).</span><span class="sxs-lookup"><span data-stu-id="64e17-151">You can conveniently convert a Table object to a Range object by using the [Table.convertToRange()](https://docs.microsoft.com/javascript/api/excel/excel.table#converttorange--) method.</span></span>

## <a name="enable-and-disable-events"></a><span data-ttu-id="64e17-152">Ative e desative eventos</span><span class="sxs-lookup"><span data-stu-id="64e17-152">Enable and disable agents</span></span>

<span data-ttu-id="64e17-p114">O desempenho de um suplemento pode ser melhorado com a desativação de eventos. Confira um exemplo de código mostrando como ativar e desativar eventos no artigo [Trabalhando com eventos](excel-add-ins-events.md#enable-and-disable-events) .</span><span class="sxs-lookup"><span data-stu-id="64e17-p114">Performance of an add-in may be improved by disabling events. A code sample showing how to enable and disable events is in the [Work with Events](excel-add-ins-events.md#enable-and-disable-events) article.</span></span>

## <a name="see-also"></a><span data-ttu-id="64e17-155">Confira também</span><span class="sxs-lookup"><span data-stu-id="64e17-155">See also</span></span>

- [<span data-ttu-id="64e17-156">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="64e17-156">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="64e17-157">Conceitos avançados de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="64e17-157">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
- [<span data-ttu-id="64e17-158">Especificação aberta da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="64e17-158">Excel JavaScript API Open Specification</span></span>](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)
- [<span data-ttu-id="64e17-159">Objeto Worksheet Functions (API JavaScript para Excel)</span><span class="sxs-lookup"><span data-stu-id="64e17-159">Worksheet Functions Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.functions)
