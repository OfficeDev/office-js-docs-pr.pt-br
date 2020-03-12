---
title: Trabalhar com tabelas usando a API JavaScript do Excel
description: ''
ms.date: 09/09/2019
localization_priority: Normal
ms.openlocfilehash: 4bc1b86f84b2002ad199174ffb37c3c080d51add
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596575"
---
# <a name="work-with-tables-using-the-excel-javascript-api"></a><span data-ttu-id="f735a-102">Trabalhar com tabelas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f735a-102">Work with tables using the Excel JavaScript API</span></span>

<span data-ttu-id="f735a-p101">Este artigo fornece exemplos de código que mostram como executar tarefas comuns com tabelas usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos que `Table` o `TableCollection` e os objetos dão suporte, consulte [objeto Table (API JavaScript para Excel)](/javascript/api/excel/excel.table) e [objeto TableCollection (API JavaScript para Excel)](/javascript/api/excel/excel.tablecollection).</span><span class="sxs-lookup"><span data-stu-id="f735a-p101">This article provides code samples that show how to perform common tasks with tables using the Excel JavaScript API. For the complete list of properties and methods that the `Table` and `TableCollection` objects support, see [Table Object (JavaScript API for Excel)](/javascript/api/excel/excel.table) and [TableCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.tablecollection).</span></span>

## <a name="create-a-table"></a><span data-ttu-id="f735a-105">Criar uma tabela</span><span class="sxs-lookup"><span data-stu-id="f735a-105">Create a table</span></span>

<span data-ttu-id="f735a-p102">O exemplo de código a seguir cria uma tabela na planilha chamada **Exemplo**. A tabela tem cabeçalhos e contém quatro colunas e sete linhas de dados. Se o aplicativo host do Excel em que o código está sendo executado der suporte ao [conjunto de requisito](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, a largura das colunas e a altura das linhas serão definidas para o melhor ajuste aos dados atuais da tabela.</span><span class="sxs-lookup"><span data-stu-id="f735a-p102">The following code sample creates a table in the worksheet named **Sample**. The table has headers and contains four columns and seven rows of data. If the Excel host application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="f735a-109">Para especificar um nome para uma tabela, primeiro você deve criar a tabela e, em seguida `name` , definir sua propriedade, conforme mostrado no exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="f735a-109">To specify a name for a table, you must first create the table and then set its `name` property, as shown in the following example.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-110">**Nova tabela**</span><span class="sxs-lookup"><span data-stu-id="f735a-110">**New table**</span></span>

![Nova tabela no Excel](../images/excel-tables-create.png)

## <a name="add-rows-to-a-table"></a><span data-ttu-id="f735a-112">Adicionar linhas a uma tabela</span><span class="sxs-lookup"><span data-stu-id="f735a-112">Add rows to a table</span></span>

<span data-ttu-id="f735a-p103">O exemplo de código a seguir adiciona sete novas linhas à tabela **ExpensesTable** na planilha **Exemplo**. As novas linhas são adicionadas ao fim da tabela. Se o aplicativo host do Excel em que o código está sendo executado der suporte ao [conjunto de requisito](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, a largura das colunas e a altura das linhas serão definidas para o melhor ajuste aos dados atuais da tabela.</span><span class="sxs-lookup"><span data-stu-id="f735a-p103">The following code sample adds seven new rows to the table named **ExpensesTable** within the worksheet named **Sample**. The new rows are added to the end of the table. If the Excel host application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

> [!NOTE]
> <span data-ttu-id="f735a-p104">A `index` propriedade de um objeto [TableRow](/javascript/api/excel/excel.tablerow) indica o número de índice da linha dentro da coleção Rows da tabela. Um `TableRow` objeto não contém uma `id` propriedade que pode ser usada como uma chave exclusiva para identificar a linha.</span><span class="sxs-lookup"><span data-stu-id="f735a-p104">The `index` property of a [TableRow](/javascript/api/excel/excel.tablerow) object indicates the index number of the row within the rows collection of the table. A `TableRow` object does not contain an `id` property that can be used as a unique key to identify the row.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.rows.add(null /*add rows to the end of the table*/, [
        ["1/16/2017", "THE PHONE COMPANY", "Communications", "$120"],
        ["1/20/2017", "NORTHWIND ELECTRIC CARS", "Transportation", "$142"],
        ["1/20/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$27"],
        ["1/21/2017", "COHO VINEYARD", "Restaurant", "$33"],
        ["1/25/2017", "BELLOWS COLLEGE", "Education", "$350"],
        ["1/28/2017", "TREY RESEARCH", "Other", "$135"],
        ["1/31/2017", "BEST FOR YOU ORGANICS COMPANY", "Groceries", "$97"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-118">**Tabela com novas linhas**</span><span class="sxs-lookup"><span data-stu-id="f735a-118">**Table with new rows**</span></span>

![Tabela com novas linhas no Excel](../images/excel-tables-add-rows.png)

## <a name="add-a-column-to-a-table"></a><span data-ttu-id="f735a-120">Adicionar uma coluna a uma tabela</span><span class="sxs-lookup"><span data-stu-id="f735a-120">Add a column to a table</span></span>

<span data-ttu-id="f735a-p105">Estes exemplos mostram como adicionar uma coluna a uma tabela. O primeiro exemplo preenche a nova coluna com valores estáticos. O segundo exemplo popula a nova coluna com fórmulas.</span><span class="sxs-lookup"><span data-stu-id="f735a-p105">These examples show how to add a column to a table. The first example populates the new column with static values; the second example populates the new column with formulas.</span></span>

> [!NOTE]
> <span data-ttu-id="f735a-p106">A propriedade **index** de um objeto [TableColumn](/javascript/api/excel/excel.tablecolumn) indica o número de índice da coluna no conjunto de colunas da tabela. A propriedade **id** de um objeto **TableColumn** contém uma chave exclusiva que identifica a coluna.</span><span class="sxs-lookup"><span data-stu-id="f735a-p106">The **index** property of a [TableColumn](/javascript/api/excel/excel.tablecolumn) object indicates the index number of the column within the columns collection of the table. The **id** property of a **TableColumn** object contains a unique key that identifies the column.</span></span>

### <a name="add-a-column-that-contains-static-values"></a><span data-ttu-id="f735a-125">Adicionar uma coluna que contém valores estáticos</span><span class="sxs-lookup"><span data-stu-id="f735a-125">Add a column that contains static values</span></span>

<span data-ttu-id="f735a-p107">O exemplo de código a seguir adiciona uma nova coluna à tabela **ExpensesTable** na planilha **Exemplo**. A nova coluna é adicionada após todas as colunas existentes na tabela e contém um cabeçalho ("Dia da Semana"), bem como dados para popular as células na coluna. Se o aplicativo host do Excel em que o código está sendo executado der suporte ao [conjunto de requisito](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, a largura das colunas e a altura das linhas serão definidas para o melhor ajuste aos dados atuais da tabela.</span><span class="sxs-lookup"><span data-stu-id="f735a-p107">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**. The new column is added after all existing columns in the table and contains a header ("Day of the Week") as well as data to populate the cells in the column. If the Excel host application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.columns.add(null /*add columns to the end of the table*/, [
        ["Day of the Week"],
        ["Saturday"],
        ["Friday"],
        ["Monday"],
        ["Thursday"],
        ["Sunday"],
        ["Saturday"],
        ["Monday"]
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-129">**Tabela com nova coluna**</span><span class="sxs-lookup"><span data-stu-id="f735a-129">**Table with new column**</span></span>

![Tabela com nova coluna no Excel](../images/excel-tables-add-column.png)

### <a name="add-a-column-that-contains-formulas"></a><span data-ttu-id="f735a-131">Adicionar uma coluna que contém fórmulas</span><span class="sxs-lookup"><span data-stu-id="f735a-131">Add a column that contains formulas</span></span>

<span data-ttu-id="f735a-p108">O exemplo de código a seguir adiciona uma nova coluna à tabela **ExpensesTable** na planilha **Exemplo**. A nova coluna é adicionada ao fim da tabela, contém um cabeçalho ("Tipo do Dia") e usa uma fórmula para popular cada célula na coluna de dados. Se o aplicativo host do Excel em que o código está sendo executado der suporte ao [conjunto de requisito](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, a largura das colunas e a altura das linhas serão definidas para o melhor ajuste aos dados atuais da tabela.</span><span class="sxs-lookup"><span data-stu-id="f735a-p108">The following code sample adds a new column to the table named **ExpensesTable** within the worksheet named **Sample**. The new column is added to the end of the table, contains a header ("Type of the Day"), and uses a formula to populate each data cell in the column. If the Excel host application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.columns.add(null /*add columns to the end of the table*/, [
        ["Type of the Day"],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")'],
        ['=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")']
    ]);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-135">**Tabela com nova coluna calculada**</span><span class="sxs-lookup"><span data-stu-id="f735a-135">**Table with new calculated column**</span></span>

![Tabela com nova coluna calculada no Excel](../images/excel-tables-add-calculated-column.png)

## <a name="update-column-name"></a><span data-ttu-id="f735a-137">Atualizar o nome da coluna</span><span class="sxs-lookup"><span data-stu-id="f735a-137">Update column name</span></span>

<span data-ttu-id="f735a-p109">O exemplo de código a seguir atualiza o nome da primeira coluna da tabela para **Data da compra**. Se o aplicativo host do Excel em que o código está sendo executado der suporte ao [conjunto de requisitos](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, a largura das colunas e a altura das linhas serão definidas para o melhor ajuste aos dados atuais da tabela.</span><span class="sxs-lookup"><span data-stu-id="f735a-p109">The following code sample updates the name of the first column in the table to **Purchase date**. If the Excel host application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var expensesTable = sheet.tables.getItem("ExpensesTable");
    expensesTable.columns.load("items");

    return context.sync()
        .then(function () {
            expensesTable.columns.items[0].name = "Purchase date";

            if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                sheet.getUsedRange().format.autofitColumns();
                sheet.getUsedRange().format.autofitRows();
            }

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-140">**Tabela com novo nome de coluna**</span><span class="sxs-lookup"><span data-stu-id="f735a-140">**Table with new column name**</span></span>

![Tabela com novo nome de coluna no Excel](../images/excel-tables-update-column-name.png)

## <a name="get-data-from-a-table"></a><span data-ttu-id="f735a-142">Obter dados de uma tabela</span><span class="sxs-lookup"><span data-stu-id="f735a-142">Get data from a table</span></span>

<span data-ttu-id="f735a-143">O exemplo de código a seguir lê dados de uma tabela chamada **ExpensesTable** na planilha **Exemplo** e inclui esses dados abaixo da tabela na mesma planilha.</span><span class="sxs-lookup"><span data-stu-id="f735a-143">The following code sample reads data from a table named **ExpensesTable** in the worksheet named **Sample** and then outputs that data below the table in the same worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Get data from the header row
    var headerRange = expensesTable.getHeaderRowRange().load("values");

    // Get data from the table
    var bodyRange = expensesTable.getDataBodyRange().load("values");

    // Get data from a single column
    var columnRange = expensesTable.columns.getItem("Merchant").getDataBodyRange().load("values");

    // Get data from a single row
    var rowRange = expensesTable.rows.getItemAt(1).load("values");

    // Sync to populate proxy objects with data from Excel
    return context.sync()
        .then(function () {
            var headerValues = headerRange.values;
            var bodyValues = bodyRange.values;
            var merchantColumnValues = columnRange.values;
            var secondRowValues = rowRange.values;

            // Write data from table back to the sheet
            sheet.getRange("A11:A11").values = [["Results"]];
            sheet.getRange("A13:D13").values = headerValues;
            sheet.getRange("A14:D20").values = bodyValues;
            sheet.getRange("B23:B29").values = merchantColumnValues;
            sheet.getRange("A32:D32").values = secondRowValues;

            // Sync to update the sheet in Excel
            return context.sync();
        });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-144">**Tabela e saída de dados**</span><span class="sxs-lookup"><span data-stu-id="f735a-144">**Table and data output**</span></span>

![Dados de tabela no Excel](../images/excel-tables-get-data.png)

## <a name="detect-data-changes"></a><span data-ttu-id="f735a-146">Detectar as alterações dos dados</span><span class="sxs-lookup"><span data-stu-id="f735a-146">Detect data changes</span></span>

<span data-ttu-id="f735a-147">O suplemento precisará reagir aos usuários alterando os dados em uma tabela.</span><span class="sxs-lookup"><span data-stu-id="f735a-147">Your add-in may need to react to users changing the data in a table.</span></span> <span data-ttu-id="f735a-148">Para detectar essas alterações, basta [Registrar um manipulador de eventos.](excel-add-ins-events.md#register-an-event-handler) para o `onChanged` evento da tabela.</span><span class="sxs-lookup"><span data-stu-id="f735a-148">To detect these changes, you can [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the `onChanged` event of a table.</span></span> <span data-ttu-id="f735a-149">Manipuladores de eventos para o `onChanged` evento recebem um objeto [TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs) quando o evento é acionado.</span><span class="sxs-lookup"><span data-stu-id="f735a-149">Event handlers for the `onChanged` event receive a [TableChangedEventArgs](/javascript/api/excel/excel.tablechangedeventargs) object when the event fires.</span></span>

<span data-ttu-id="f735a-150">O `TableChangedEventArgs` objeto fornece informações sobre as alterações e a fonte.</span><span class="sxs-lookup"><span data-stu-id="f735a-150">The `TableChangedEventArgs` object provides information about the changes and the source.</span></span> <span data-ttu-id="f735a-151">Como `onChanged` o acionamento ocorre quando o formato ou o valor dos dados mudam, pode ser útil checar com o suplemento se os valores realmente foram alterados.</span><span class="sxs-lookup"><span data-stu-id="f735a-151">Since `onChanged` fires when either the format or value of the data changes, it can be useful to have your add-in check if the values have actually changed.</span></span> <span data-ttu-id="f735a-152">A `details` propriedade encapsula estas informações como um [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span><span class="sxs-lookup"><span data-stu-id="f735a-152">The `details` property encapsulates this information as a [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span></span> <span data-ttu-id="f735a-153">O exemplo a seguir mostra como exibir o antes e depois dos valores e tipos de uma célula que foi alterada.</span><span class="sxs-lookup"><span data-stu-id="f735a-153">The following code sample shows how to display the before and after values and types of a cell that has been changed.</span></span>

```js
// This function would be used as an event handler for the Table.onChanged event.
function onTableChanged(eventArgs) {
    Excel.run(function (context) {
        var details = eventArgs.details;
        var address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        return context.sync();
    });
}
```

## <a name="sort-data-in-a-table"></a><span data-ttu-id="f735a-154">Classificar dados em uma tabela</span><span class="sxs-lookup"><span data-stu-id="f735a-154">Sort data in a table</span></span>

<span data-ttu-id="f735a-155">O exemplo de código a seguir classifica os dados da tabela em ordem decrescente de acordo com os valores na quarta coluna da tabela.</span><span class="sxs-lookup"><span data-stu-id="f735a-155">The following code sample sorts table data in descending order according to the values in the fourth column of the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to sort data by the fourth column of the table (descending)
    var sortRange = expensesTable.getDataBodyRange();
    sortRange.sort.apply([
        {
            key: 3,
            ascending: false,
        },
    ]);

    // Sync to run the queued command in Excel
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-156">**Dados de tabela classificados por Valor (decrescente)**</span><span class="sxs-lookup"><span data-stu-id="f735a-156">**Table data sorted by Amount (descending)**</span></span>

![Dados de tabela no Excel](../images/excel-tables-sort.png)

<span data-ttu-id="f735a-158">Quando os dados são classificados em uma planilha, uma notificação de evento é acionada.</span><span class="sxs-lookup"><span data-stu-id="f735a-158">When data is sorted in a worksheet, an event notification fires.</span></span> <span data-ttu-id="f735a-159">Para saber mais sobre os eventos relacionados à classificação e como seu suplemento pode registrar manipuladores de eventos para responder a esses eventos, consulte [Manipular eventos de classificação](excel-add-ins-worksheets.md#handle-sorting-events).</span><span class="sxs-lookup"><span data-stu-id="f735a-159">To learn more about sort-related events and how your add-in can register event handlers to respond to such events, see [Handle sorting events](excel-add-ins-worksheets.md#handle-sorting-events).</span></span>

## <a name="apply-filters-to-a-table"></a><span data-ttu-id="f735a-160">Aplicar filtros a uma tabela</span><span class="sxs-lookup"><span data-stu-id="f735a-160">Apply filters to a table</span></span>

<span data-ttu-id="f735a-p113">O exemplo de código a seguir aplica filtros à coluna **Valor** e à coluna **Categoria** em uma tabela. Como resultado dos filtros, são mostradas apenas linhas em que **Categoria** é um dos valores especificados e **Valor** está abaixo do valor médio para todas as linhas.</span><span class="sxs-lookup"><span data-stu-id="f735a-p113">The following code sample applies filters to the **Amount** column and the **Category** column within a table. As a result of the filters, only rows where **Category** is one of the specified values and **Amount** is below the average value for all rows is shown.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    // Queue a command to apply a filter on the Category column
    filter = expensesTable.columns.getItem("Category").filter;
    filter.apply({
        filterOn: Excel.FilterOn.values,
        values: ["Restaurant", "Groceries"]
    });

    // Queue a command to apply a filter on the Amount column
    var filter = expensesTable.columns.getItem("Amount").filter;
    filter.apply({
        filterOn: Excel.FilterOn.dynamic,
        dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    // Sync to run the queued commands in Excel
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-163">**Dados de tabela com filtros aplicados para Categoria e Valor**</span><span class="sxs-lookup"><span data-stu-id="f735a-163">**Table data with filters applied for Category and Amount**</span></span>

![Dados de tabela filtrados no Excel](../images/excel-tables-filters-apply.png)

## <a name="clear-table-filters"></a><span data-ttu-id="f735a-165">Limpar filtros de tabela</span><span class="sxs-lookup"><span data-stu-id="f735a-165">Clear table filters</span></span>

<span data-ttu-id="f735a-166">O exemplo de código a seguir limpa todos os filtros aplicados atualmente à tabela.</span><span class="sxs-lookup"><span data-stu-id="f735a-166">The following code sample clears any filters currently applied on the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.clearFilters();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-167">**Dados de tabela sem filtros aplicados**</span><span class="sxs-lookup"><span data-stu-id="f735a-167">**Table data with no filters applied**</span></span>

![Dados de tabela não filtrados no Excel](../images/excel-tables-filters-clear.png)

## <a name="get-the-visible-range-from-a-filtered-table"></a><span data-ttu-id="f735a-169">Obter o intervalo visível de uma tabela filtrada</span><span class="sxs-lookup"><span data-stu-id="f735a-169">Get the visible range from a filtered table</span></span>

<span data-ttu-id="f735a-p114">O exemplo de código a seguir obtém um intervalo que contém dados somente para as células que estão visíveis no momento dentro da tabela especificada e, em seguida, grava os valores desse intervalo no console. Você pode usar o `getVisibleView()` método conforme mostrado abaixo para obter o conteúdo visível de uma tabela sempre que os filtros de coluna tiverem sido aplicados.</span><span class="sxs-lookup"><span data-stu-id="f735a-p114">The following code sample gets a range that contains data only for cells that are currently visible within the specified table, and then writes the values of that range to the console. You can use the `getVisibleView()` method as shown below to get the visible contents of a table whenever column filters have been applied.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    var visibleRange = expensesTable.getDataBodyRange().getVisibleView();
    visibleRange.load("values");

    return context.sync()
        .then(function() {
            console.log(visibleRange.values);
        });
}).catch(errorHandlerFunction);
```

## <a name="autofilter"></a><span data-ttu-id="f735a-172">Filtro Automático</span><span class="sxs-lookup"><span data-stu-id="f735a-172">AutoFilter</span></span>

<span data-ttu-id="f735a-173">Um suplemento pode usar o objeto [AutoFilter](/javascript/api/excel/excel.autofilter) da tabela para filtrar dados.</span><span class="sxs-lookup"><span data-stu-id="f735a-173">An add-in can use the table's [AutoFilter](/javascript/api/excel/excel.autofilter) object to filter data.</span></span> <span data-ttu-id="f735a-174">Um `AutoFilter` objeto é toda a estrutura de filtro de uma tabela ou intervalo.</span><span class="sxs-lookup"><span data-stu-id="f735a-174">An `AutoFilter` object is the entire filter structure of a table or range.</span></span> <span data-ttu-id="f735a-175">Todas as operações de filtros abordadas anteriormente neste artigo são compatíveis com o filtro automático.</span><span class="sxs-lookup"><span data-stu-id="f735a-175">All of the filter operations discussed earlier in this article are compatible with the auto-filter.</span></span> <span data-ttu-id="f735a-176">O ponto de acesso único facilita o acesso e o gerenciamento de múltiplos filtros.</span><span class="sxs-lookup"><span data-stu-id="f735a-176">The single access point does make it easier to access and manage multiple filters.</span></span>

<span data-ttu-id="f735a-177">O exemplo de código a seguir mostra a mesma [filtragem de dados como o exemplo de código anterior](#apply-filters-to-a-table), mas concluído totalmente pelo filtro automático.</span><span class="sxs-lookup"><span data-stu-id="f735a-177">The following code sample shows the same [data filtering as the earlier code sample](#apply-filters-to-a-table), but done entirely through the auto-filter.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.autoFilter.apply(expensesTable.getRange(), 2, {
        filterOn: Excel.FilterOn.values,
        values: ["Restaurant", "Groceries"]
    });
    expensesTable.autoFilter.apply(expensesTable.getRange(), 3, {
        filterOn: Excel.FilterOn.dynamic,
        dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
    });

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-178">Um `AutoFilter` também pode ser aplicado a um intervalo no nível da planilha.</span><span class="sxs-lookup"><span data-stu-id="f735a-178">An `AutoFilter` can also be applied to a range at the worksheet level.</span></span> <span data-ttu-id="f735a-179">Consulte [Trabalhar com tabelas usando o API JavaScript do Excel](excel-add-ins-worksheets.md#filter-data) para mais informações.</span><span class="sxs-lookup"><span data-stu-id="f735a-179">See [Work with worksheets using the Excel JavaScript API](excel-add-ins-worksheets.md#filter-data) for more information.</span></span>

## <a name="format-a-table"></a><span data-ttu-id="f735a-180">Formatar uma tabela</span><span class="sxs-lookup"><span data-stu-id="f735a-180">Format a table</span></span>

<span data-ttu-id="f735a-p117">O código de exemplo a seguir aplica formatação a uma tabela. Ele especifica cores de preenchimento diferentes para a linha de cabeçalho, o corpo, a segunda linha e a primeira coluna da tabela. Para obter informações sobre as propriedades que você pode usar para especificar o formato, confira [Objeto RangeFormat (API do JavaScript para Excel)](/javascript/api/excel/excel.rangeformat).</span><span class="sxs-lookup"><span data-stu-id="f735a-p117">The following code sample applies formatting to a table. It specifies different fill colors for the header row of the table, the body of the table, the second row of the table, and the first column of the table. For information about the properties you can use to specify format, see [RangeFormat Object (JavaScript API for Excel)](/javascript/api/excel/excel.rangeformat).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.getHeaderRowRange().format.fill.color = "#C70039";
    expensesTable.getDataBodyRange().format.fill.color = "#DAF7A6";
    expensesTable.rows.getItemAt(1).getRange().format.fill.color = "#FFC300";
    expensesTable.columns.getItemAt(0).getDataBodyRange().format.fill.color = "#FFA07A";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-184">**Tabela depois que a formatação é aplicada**</span><span class="sxs-lookup"><span data-stu-id="f735a-184">**Table after formatting is applied**</span></span>

![Tabela depois que a formatação é aplicada no Excel](../images/excel-tables-formatting-after.png)

## <a name="convert-a-range-to-a-table"></a><span data-ttu-id="f735a-186">Converter um intervalo em uma tabela</span><span class="sxs-lookup"><span data-stu-id="f735a-186">Convert a range to a table</span></span>

<span data-ttu-id="f735a-187">O exemplo de código a seguir cria um intervalo de dados e o converte em uma tabela.</span><span class="sxs-lookup"><span data-stu-id="f735a-187">The following code sample creates a range of data and then converts that range to a table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Define values for the range
    var values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
    ["Frames", 5000, 7000, 6544, 4377],
    ["Saddles", 400, 323, 276, 651],
    ["Brake levers", 12000, 8766, 8456, 9812],
    ["Chains", 1550, 1088, 692, 853],
    ["Mirrors", 225, 600, 923, 544],
    ["Spokes", 6005, 7634, 4589, 8765]];

    // Create the range
    var range = sheet.getRange("A1:E7");
    range.values = values;

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    // Convert the range to a table
    var expensesTable = sheet.tables.add('A1:E7', true);
    expensesTable.name = "ExpensesTable";

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-188">**Dados no intervalo (antes que o intervalo seja convertido em uma tabela)**</span><span class="sxs-lookup"><span data-stu-id="f735a-188">**Data in the range (before the range is converted to a table)**</span></span>

![Dados no intervalo no Excel](../images/excel-ranges.png)

<span data-ttu-id="f735a-190">**Dados da tabela (depois que o intervalo é convertido em uma tabela)**</span><span class="sxs-lookup"><span data-stu-id="f735a-190">**Data in the table (after the range is converted to a table)**</span></span>

![Dados na tabela no Excel](../images/excel-tables-from-range.png)

## <a name="import-json-data-into-a-table"></a><span data-ttu-id="f735a-192">Importar dados JSON em uma tabela</span><span class="sxs-lookup"><span data-stu-id="f735a-192">Import JSON data into a table</span></span>

<span data-ttu-id="f735a-p118">O exemplo de código a seguir cria uma tabela na planilha **Exemplo** e popula a tabela usando um objeto JSON que define duas linhas de dados. Se o aplicativo host do Excel em que o código está sendo executado der suporte ao [conjunto de requisito](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, a largura das colunas e a altura das linhas serão definidas para o melhor ajuste aos dados atuais da tabela.</span><span class="sxs-lookup"><span data-stu-id="f735a-p118">The following code sample creates a table in the worksheet named **Sample** and then populates the table by using a JSON object that defines two rows of data. If the Excel host application where the code is running supports [requirement set](../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, the width of the columns and height of the rows are set to best fit the current data in the table.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    var transactions = [
      {
        "DATE": "1/1/2017",
        "MERCHANT": "The Phone Company",
        "CATEGORY": "Communications",
        "AMOUNT": "$120"
      },
      {
        "DATE": "1/1/2017",
        "MERCHANT": "Southridge Video",
        "CATEGORY": "Entertainment",
        "AMOUNT": "$40"
      }
    ];

    var newData = transactions.map(item =>
        [item.DATE, item.MERCHANT, item.CATEGORY, item.AMOUNT]);

    expensesTable.rows.add(null, newData);

    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f735a-195">**Nova tabela**</span><span class="sxs-lookup"><span data-stu-id="f735a-195">**New table**</span></span>

![Nova tabela no Excel](../images/excel-tables-create-from-json.png)

## <a name="see-also"></a><span data-ttu-id="f735a-197">Confira também</span><span class="sxs-lookup"><span data-stu-id="f735a-197">See also</span></span>

- [<span data-ttu-id="f735a-198">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f735a-198">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
