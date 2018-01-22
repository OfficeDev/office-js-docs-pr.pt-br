# <a name="work-with-tables-using-the-excel-javascript-api"></a>Trabalhar com tabelas usando a API JavaScript do Excel

Este artigo fornece exemplos de código que mostram como executar tarefas comuns com tabelas usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos aos quais os objetos **Table** e **TableCollection** dão suporte, confira [Objeto Table (API do JavaScript para Excel)](../../reference/excel/table.md) e [Objeto TableCollection (API do JavaScript para Excel)](../../reference/excel/tablecollection.md).

## <a name="create-a-table"></a>Criar uma tabela

O exemplo de código a seguir cria uma tabela na planilha chamada **Exemplo**. A tabela tem cabeçalhos e contém quatro colunas e sete linhas de dados. Se o aplicativo host do Excel em que o código está sendo executado der suporte ao [conjunto de requisito](../../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, a largura das colunas e a altura das linhas serão definidas para o melhor ajuste aos dados atuais da tabela.

>**Observação**: Para especificar um nome para uma tabela, primeiro crie a tabela e defina sua propriedade **name**, conforme mostrado no exemplo a seguir.

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

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Nova tabela**

![Nova tabela no Excel](../images/Excel-table-create.png)

## <a name="add-rows-to-a-table"></a>Adicionar linhas a uma tabela

O exemplo de código a seguir adiciona sete novas linhas à tabela **ExpensesTable** na planilha **Exemplo**. As novas linhas são adicionadas ao fim da tabela. Se o aplicativo host do Excel em que o código está sendo executado der suporte ao [conjunto de requisito](../../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, a largura das colunas e a altura das linhas serão definidas para o melhor ajuste aos dados atuais da tabela.

>**Observação**: A propriedade **index** de um objeto [column](../../reference/excel/tablerow.md) indica o número de índice da linha no conjunto de linhas da tabela. Um objeto **TableRow** não contém uma propriedade **id** que pode ser usada como chave exclusiva para identificar a linha.

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

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

**Tabela com novas linhas**

![Tabela com novas linhas no Excel](../images/Excel-table-add-rows.png)

## <a name="add-a-column-to-a-table"></a>Adicionar uma coluna a uma tabela

Estes exemplos mostram como adicionar uma coluna a uma tabela. O primeiro exemplo popula a nova coluna com valores estáticos. O segundo exemplo popula a nova coluna com fórmulas.

>**Observação**: A propriedade **index** de um objeto [TableColumn](../../reference/excel/tablecolumn.md) indica o número de índice da coluna no conjunto de colunas da tabela. A propriedade **id** de um objeto **TableColumn** contém uma chave exclusiva que identifica a coluna.

### <a name="add-a-column-that-contains-static-values"></a>Adicionar uma coluna que contém valores estáticos

O exemplo de código a seguir adiciona uma nova coluna à tabela **ExpensesTable** na planilha **Exemplo**. A nova coluna é adicionada após todas as colunas existentes na tabela e contém um cabeçalho ("Dia da Semana"), bem como dados para popular as células na coluna. Se o aplicativo host do Excel em que o código está sendo executado der suporte ao [conjunto de requisito](../../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, a largura das colunas e a altura das linhas serão definidas para o melhor ajuste aos dados atuais da tabela.

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

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

**Tabela com nova coluna**

![Tabela com nova coluna no Excel](../images/Excel-table-add-column.png)

### <a name="add-a-column-that-contains-formulas"></a>Adicionar uma coluna que contém fórmulas

O exemplo de código a seguir adiciona uma nova coluna à tabela **ExpensesTable** na planilha **Exemplo**. A nova coluna é adicionada ao fim da tabela, contém um cabeçalho ("Tipo do Dia") e usa uma fórmula para popular cada célula na coluna de dados. Se o aplicativo host do Excel em que o código está sendo executado der suporte ao [conjunto de requisito](../../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, a largura das colunas e a altura das linhas serão definidas para o melhor ajuste aos dados atuais da tabela.

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

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    return context.sync();
}).catch(errorHandlerFunction);
```

**Tabela com nova coluna calculada**

![Tabela com nova coluna calculada no Excel](../images/Excel-table-add-calculated-column.png)

## <a name="update-column-name"></a>Atualizar o nome da coluna

O exemplo de código a seguir atualiza o nome da primeira coluna da tabela para **Data da compra**. Se o aplicativo host do Excel em que o código está sendo executado der suporte ao [conjunto de requisito](../../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, a largura das colunas e a altura das linhas serão definidas para o melhor ajuste aos dados atuais da tabela.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var expensesTable = sheet.tables.getItem("ExpensesTable");
    expensesTable.columns.load("items");

    return context.sync()
        .then(function () {
            expensesTable.columns.items[0].name = "Purchase date";

            if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
                sheet.getUsedRange().format.autofitColumns();
                sheet.getUsedRange().format.autofitRows();
            }

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

**Tabela com novo nome de coluna**

![Tabela com novo nome de coluna no Excel](../images/Excel-table-update-column-name.png)

## <a name="get-data-from-a-table"></a>Obter dados de uma tabela

O exemplo de código a seguir lê dados de uma tabela chamada **ExpensesTable** na planilha **Exemplo** e inclui esses dados abaixo da tabela na mesma planilha.

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

**Tabela e saída de dados**

![Dados de tabela no Excel](../images/Excel-table-get-data.png)

## <a name="sort-data-in-a-table"></a>Classificar dados em uma tabela

O exemplo de código a seguir classifica os dados da tabela em ordem decrescente de acordo com os valores na quarta coluna da tabela.

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

**Dados de tabela classificados por Valor (decrescente)**

![Dados de tabela no Excel](../images/Excel-table-sort.png)

## <a name="apply-filters-to-a-table"></a>Aplicar filtros a uma tabela

O exemplo de código a seguir aplica filtros à coluna **Valor** e à coluna **Categoria** em uma tabela. Como resultado dos filtros, são mostradas apenas linhas em que **Categoria** é um dos valores especificados e **Valor** está abaixo do valor médio para todas as linhas.

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

**Dados de tabela com filtros aplicados para Categoria e Valor**

![Dados de tabela filtrados no Excel](../images/Excel-table-filters-apply.png)

## <a name="clear-table-filters"></a>Limpar filtros de tabela

O exemplo de código a seguir limpa todos os filtros aplicados atualmente à tabela.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var expensesTable = sheet.tables.getItem("ExpensesTable");

    expensesTable.clearFilters();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Dados de tabela sem filtros aplicados**

![Dados de tabela não filtrados no Excel](../images/Excel-table-filters-clear.png)

## <a name="get-the-visible-range-from-a-filtered-table"></a>Obter o intervalo visível de uma tabela filtrada

O exemplo de código a seguir obtém um intervalo que contém dados somente para células que estão visíveis atualmente na tabela especificada e grava os valores do intervalo no console. Você pode usar o método **getVisibleView()** conforme mostrado abaixo para obter o conteúdo visível de uma tabela sempre que filtros de coluna tiverem sido aplicados.

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

## <a name="format-a-table"></a>Formatar uma tabela

O código de exemplo a seguir aplica formatação a uma tabela. Ele especifica cores de preenchimento diferentes para a linha de cabeçalho, o corpo, a segunda linha e a primeira coluna da tabela. Para obter informações sobre as propriedades que você pode usar para especificar o formato, confira [Objeto RangeFormat (API do JavaScript para Excel)](../../reference/excel/rangeformat.md).

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

**Tabela depois que a formatação é aplicada**

![Tabela depois que a formatação é aplicada no Excel](../images/Excel-table-formatting-after.png)

## <a name="convert-a-range-to-a-table"></a>Converter um intervalo em uma tabela

O exemplo de código a seguir cria um intervalo de dados e o converte em uma tabela.

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

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
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

**Dados no intervalo (antes que o intervalo seja convertido em uma tabela)**

![Dados no intervalo no Excel](../images/Excel-range.png)

**Dados da tabela (depois que o intervalo é convertido em uma tabela)**

![Dados na tabela no Excel](../images/Excel-table-from-range.png)

## <a name="import-json-data-into-a-table"></a>Importar dados JSON em uma tabela

O exemplo de código a seguir cria uma tabela na planilha **Exemplo** e popula a tabela usando um objeto JSON que define duas linhas de dados. Se o aplicativo host do Excel em que o código está sendo executado der suporte ao [conjunto de requisito](../../reference/requirement-sets/excel-api-requirement-sets.md) **ExcelApi 1.2**, a largura das colunas e a altura das linhas serão definidas para o melhor ajuste aos dados atuais da tabela.

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

    if (Office.context.requirements.isSetSupported("ExcelApi", 1.2)) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
    }

    sheet.activate();

    return context.sync();
}).catch(errorHandlerFunction);
```

**Nova tabela**

![Nova tabela no Excel](../images/Excel-table-create-from-json.png)

## <a name="additional-resources"></a>Recursos adicionais

- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Objeto Table (API JavaScript para Excel)](../../reference/excel/table.md)
- [Objeto TableCollection (API JavaScript para Excel)](../../reference/excel/tablecollection.md)
