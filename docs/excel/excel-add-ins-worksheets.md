---
title: Trabalhe com planilhas usando a API JavaScript do Excel
description: ''
ms.date: 09/09/2019
localization_priority: Priority
ms.openlocfilehash: 3c06e3660c2c8d6bf362b38185b96c8012dc4b90
ms.sourcegitcommit: 24303ca235ebd7144a1d913511d8e4fb7c0e8c0d
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2019
ms.locfileid: "36838557"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="a083a-102">Trabalhe com planilhas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="a083a-102">Work with worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="a083a-p101">Este artigo fornece exemplos de código que mostram como executar tarefas comuns com planilhas usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos aos quais os objetos **Worksheet** e **WorksheetCollection** dão suporte, confira [Objeto Worksheet (API JavaScript para Excel)](/javascript/api/excel/excel.worksheet) e [Objeto WorksheetCollection (API JavaScript para Excel)](/javascript/api/excel/excel.worksheetcollection).</span><span class="sxs-lookup"><span data-stu-id="a083a-p101">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API. For the complete list of properties and methods that the **Worksheet** and **WorksheetCollection** objects support, see [Worksheet Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheet) and [WorksheetCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheetcollection).</span></span>

> [!NOTE]
> <span data-ttu-id="a083a-105">As informações deste artigo se aplicam apenas a planilhas regulares; elas não se aplicam às folhas "gráfico" ou "macro".</span><span class="sxs-lookup"><span data-stu-id="a083a-105">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="a083a-106">Obter planilhas</span><span class="sxs-lookup"><span data-stu-id="a083a-106">Get worksheets</span></span>

<span data-ttu-id="a083a-107">O exemplo de código a seguir obtém a coleção de planilhas, carrega a propriedade **name** de cada planilha e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="a083a-107">The following code sample gets the collection of worksheets, loads the **name** property of each worksheet, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
        .then(function () {
            if (sheets.items.length > 1) {
                console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
            } else {
                console.log(`There is one worksheet in the workbook:`);
            }
            for (var i in sheets.items) {
                console.log(sheets.items[i].name);
            }
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="a083a-p102">A propriedade **id** de uma planilha identifica exclusivamente a planilha em uma determinada pasta de trabalho e seu valor permanecerá igual, mesmo quando a planilha for renomeada ou movida. Quando uma planilha é excluída de uma pasta de trabalho no Excel para Mac, a **id** da planilha excluída pode ser reatribuída a uma nova planilha que é subsequentemente criada.</span><span class="sxs-lookup"><span data-stu-id="a083a-p102">The **id** property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved. When a worksheet is deleted from a workbook in Excel for Mac, the **id** of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="a083a-110">Obter a planilha ativa</span><span class="sxs-lookup"><span data-stu-id="a083a-110">Get the active worksheet</span></span>

<span data-ttu-id="a083a-111">O exemplo de código a seguir obtém a planilha ativa, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="a083a-111">The following code sample gets the active worksheet, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The active worksheet is "${sheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-the-active-worksheet"></a><span data-ttu-id="a083a-112">Definir a planilha ativa</span><span class="sxs-lookup"><span data-stu-id="a083a-112">Set the active worksheet</span></span>

<span data-ttu-id="a083a-p103">O exemplo de código a seguir define a planilha ativa para a planilha chamada **Amostra**, carrega sua propriedade **name** e grava uma mensagem no console. Se não houver planilha com esse nome, o método **activate()** gerará um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="a083a-p103">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its **name** property, and writes a message to the console. If there is no worksheet with that name, the **activate()** method throws an **ItemNotFound** error.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.activate();
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The active worksheet is "${sheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="a083a-115">Planilhas de referência por posição relativa</span><span class="sxs-lookup"><span data-stu-id="a083a-115">Reference worksheets by relative position</span></span>

<span data-ttu-id="a083a-116">Esses exemplos mostram como fazer referência a uma planilha por sua posição relativa.</span><span class="sxs-lookup"><span data-stu-id="a083a-116">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="a083a-117">Obter a primeira planilha</span><span class="sxs-lookup"><span data-stu-id="a083a-117">Get the first worksheet</span></span>

<span data-ttu-id="a083a-118">O exemplo de código a seguir obtém a primeira planilha na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="a083a-118">The following code sample gets the first worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var firstSheet = context.workbook.worksheets.getFirst();
    firstSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the first worksheet is "${firstSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-last-worksheet"></a><span data-ttu-id="a083a-119">Obter a última planilha</span><span class="sxs-lookup"><span data-stu-id="a083a-119">Get the last worksheet</span></span>

<span data-ttu-id="a083a-120">O exemplo de código a seguir obtém a última planilha na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="a083a-120">The following code sample gets the last worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var lastSheet = context.workbook.worksheets.getLast();
    lastSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the last worksheet is "${lastSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-next-worksheet"></a><span data-ttu-id="a083a-121">Obter a próxima planilha</span><span class="sxs-lookup"><span data-stu-id="a083a-121">Get the next worksheet</span></span>

<span data-ttu-id="a083a-p104">O exemplo de código a seguir obtém a planilha que vem depois da planilha ativa na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console. Se não houver planilha após a planilha ativa, o método **getNext()** gerará um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="a083a-p104">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its **name** property, and writes a message to the console. If there is no worksheet after the active worksheet, the **getNext()** method throws an **ItemNotFound** error.</span></span>

```js
 Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    var nextSheet = currentSheet.getNext();
    nextSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the sheet that follows the active worksheet is "${nextSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="a083a-124">Obter a planilha anterior</span><span class="sxs-lookup"><span data-stu-id="a083a-124">Get the previous worksheet</span></span>

<span data-ttu-id="a083a-p105">O exemplo de código a seguir obtém a planilha que precede a planilha ativa na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console. Se não houver planilha antes da planilha ativa, o método **getPrevious()** gerará um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="a083a-p105">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its **name** property, and writes a message to the console. If there is no worksheet before the active worksheet, the **getPrevious()** method throws an **ItemNotFound** error.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    var previousSheet = currentSheet.getPrevious();
    previousSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the sheet that precedes the active worksheet is "${previousSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="add-a-worksheet"></a><span data-ttu-id="a083a-127">Adicionar uma planilha</span><span class="sxs-lookup"><span data-stu-id="a083a-127">Add a worksheet</span></span>

<span data-ttu-id="a083a-p106">O exemplo de código a seguir adiciona uma nova planilha chamada **Amostra** à pasta de trabalho, carrega suas propriedades **name** e **position** e grava uma mensagem no console. A nova planilha é adicionada após todas as planilhas existentes.</span><span class="sxs-lookup"><span data-stu-id="a083a-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its **name** and **position** properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;

    var sheet = sheets.add("Sample");
    sheet.load("name, position");

    return context.sync()
        .then(function () {
            console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
        });
}).catch(errorHandlerFunction);
```

## <a name="delete-a-worksheet"></a><span data-ttu-id="a083a-130">Excluir uma planilha</span><span class="sxs-lookup"><span data-stu-id="a083a-130">Delete a worksheet</span></span>

<span data-ttu-id="a083a-131">O exemplo de código a seguir exclui a planilha final na pasta de trabalho (desde que ela não seja a única folha na pasta de trabalho) e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="a083a-131">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
        .then(function () {
            if (sheets.items.length === 1) {
                console.log("Unable to delete the only worksheet in the workbook");
            } else {
                var lastSheet = sheets.items[sheets.items.length - 1];

                console.log(`Deleting worksheet named "${lastSheet.name}"`);
                lastSheet.delete();

                return context.sync();
            };
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="a083a-132">Uma planilha com visibilidade de "[Muito oculta](/javascript/api/excel/excel.sheetvisibility)" não pode ser excluída com o método `delete`.</span><span class="sxs-lookup"><span data-stu-id="a083a-132">A worksheet with a visibility of "[Very Hidden](/javascript/api/excel/excel.sheetvisibility)" cannot be deleted with the `delete` method.</span></span> <span data-ttu-id="a083a-133">Se você quiser excluir a planilha de qualquer forma, deverá primeiro alterar a visibilidade.</span><span class="sxs-lookup"><span data-stu-id="a083a-133">If you wish to delete the worksheet anyway, you must first change the visibility.</span></span>

## <a name="rename-a-worksheet"></a><span data-ttu-id="a083a-134">Renomear uma planilha</span><span class="sxs-lookup"><span data-stu-id="a083a-134">Rename a worksheet</span></span>

<span data-ttu-id="a083a-135">O exemplo de código a seguir altera o nome da planilha ativa para **Novo Nome**.</span><span class="sxs-lookup"><span data-stu-id="a083a-135">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="a083a-136">Mover uma planilha</span><span class="sxs-lookup"><span data-stu-id="a083a-136">Move a worksheet</span></span>

<span data-ttu-id="a083a-137">O exemplo de código a seguir move uma planilha da última posição para a primeira posição na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="a083a-137">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items");

    return context.sync()
        .then(function () {
            var lastSheet = sheets.items[sheets.items.length - 1];
            lastSheet.position = 0;

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

## <a name="set-worksheet-visibility"></a><span data-ttu-id="a083a-138">Definir visibilidade da planilha</span><span class="sxs-lookup"><span data-stu-id="a083a-138">Set worksheet visibility</span></span>

<span data-ttu-id="a083a-139">Esses exemplos mostram como definir a visibilidade de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="a083a-139">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="a083a-140">Ocultar uma planilha</span><span class="sxs-lookup"><span data-stu-id="a083a-140">Hide a worksheet</span></span>

<span data-ttu-id="a083a-141">O exemplo de código a seguir define a visibilidade da planilha chamada **Amostra** para oculta, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="a083a-141">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.hidden;
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`Worksheet with name "${sheet.name}" is hidden`);
        });
}).catch(errorHandlerFunction);
```

### <a name="unhide-a-worksheet"></a><span data-ttu-id="a083a-142">Reexibir uma planilha</span><span class="sxs-lookup"><span data-stu-id="a083a-142">Unhide a worksheet</span></span>

<span data-ttu-id="a083a-143">O exemplo de código a seguir define a visibilidade da planilha chamada **Amostra** para visível, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="a083a-143">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.visible;
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`Worksheet with name "${sheet.name}" is visible`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-a-single-cell-within-a-worksheet"></a><span data-ttu-id="a083a-144">Obter uma única célula em uma planilha</span><span class="sxs-lookup"><span data-stu-id="a083a-144">Get a single cell within a worksheet</span></span>

<span data-ttu-id="a083a-145">O exemplo de código a seguir obtém a célula que está localizada na linha 2, coluna 5 da planilha chamada **Amostra**, carrega suas propriedades **address** e **values** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="a083a-145">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its **address** and **values** properties, and writes a message to the console.</span></span> <span data-ttu-id="a083a-146">Os valores que são passados no método `getCell(row: number, column:number)` são número de linha e número de coluna indexados por zero para a célula que está sendo recuperada.</span><span class="sxs-lookup"><span data-stu-id="a083a-146">The values that are passed into the `getCell(row: number, column:number)` method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var cell = sheet.getCell(1, 4);
    cell.load("address, values");

    return context.sync()
        .then(function() {
            console.log(`The value of the cell in row 2, column 5 is "${cell.values[0][0]}" and the address of that cell is "${cell.address}"`);
        })
}).catch(errorHandlerFunction);
```

## <a name="detect-data-changes"></a><span data-ttu-id="a083a-147">Detectar as alterações dos dados</span><span class="sxs-lookup"><span data-stu-id="a083a-147">Detect data changes</span></span>

<span data-ttu-id="a083a-148">O suplemento precisará reagir aos usuários alterando os dados em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="a083a-148">Your add-in may need to react to users changing the data in a worksheet.</span></span> <span data-ttu-id="a083a-149">Para detectar essas alterações, basta [Registrar um manipulador de eventos.](excel-add-ins-events.md#register-an-event-handler) para o `onChanged` evento da planilha.</span><span class="sxs-lookup"><span data-stu-id="a083a-149">To detect these changes, you can [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the `onChanged` event of a worksheet.</span></span> <span data-ttu-id="a083a-150">Manipuladores de eventos para o `onChanged` evento recebem um objeto [WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs) quando o evento é acionado.</span><span class="sxs-lookup"><span data-stu-id="a083a-150">Event handlers for the `onChanged` event receive a [WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs) object when the event fires.</span></span>

<span data-ttu-id="a083a-151">O `WorksheetChangedEventArgs` objeto fornece informações sobre as alterações e a fonte.</span><span class="sxs-lookup"><span data-stu-id="a083a-151">The `WorksheetChangedEventArgs` object provides information about the changes and the source.</span></span> <span data-ttu-id="a083a-152">Como `onChanged` o acionamento ocorre quando o formato ou o valor dos dados mudam, pode ser útil checar com o suplemento se os valores realmente foram alterados.</span><span class="sxs-lookup"><span data-stu-id="a083a-152">Since `onChanged` fires when either the format or value of the data changes, it can be useful to have your add-in check if the values have actually changed.</span></span> <span data-ttu-id="a083a-153">A `details` propriedade encapsula estas informações como um [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span><span class="sxs-lookup"><span data-stu-id="a083a-153">The `details` property encapsulates this information as a [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span></span> <span data-ttu-id="a083a-154">O exemplo a seguir mostra como exibir o antes e depois dos valores e tipos de uma célula que foi alterada.</span><span class="sxs-lookup"><span data-stu-id="a083a-154">The following code sample shows how to display the before and after values and types of a cell that has been changed.</span></span>

```js
// This function would be used as an event handler for the Worksheet.onChanged event.
function onWorksheetChanged(eventArgs) {
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

## <a name="handle-sorting-events-preview"></a><span data-ttu-id="a083a-155">Manipulação de eventos de classificação (visualização)</span><span class="sxs-lookup"><span data-stu-id="a083a-155">Handle sorting events (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="a083a-156">As APIs para esses eventos relacionados à classificação estão disponíveis atualmente apenas na visualização pública.</span><span class="sxs-lookup"><span data-stu-id="a083a-156">The APIs for these sort-related events are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="a083a-157">Os eventos `onColumnSorted` e `onRowSorted` indicam quando quaisquer dados de planilha são classificados.</span><span class="sxs-lookup"><span data-stu-id="a083a-157">The `onColumnSorted` and `onRowSorted` events indicate when any worksheet data is sorted.</span></span> <span data-ttu-id="a083a-158">Esses eventos estão conectados a objetos `Worksheet` individuais e à `WorkbookCollection` da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="a083a-158">These events are connected to individual `Worksheet` objects and to the workbook's `WorkbookCollection`.</span></span> <span data-ttu-id="a083a-159">Eles são acionados independentemente da classificação ser realizada de forma programática ou manualmente por meio da interface de usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="a083a-159">They fire whether the sorting is done programmatically or manually through the Excel user interface.</span></span>

> [!NOTE]
> <span data-ttu-id="a083a-160">`onColumnSorted` aciona quando as colunas são classificadas como resultado de uma operação de classificação da esquerda para a direita.</span><span class="sxs-lookup"><span data-stu-id="a083a-160">`onColumnSorted` fires when columns are sorted as the result of a left-to-right sort operation.</span></span> <span data-ttu-id="a083a-161">`onRowSorted` aciona quando as linhas são classificadas como resultado de uma operação de classificação de cima para baixo.</span><span class="sxs-lookup"><span data-stu-id="a083a-161">`onRowSorted` fires when rows are sorted as the result of a top-to-bottom sort operation.</span></span> <span data-ttu-id="a083a-162">Classificar uma tabela usando o menu suspenso em um cabeçalho da coluna resulta em um evento `onRowSorted`.</span><span class="sxs-lookup"><span data-stu-id="a083a-162">Sorting a table using the drop-down menu on a column header results in an `onRowSorted` event.</span></span> <span data-ttu-id="a083a-163">O evento corresponde ao que está movendo, não ao que está sendo considerado como os critérios de classificação.</span><span class="sxs-lookup"><span data-stu-id="a083a-163">The event corresponds with what is moving, not what is being considered as the sorting criteria.</span></span>

<span data-ttu-id="a083a-164">Os eventos `onColumnSorted` e `onRowSorted` fornecem seus retornos de chamadas com [WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs) ou [WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs), respectivamente.</span><span class="sxs-lookup"><span data-stu-id="a083a-164">The `onColumnSorted` and `onRowSorted` events provide their callbacks with [WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs) or [WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs), respectively.</span></span> <span data-ttu-id="a083a-165">Isso fornece mais detalhes sobre o evento.</span><span class="sxs-lookup"><span data-stu-id="a083a-165">These give more details about the event.</span></span> <span data-ttu-id="a083a-166">Em particular, ambos `EventArgs` têm uma propriedade `address` que representa as linhas ou as colunas movidas como resultado da operação de classificação.</span><span class="sxs-lookup"><span data-stu-id="a083a-166">In particular, both `EventArgs` have an `address` property that represents the rows or columns moved as a result of the sort operation.</span></span> <span data-ttu-id="a083a-167">Qualquer célula com o conteúdo classificado será incluída, mesmo que o valor da célula não seja parte do critério de classificação.</span><span class="sxs-lookup"><span data-stu-id="a083a-167">Any cell with sorted content is included, even if that cell's value was not part of the sorting criteria.</span></span>

<span data-ttu-id="a083a-168">As imagens a seguir mostram os intervalos retornados pela propriedade `address` para eventos de classificação.</span><span class="sxs-lookup"><span data-stu-id="a083a-168">The following images show the ranges returned by the `address` property for sort events.</span></span> <span data-ttu-id="a083a-169">Primeiro, aqui estão os dados de exemplo antes da classificação:</span><span class="sxs-lookup"><span data-stu-id="a083a-169">First, here is the sample data before sorting:</span></span>

![Dados da tabela no Excel antes de serem classificados](../images/excel-sort-event-before.png)

<span data-ttu-id="a083a-171">Se uma classificação de cima para baixo for realizada no "**Q1**" (os valores em "**B**"), as seguintes linhas realçadas serão retornadas por `WorksheetRowSortedEventArgs.address`:</span><span class="sxs-lookup"><span data-stu-id="a083a-171">If a top-to-bottom sort is performed on "**Q1**" (the values in "**B**"), the following highlighted rows are returned by `WorksheetRowSortedEventArgs.address`:</span></span>

![Dados da tabela no Excel após uma classificação de cima para baixo.](../images/excel-sort-event-after-row.png)

<span data-ttu-id="a083a-174">Se uma classificação da esquerda para a direita for executada em "**Quinces**" (os valores em "**4**") nos dados originais, as seguintes colunas realçadas serão retornadas por `WorksheetColumnsSortedEventArgs.address`:</span><span class="sxs-lookup"><span data-stu-id="a083a-174">If a left-to-right sort is performed on "**Quinces**" (the values in "**4**") on the original data, the following highlighted columns are returned by `WorksheetColumnsSortedEventArgs.address`:</span></span>

![Dados da tabela no Excel após uma classificação da esquerda para a direita.](../images/excel-sort-event-after-column.png)

<span data-ttu-id="a083a-177">O exemplo de código a seguir mostra como registrar um manipulador de eventos para o evento `Worksheet.onRowSorted`.</span><span class="sxs-lookup"><span data-stu-id="a083a-177">The following code sample shows how to register an event handler for the `Worksheet.onRowSorted` event.</span></span> <span data-ttu-id="a083a-178">O retorno de chamada do manipulador limpa a cor de preenchimento do intervalo, e depois preenche as células das linhas movidas.</span><span class="sxs-lookup"><span data-stu-id="a083a-178">The handler's callback clears the fill color for the range, then fills the cells of the moved rows.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // This will fire whenever a row has been moved as the result of a sort action.
    sheet.onRowSorted.add(function (event) {
        return Excel.run(function (context) {
            console.log("Row sorted: " + event.address);
            var sheet = context.workbook.worksheets.getActiveWorksheet();

            // Clear formatting for section, then highlight the sorted area.
            sheet.getRange("A1:E5").format.fill.clear();
            if (event.address !== "") {
                sheet.getRanges(event.address).format.fill.color = "yellow";
            }

            return context.sync();
        });
    });

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="find-all-cells-with-matching-text"></a><span data-ttu-id="a083a-179">Localizar todas as células com texto correspondente</span><span class="sxs-lookup"><span data-stu-id="a083a-179">Find all cells with matching text (preview)</span></span>

<span data-ttu-id="a083a-180">O objeto `Worksheet` tem o método `find` para pesquisar uma cadeia especificada dentro da planilha.</span><span class="sxs-lookup"><span data-stu-id="a083a-180">The `Worksheet` object has a `find` method to search for a specified string within the worksheet.</span></span> <span data-ttu-id="a083a-181">Ele retorna um objeto `RangeAreas`, que é um conjunto de objetos `Range` que podem ser editados ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="a083a-181">It returns a `RangeAreas` object, which is a collection of `Range` objects that can be edited all at once.</span></span> <span data-ttu-id="a083a-182">O exemplo de código a seguir localiza todas as células com valores iguais à cadeia de caracteres **Concluída** e os marca de verde.</span><span class="sxs-lookup"><span data-stu-id="a083a-182">The following code sample finds all cells with values equal to the string **Complete** and colors them green.</span></span> <span data-ttu-id="a083a-183">Observe que `findAll` exibirá um erro `ItemNotFound` se a cadeia especificada não existir na planilha.</span><span class="sxs-lookup"><span data-stu-id="a083a-183">Note that `findAll` will throw an `ItemNotFound` error if the specified string doesn't exist in the worksheet.</span></span> <span data-ttu-id="a083a-184">Se você acha que a cadeia especificada pode não estar na planilha, use o método [findAllOrNullObject](excel-add-ins-advanced-concepts.md#ornullobject-methods) para que seu código manipule normalmente esse cenário.</span><span class="sxs-lookup"><span data-stu-id="a083a-184">If you expect that the specified string may not exist in the worksheet, use the [findAllOrNullObject](excel-add-ins-advanced-concepts.md#ornullobject-methods) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var foundRanges = sheet.findAll("Complete", {
        completeMatch: true, // findAll will match the whole cell value
        matchCase: false // findAll will not match case
    });

    return context.sync()
        .then(function() {
            foundRanges.format.fill.color = "green"
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="a083a-185">Esta seção descreve como localizar as células e intervalos usando as funções do objeto `Worksheet`.</span><span class="sxs-lookup"><span data-stu-id="a083a-185">This section describes how to find cells and ranges using the `Worksheet` object's functions.</span></span> <span data-ttu-id="a083a-186">Encontre mais informações de recuperação de intervalo nos artigos específicos do objeto.</span><span class="sxs-lookup"><span data-stu-id="a083a-186">More range retrieval information can be found in object-specific articles.</span></span>
> - <span data-ttu-id="a083a-187">Confira os exemplos que mostram como obter um intervalo em uma planilha usando o objeto `Range` em [Trabalhar com intervalos usando a API JavaScript do Excel](excel-add-ins-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="a083a-187">For examples that show how to get a range within a worksheet using the `Range` object, see [Work with ranges using the Excel JavaScript API](excel-add-ins-ranges.md).</span></span>
> - <span data-ttu-id="a083a-188">Para obter exemplos que mostram como obter intervalos de um objeto `Table`, confira [Trabalhar com tabelas usando a API JavaScript do Excel](excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="a083a-188">For examples that show how to get ranges from a `Table` object, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).</span></span>
> - <span data-ttu-id="a083a-189">Para obter exemplos que mostram como pesquisar um grande intervalo para vários subgrupos com base nas características da célula, confira [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="a083a-189">For examples that show how to search a large range for multiple sub-ranges based on cell characteristics, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="filter-data"></a><span data-ttu-id="a083a-190">Filtrar dados</span><span class="sxs-lookup"><span data-stu-id="a083a-190">Filter data</span></span>

<span data-ttu-id="a083a-191">Um [AutoFiltro](/javascript/api/excel/excel.autofilter) aplica filtros de data em um intervalo dentro da planilha.</span><span class="sxs-lookup"><span data-stu-id="a083a-191">An [AutoFilter](/javascript/api/excel/excel.autofilter) applies data filters across a range within the worksheet.</span></span> <span data-ttu-id="a083a-192">Isso é criado com `Worksheet.autoFilter.apply`, que possui os seguintes parâmetros:</span><span class="sxs-lookup"><span data-stu-id="a083a-192">This is created with `Worksheet.autoFilter.apply`, which has the following parameters:</span></span>

- <span data-ttu-id="a083a-193">`range`: O intervalo para o qual o filtro é aplicado, especificado como um `Range` objeto ou uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="a083a-193">`range`: The range to which the filter is applied, specified as either a `Range` object or a string.</span></span>
- <span data-ttu-id="a083a-194">`columnIndex`: O índice da coluna com base em zero contra os quais o critério de filtro é avaliado.</span><span class="sxs-lookup"><span data-stu-id="a083a-194">`columnIndex`: The zero-based column index against which the filter criteria is evaluated.</span></span>
- <span data-ttu-id="a083a-195">`criteria`: Um [FilterCriteria](/javascript/api/excel/excel.filtercriteria) objeto determinando quais linhas devem ser filtradas com base na célula da coluna.</span><span class="sxs-lookup"><span data-stu-id="a083a-195">`criteria`: A [FilterCriteria](/javascript/api/excel/excel.filtercriteria) object determining which rows should be filtered based on the column's cell.</span></span>

<span data-ttu-id="a083a-196">O exemplo do primeiro código mostra como adicionar um filtro de intervalo usado na planilha.</span><span class="sxs-lookup"><span data-stu-id="a083a-196">The first code sample shows how to add a filter to the worksheet's used range.</span></span> <span data-ttu-id="a083a-197">Esse filtro ocultará as entradas que não estiverem superior a 25%, com base nos valores na coluna **3**.</span><span class="sxs-lookup"><span data-stu-id="a083a-197">This filter will hide entries that are not in the top 25%, based on the values in column **3**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var farmData = sheet.getUsedRange();

    // This filter will only show the rows with the top 25% of values in column 3.
    sheet.autoFilter.apply(farmData, 3, { criterion1: "25", filterOn: Excel.FilterOn.topPercent });
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a083a-198">O exemplo do código seguinte mostra como atualizar o filtro automático usando o método `reapply`.</span><span class="sxs-lookup"><span data-stu-id="a083a-198">The next code sample shows how to refresh the auto-filter using the `reapply` method.</span></span> <span data-ttu-id="a083a-199">Isso deve ser feito quando os dados no intervalo forem alterados.</span><span class="sxs-lookup"><span data-stu-id="a083a-199">This should be done when the data in the range changes.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.reapply();
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a083a-200">O exemplo de código final de filtro automático mostra como remover o filtro automático de planilha com o método `remove`.</span><span class="sxs-lookup"><span data-stu-id="a083a-200">The final auto-filter code sample shows how to remove the auto-filter from the worksheet with the `remove` method.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.remove();
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a083a-201">Um `AutoFilter` também pode ser aplicado em tabelas individuais.</span><span class="sxs-lookup"><span data-stu-id="a083a-201">An `AutoFilter` can also be applied to individual tables.</span></span> <span data-ttu-id="a083a-202">Consulte [Trabalhar com tabelas usando o API JavaScript do Excel](excel-add-ins-tables.md#autofilter) para mais informações.</span><span class="sxs-lookup"><span data-stu-id="a083a-202">See [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md#autofilter) for more information.</span></span>

## <a name="data-protection"></a><span data-ttu-id="a083a-203">Proteção de dados</span><span class="sxs-lookup"><span data-stu-id="a083a-203">Data protection</span></span>

<span data-ttu-id="a083a-204">O suplemento pode controlar a capacidade de um usuário de editar dados em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="a083a-204">Your add-in can control a user's ability to edit data in a worksheet.</span></span> <span data-ttu-id="a083a-205">A propriedade `protection` da planilha é um objeto [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) com um método `protect()`.</span><span class="sxs-lookup"><span data-stu-id="a083a-205">The worksheet's `protection` property is a [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) object with a `protect()` method.</span></span> <span data-ttu-id="a083a-206">O exemplo a seguir mostra um cenário básico ativando/desativando a proteção completa da planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="a083a-206">The following example shows a basic scenario toggling the complete protection of the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("protection/protected");

    return context.sync().then(function() {
        if (!activeSheet.protection.protected) {
            activeSheet.protection.protect();
        }
    })
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a083a-207">O método `protect` tem dois parâmetros opcionais:</span><span class="sxs-lookup"><span data-stu-id="a083a-207">The `protect` method has two optional parameters:</span></span>

- <span data-ttu-id="a083a-208">`options`: Um objeto [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) definindo restrições de edição de específicas.</span><span class="sxs-lookup"><span data-stu-id="a083a-208">`options`: A [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) object defining specific editing restrictions.</span></span>
- <span data-ttu-id="a083a-209">`password`: Uma cadeia de caracteres que representa a senha necessária para um usuário ignorar a proteção e editar a planilha.</span><span class="sxs-lookup"><span data-stu-id="a083a-209">`password`: A string representing the password needed for a user to bypass protection and edit the worksheet.</span></span>

<span data-ttu-id="a083a-210">O artigo [Proteger uma planilha](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) tem mais informações sobre a proteção de planilhas e sobre como alterar na interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="a083a-210">The article [Protect a worksheet](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) has more information about worksheet protection and how to change it through the Excel UI.</span></span>

## <a name="page-layout-and-print-settings"></a><span data-ttu-id="a083a-211">Configurações de impressão e layout da página</span><span class="sxs-lookup"><span data-stu-id="a083a-211">Page layout and print settings</span></span>

<span data-ttu-id="a083a-212">Os suplementos tem acesso às configurações de layout de página em um nível de planilha.</span><span class="sxs-lookup"><span data-stu-id="a083a-212">Add-ins have access to page layout settings at a worksheet level.</span></span> <span data-ttu-id="a083a-213">Estes controlam como a planilha é impressa.</span><span class="sxs-lookup"><span data-stu-id="a083a-213">These control how the sheet is printed.</span></span> <span data-ttu-id="a083a-214">Um `Worksheet` objeto tem três propriedades de layout relacionadas: `horizontalPageBreaks`, `verticalPageBreaks`, `pageLayout`.</span><span class="sxs-lookup"><span data-stu-id="a083a-214">A `Worksheet` object has three layout-related properties: `horizontalPageBreaks`, `verticalPageBreaks`, `pageLayout`.</span></span>

<span data-ttu-id="a083a-215">`Worksheet.horizontalPageBreaks` e `Worksheet.verticalPageBreaks` são [PageBreakCollections](/javascript/api/excel/excel.pagebreakcollection).</span><span class="sxs-lookup"><span data-stu-id="a083a-215">`Worksheet.horizontalPageBreaks` and `Worksheet.verticalPageBreaks` are [PageBreakCollections](/javascript/api/excel/excel.pagebreakcollection).</span></span> <span data-ttu-id="a083a-216">Estes são conjuntos [Quebras de página](/javascript/api/excel/excel.pagebreak), que especificam os intervalos em que as quebras de página manuais são inseridas.</span><span class="sxs-lookup"><span data-stu-id="a083a-216">These are collections of [PageBreaks](/javascript/api/excel/excel.pagebreak), which specify ranges where manual page breaks are inserted.</span></span> <span data-ttu-id="a083a-217">O exemplo de código a seguir adiciona uma quebra de página horizontal acima da linha **21**.</span><span class="sxs-lookup"><span data-stu-id="a083a-217">The following code sample adds a horizontal page break above row **21**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.horizontalPageBreaks.add("A21:E21"); // The page break is added above this range.
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a083a-218">`Worksheet.pageLayout` é um objeto [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="a083a-218">`Worksheet.pageLayout` is a [PageLayout](/javascript/api/excel/excel.pagelayout) object.</span></span> <span data-ttu-id="a083a-219">Esse objeto contém as configurações de layout e impressão que não são dependentes da implementação de qualquer impressora específica.</span><span class="sxs-lookup"><span data-stu-id="a083a-219">This object contains layout and print settings that are not dependant any printer-specific implementation.</span></span> <span data-ttu-id="a083a-220">Essas configurações incluem margens, orientação, numeração de página, linhas de título e a área de impressão.</span><span class="sxs-lookup"><span data-stu-id="a083a-220">These settings include margins, orientation, page numbering, title rows, and print area.</span></span>

<span data-ttu-id="a083a-221">O exemplo de código a seguir centraliza a página (tanto verticalmente quanto horizontalmente), define uma linha de título que será impressa na parte superior de cada página e define a área impressa para a subseção da planilha.</span><span class="sxs-lookup"><span data-stu-id="a083a-221">The following code sample centers the page (both vertically and horizontally), sets a title row that will be printed at the top of every page, and sets the printed area to a subsection of the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Center the page in both directions.
    sheet.pageLayout.centerHorizontally = true;
    sheet.pageLayout.centerVertically = true;

    // Set the first row as the title row for every page.
    sheet.pageLayout.setPrintTitleRows("$1:$1");

    // Limit the area to be printed to the range "A1:D100".
    sheet.pageLayout.setPrintArea("A1:D100");

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="a083a-222">Confira também</span><span class="sxs-lookup"><span data-stu-id="a083a-222">See also</span></span>

- [<span data-ttu-id="a083a-223">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="a083a-223">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
