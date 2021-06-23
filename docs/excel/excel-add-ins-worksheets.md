---
title: Trabalhe com planilhas usando a API JavaScript do Excel
description: Exemplos de código que mostram como executar tarefas comuns com planilhas usando Excel API JavaScript.
ms.date: 06/03/2021
localization_priority: Normal
ms.openlocfilehash: 9e181ec800eccb938fa152bb28772b11961c7a40
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075548"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="04f50-103">Trabalhe com planilhas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="04f50-103">Work with worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="04f50-p101">Este artigo fornece exemplos de código que mostram como executar tarefas comuns com planilhas usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos aos quais os objetos `Worksheet` e `WorksheetCollection` dão suporte, confira [Objeto Worksheet (API JavaScript para Excel)](/javascript/api/excel/excel.worksheet) e [Objeto WorksheetCollection (API JavaScript para Excel)](/javascript/api/excel/excel.worksheetcollection).</span><span class="sxs-lookup"><span data-stu-id="04f50-p101">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API. For the complete list of properties and methods that the `Worksheet` and `WorksheetCollection` objects support, see [Worksheet Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheet) and [WorksheetCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheetcollection).</span></span>

> [!NOTE]
> <span data-ttu-id="04f50-106">As informações deste artigo se aplicam apenas a planilhas regulares; elas não se aplicam às folhas "gráfico" ou "macro".</span><span class="sxs-lookup"><span data-stu-id="04f50-106">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="04f50-107">Obter planilhas</span><span class="sxs-lookup"><span data-stu-id="04f50-107">Get worksheets</span></span>

<span data-ttu-id="04f50-108">O exemplo de código a seguir obtém a coleção de planilhas, carrega a propriedade `name` de cada planilha e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="04f50-108">The following code sample gets the collection of worksheets, loads the `name` property of each worksheet, and writes a message to the console.</span></span>

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
            sheets.items.forEach(function (sheet) {
              console.log(sheet.name);
            });
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="04f50-p102">A propriedade `id` de uma planilha identifica exclusivamente a planilha em uma determinada pasta de trabalho e seu valor permanecerá igual, mesmo quando a planilha for renomeada ou movida. Quando uma planilha é excluída de uma pasta de trabalho no Excel para Mac, a `id` da planilha excluída pode ser reatribuída a uma nova planilha que é subsequentemente criada.</span><span class="sxs-lookup"><span data-stu-id="04f50-p102">The `id` property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved. When a worksheet is deleted from a workbook in Excel on Mac, the `id` of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="04f50-111">Obter a planilha ativa</span><span class="sxs-lookup"><span data-stu-id="04f50-111">Get the active worksheet</span></span>

<span data-ttu-id="04f50-112">O exemplo de código a seguir obtém a planilha ativa, carrega sua propriedade `name` e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="04f50-112">The following code sample gets the active worksheet, loads its `name` property, and writes a message to the console.</span></span>

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

## <a name="set-the-active-worksheet"></a><span data-ttu-id="04f50-113">Definir a planilha ativa</span><span class="sxs-lookup"><span data-stu-id="04f50-113">Set the active worksheet</span></span>

<span data-ttu-id="04f50-p103">O exemplo de código a seguir define a planilha ativa para a planilha chamada **Amostra**, carrega sua propriedade `name` e grava uma mensagem no console. Se não houver planilha com esse nome, o método `activate()` gerará um erro `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="04f50-p103">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its `name` property, and writes a message to the console. If there is no worksheet with that name, the `activate()` method throws an `ItemNotFound` error.</span></span>

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

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="04f50-116">Planilhas de referência por posição relativa</span><span class="sxs-lookup"><span data-stu-id="04f50-116">Reference worksheets by relative position</span></span>

<span data-ttu-id="04f50-117">Esses exemplos mostram como fazer referência a uma planilha por sua posição relativa.</span><span class="sxs-lookup"><span data-stu-id="04f50-117">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="04f50-118">Obter a primeira planilha</span><span class="sxs-lookup"><span data-stu-id="04f50-118">Get the first worksheet</span></span>

<span data-ttu-id="04f50-119">O exemplo de código a seguir obtém a primeira planilha na pasta de trabalho, carrega sua propriedade `name` e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="04f50-119">The following code sample gets the first worksheet in the workbook, loads its `name` property, and writes a message to the console.</span></span>

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

### <a name="get-the-last-worksheet"></a><span data-ttu-id="04f50-120">Obter a última planilha</span><span class="sxs-lookup"><span data-stu-id="04f50-120">Get the last worksheet</span></span>

<span data-ttu-id="04f50-121">O exemplo de código a seguir obtém a última planilha na pasta de trabalho, carrega sua propriedade `name` e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="04f50-121">The following code sample gets the last worksheet in the workbook, loads its `name` property, and writes a message to the console.</span></span>

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

### <a name="get-the-next-worksheet"></a><span data-ttu-id="04f50-122">Obter a próxima planilha</span><span class="sxs-lookup"><span data-stu-id="04f50-122">Get the next worksheet</span></span>

<span data-ttu-id="04f50-p104">O exemplo de código a seguir obtém a planilha que vem depois da planilha ativa na pasta de trabalho, carrega sua propriedade `name` e grava uma mensagem no console. Se não houver planilha após a planilha ativa, o método `getNext()` gerará um erro `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="04f50-p104">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its `name` property, and writes a message to the console. If there is no worksheet after the active worksheet, the `getNext()` method throws an `ItemNotFound` error.</span></span>

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

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="04f50-125">Obter a planilha anterior</span><span class="sxs-lookup"><span data-stu-id="04f50-125">Get the previous worksheet</span></span>

<span data-ttu-id="04f50-p105">O exemplo de código a seguir obtém a planilha que precede a planilha ativa na pasta de trabalho, carrega sua propriedade `name` e grava uma mensagem no console. Se não houver planilha antes da planilha ativa, o método `getPrevious()` gerará um erro `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="04f50-p105">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its `name` property, and writes a message to the console. If there is no worksheet before the active worksheet, the `getPrevious()` method throws an `ItemNotFound` error.</span></span>

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

## <a name="add-a-worksheet"></a><span data-ttu-id="04f50-128">Adicionar uma planilha</span><span class="sxs-lookup"><span data-stu-id="04f50-128">Add a worksheet</span></span>

<span data-ttu-id="04f50-p106">O exemplo de código a seguir adiciona uma nova planilha chamada **Amostra** à pasta de trabalho, carrega suas propriedades `name` e `position` e grava uma mensagem no console. A nova planilha é adicionada após todas as planilhas existentes.</span><span class="sxs-lookup"><span data-stu-id="04f50-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its `name` and `position` properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

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

### <a name="copy-an-existing-worksheet"></a><span data-ttu-id="04f50-131">Copiar uma planilha existente</span><span class="sxs-lookup"><span data-stu-id="04f50-131">Copy an existing worksheet</span></span>

<span data-ttu-id="04f50-132">`Worksheet.copy` adiciona uma nova planilha que é uma cópia de uma planilha existente.</span><span class="sxs-lookup"><span data-stu-id="04f50-132">`Worksheet.copy` adds a new worksheet that is a copy of an existing worksheet.</span></span> <span data-ttu-id="04f50-133">O nome da nova planilha terá um número anexado ao final, consistente com a cópia de uma planilha feita pela Interface do Usuário do Excel (por exemplo, **MySheet (2)**).</span><span class="sxs-lookup"><span data-stu-id="04f50-133">The new worksheet's name will have a number appended to the end, in a manner consistent with copying a worksheet through the Excel UI (for example, **MySheet (2)**).</span></span> <span data-ttu-id="04f50-134">`Worksheet.copy` pode-se usar dois parâmetros, ambos opcionais:</span><span class="sxs-lookup"><span data-stu-id="04f50-134">`Worksheet.copy` can take two parameters, both of which are optional:</span></span>

- <span data-ttu-id="04f50-135">`positionType` -Um [WorksheetPositionType](/javascript/api/excel/excel.worksheetpositiontype) enum especificando o local da pasta de trabalho em que a nova planilha deve ser adicionada.</span><span class="sxs-lookup"><span data-stu-id="04f50-135">`positionType` - A [WorksheetPositionType](/javascript/api/excel/excel.worksheetpositiontype) enum specifying where in the workbook the new worksheet is to be added.</span></span>
- <span data-ttu-id="04f50-136">`relativeTo` -Se o `positionType` for `Before` ou `After`, você precisa especificar uma planilha relativa à qual a nova planilha deve ser adicionada (esse parâmetro responde a pergunta "antes ou depois?").</span><span class="sxs-lookup"><span data-stu-id="04f50-136">`relativeTo` - If the `positionType` is `Before` or `After`, you need to specify a worksheet relative to which the new sheet is to be added (this parameter answers the question "Before or after what?").</span></span>

<span data-ttu-id="04f50-137">O exemplo de código a seguir copia a planilha atual e insere a nova planilha logo após a planilha atual.</span><span class="sxs-lookup"><span data-stu-id="04f50-137">The following code sample copies the current worksheet and inserts the new sheet directly after the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var myWorkbook = context.workbook;
    var sampleSheet = myWorkbook.worksheets.getActiveWorksheet();
    var copiedSheet = sampleSheet.copy(Excel.WorksheetPositionType.after, sampleSheet);
    return context.sync();
});
```

## <a name="delete-a-worksheet"></a><span data-ttu-id="04f50-138">Excluir uma planilha</span><span class="sxs-lookup"><span data-stu-id="04f50-138">Delete a worksheet</span></span>

<span data-ttu-id="04f50-139">O exemplo de código a seguir exclui a planilha final na pasta de trabalho (desde que ela não seja a única folha na pasta de trabalho) e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="04f50-139">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

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
> <span data-ttu-id="04f50-140">Uma planilha com visibilidade de "[Muito oculta](/javascript/api/excel/excel.sheetvisibility)" não pode ser excluída com o método `delete`.</span><span class="sxs-lookup"><span data-stu-id="04f50-140">A worksheet with a visibility of "[Very Hidden](/javascript/api/excel/excel.sheetvisibility)" cannot be deleted with the `delete` method.</span></span> <span data-ttu-id="04f50-141">Se você quiser excluir a planilha de qualquer forma, deverá primeiro alterar a visibilidade.</span><span class="sxs-lookup"><span data-stu-id="04f50-141">If you wish to delete the worksheet anyway, you must first change the visibility.</span></span>

## <a name="rename-a-worksheet"></a><span data-ttu-id="04f50-142">Renomear uma planilha</span><span class="sxs-lookup"><span data-stu-id="04f50-142">Rename a worksheet</span></span>

<span data-ttu-id="04f50-143">O exemplo de código a seguir altera o nome da planilha ativa para **Novo Nome**.</span><span class="sxs-lookup"><span data-stu-id="04f50-143">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="04f50-144">Mover uma planilha</span><span class="sxs-lookup"><span data-stu-id="04f50-144">Move a worksheet</span></span>

<span data-ttu-id="04f50-145">O exemplo de código a seguir move uma planilha da última posição para a primeira posição na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="04f50-145">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

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

## <a name="set-worksheet-visibility"></a><span data-ttu-id="04f50-146">Definir visibilidade da planilha</span><span class="sxs-lookup"><span data-stu-id="04f50-146">Set worksheet visibility</span></span>

<span data-ttu-id="04f50-147">Esses exemplos mostram como definir a visibilidade de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-147">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="04f50-148">Ocultar uma planilha</span><span class="sxs-lookup"><span data-stu-id="04f50-148">Hide a worksheet</span></span>

<span data-ttu-id="04f50-149">O exemplo de código a seguir define a visibilidade da planilha chamada **Amostra** para oculta, carrega sua propriedade `name` e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="04f50-149">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its `name` property, and writes a message to the console.</span></span>

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

### <a name="unhide-a-worksheet"></a><span data-ttu-id="04f50-150">Reexibir uma planilha</span><span class="sxs-lookup"><span data-stu-id="04f50-150">Unhide a worksheet</span></span>

<span data-ttu-id="04f50-151">O exemplo de código a seguir define a visibilidade da planilha chamada **Amostra** para visível, carrega sua propriedade `name` e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="04f50-151">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its `name` property, and writes a message to the console.</span></span>

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

## <a name="get-a-single-cell-within-a-worksheet"></a><span data-ttu-id="04f50-152">Obter uma única célula em uma planilha</span><span class="sxs-lookup"><span data-stu-id="04f50-152">Get a single cell within a worksheet</span></span>

<span data-ttu-id="04f50-153">O exemplo de código a seguir obtém a célula que está localizada na linha 2, coluna 5 da planilha chamada **Amostra**, carrega suas propriedades `address` e `values` e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="04f50-153">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its `address` and `values` properties, and writes a message to the console.</span></span> <span data-ttu-id="04f50-154">Os valores que são passados no método `getCell(row: number, column:number)` são número de linha e número de coluna indexados por zero para a célula que está sendo recuperada.</span><span class="sxs-lookup"><span data-stu-id="04f50-154">The values that are passed into the `getCell(row: number, column:number)` method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

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

## <a name="detect-data-changes"></a><span data-ttu-id="04f50-155">Detectar as alterações dos dados</span><span class="sxs-lookup"><span data-stu-id="04f50-155">Detect data changes</span></span>

<span data-ttu-id="04f50-156">O suplemento precisará reagir aos usuários alterando os dados em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-156">Your add-in may need to react to users changing the data in a worksheet.</span></span> <span data-ttu-id="04f50-157">Para detectar essas alterações, basta [Registrar um manipulador de eventos.](excel-add-ins-events.md#register-an-event-handler) para o `onChanged` evento da planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-157">To detect these changes, you can [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the `onChanged` event of a worksheet.</span></span> <span data-ttu-id="04f50-158">Manipuladores de eventos para o `onChanged` evento recebem um objeto [WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs) quando o evento é acionado.</span><span class="sxs-lookup"><span data-stu-id="04f50-158">Event handlers for the `onChanged` event receive a [WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs) object when the event fires.</span></span>

<span data-ttu-id="04f50-159">O `WorksheetChangedEventArgs` objeto fornece informações sobre as alterações e a fonte.</span><span class="sxs-lookup"><span data-stu-id="04f50-159">The `WorksheetChangedEventArgs` object provides information about the changes and the source.</span></span> <span data-ttu-id="04f50-160">Como `onChanged` o acionamento ocorre quando o formato ou o valor dos dados mudam, pode ser útil checar com o suplemento se os valores realmente foram alterados.</span><span class="sxs-lookup"><span data-stu-id="04f50-160">Since `onChanged` fires when either the format or value of the data changes, it can be useful to have your add-in check if the values have actually changed.</span></span> <span data-ttu-id="04f50-161">A `details` propriedade encapsula estas informações como um [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span><span class="sxs-lookup"><span data-stu-id="04f50-161">The `details` property encapsulates this information as a [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span></span> <span data-ttu-id="04f50-162">O exemplo a seguir mostra como exibir o antes e depois dos valores e tipos de uma célula que foi alterada.</span><span class="sxs-lookup"><span data-stu-id="04f50-162">The following code sample shows how to display the before and after values and types of a cell that has been changed.</span></span>

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

## <a name="detect-formula-changes-preview"></a><span data-ttu-id="04f50-163">Detectar alterações de fórmula (visualização)</span><span class="sxs-lookup"><span data-stu-id="04f50-163">Detect formula changes (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="04f50-164">No `Worksheet.onFormulaChanged` momento, o evento só está disponível na visualização pública.</span><span class="sxs-lookup"><span data-stu-id="04f50-164">The `Worksheet.onFormulaChanged` event is currently only available in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

<span data-ttu-id="04f50-165">Seu complemento pode acompanhar as alterações nas fórmulas em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-165">Your add-in can track changes to formulas in a worksheet.</span></span> <span data-ttu-id="04f50-166">Isso é útil quando uma planilha está conectada a um banco de dados externo.</span><span class="sxs-lookup"><span data-stu-id="04f50-166">This is useful when a worksheet is connected to an external database.</span></span> <span data-ttu-id="04f50-167">Quando a fórmula é mudada na planilha, o evento nesse cenário dispara atualizações correspondentes no banco de dados externo.</span><span class="sxs-lookup"><span data-stu-id="04f50-167">When the formula changes in the worksheet, the event in this scenario triggers corresponding updates in the external database.</span></span>

<span data-ttu-id="04f50-168">Para detectar alterações nas fórmulas, [registre](excel-add-ins-events.md#register-an-event-handler) um manipulador de eventos para o [evento onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged) de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-168">To detect changes to formulas, [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the [onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged) event of a worksheet.</span></span> <span data-ttu-id="04f50-169">Os manipuladores de eventos `onFormulaChanged` do evento recebem um objeto [WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs) quando o evento é ativos.</span><span class="sxs-lookup"><span data-stu-id="04f50-169">Event handlers for the `onFormulaChanged` event receive a [WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs) object when the event fires.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="04f50-170">O evento detecta quando uma fórmula em si muda, não o valor de dados `onFormulaChanged` resultante do cálculo da fórmula.</span><span class="sxs-lookup"><span data-stu-id="04f50-170">The `onFormulaChanged` event detects when a formula itself changes, not the data value resulting from the formula's calculation.</span></span>

<span data-ttu-id="04f50-171">O exemplo de código a seguir mostra como registrar o manipulador de eventos, usar o objeto para recuperar a matriz formulaDetails da fórmula alterada e imprimir detalhes sobre a fórmula alterada com as propriedades `onFormulaChanged` `WorksheetFormulaChangedEventArgs` [FormulaChangedEventDetail.](/javascript/api/excel/excel.formulachangedeventdetail) [](/javascript/api/excel/excel.worksheetformulachangedeventargs#formulaDetails)</span><span class="sxs-lookup"><span data-stu-id="04f50-171">The following code sample shows how to register the `onFormulaChanged` event handler, use the `WorksheetFormulaChangedEventArgs` object to retrieve the [formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formulaDetails) array of the changed formula, and then print out details about the changed formula with the [FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail) properties.</span></span>

> [!NOTE]
> <span data-ttu-id="04f50-172">Esse exemplo de código só funciona quando uma única fórmula é alterada.</span><span class="sxs-lookup"><span data-stu-id="04f50-172">This code sample only works when a single formula is changed.</span></span>

```js
Excel.run(function (context) {
    // Retrieve the worksheet named "Sample".
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Register the formula changed event handler for this worksheet.
    sheet.onFormulaChanged.add(formulaChangeHandler);

    return context.sync();
});

function formulaChangeHandler(event) {
    Excel.run(function (context) {
        // Retrieve details about the formula change event.
        // Note: This method assumes only a single formula is changed at a time. 
        var cellAddress = event.formulaDetails[0].cellAddress;
        var previousFormula = event.formulaDetails[0].previousFormula;
        var source = event.source;
    
        // Print out the change event details.
        console.log(
          `The formula in cell ${cellAddress} changed. 
          The previous formula was: ${previousFormula}. 
          The source of the change was: ${source}.`
        );         
    });
}
```

## <a name="handle-sorting-events"></a><span data-ttu-id="04f50-173">Manipulação de eventos de classificação</span><span class="sxs-lookup"><span data-stu-id="04f50-173">Handle sorting events</span></span>

<span data-ttu-id="04f50-174">Os eventos `onColumnSorted` e `onRowSorted` indicam quando quaisquer dados de planilha são classificados.</span><span class="sxs-lookup"><span data-stu-id="04f50-174">The `onColumnSorted` and `onRowSorted` events indicate when any worksheet data is sorted.</span></span> <span data-ttu-id="04f50-175">Esses eventos estão conectados a objetos `Worksheet` individuais e à `WorkbookCollection` da pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="04f50-175">These events are connected to individual `Worksheet` objects and to the workbook's `WorkbookCollection`.</span></span> <span data-ttu-id="04f50-176">Eles são acionados independentemente da classificação ser realizada de forma programática ou manualmente por meio da interface de usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="04f50-176">They fire whether the sorting is done programmatically or manually through the Excel user interface.</span></span>

> [!NOTE]
> <span data-ttu-id="04f50-177">`onColumnSorted` aciona quando as colunas são classificadas como resultado de uma operação de classificação da esquerda para a direita.</span><span class="sxs-lookup"><span data-stu-id="04f50-177">`onColumnSorted` fires when columns are sorted as the result of a left-to-right sort operation.</span></span> <span data-ttu-id="04f50-178">`onRowSorted` aciona quando as linhas são classificadas como resultado de uma operação de classificação de cima para baixo.</span><span class="sxs-lookup"><span data-stu-id="04f50-178">`onRowSorted` fires when rows are sorted as the result of a top-to-bottom sort operation.</span></span> <span data-ttu-id="04f50-179">Classificar uma tabela usando o menu suspenso em um cabeçalho da coluna resulta em um evento `onRowSorted`.</span><span class="sxs-lookup"><span data-stu-id="04f50-179">Sorting a table using the drop-down menu on a column header results in an `onRowSorted` event.</span></span> <span data-ttu-id="04f50-180">O evento corresponde ao que está movendo, não ao que está sendo considerado como os critérios de classificação.</span><span class="sxs-lookup"><span data-stu-id="04f50-180">The event corresponds with what is moving, not what is being considered as the sorting criteria.</span></span>

<span data-ttu-id="04f50-181">Os eventos `onColumnSorted` e `onRowSorted` fornecem seus retornos de chamadas com [WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs) ou [WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs), respectivamente.</span><span class="sxs-lookup"><span data-stu-id="04f50-181">The `onColumnSorted` and `onRowSorted` events provide their callbacks with [WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs) or [WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs), respectively.</span></span> <span data-ttu-id="04f50-182">Isso fornece mais detalhes sobre o evento.</span><span class="sxs-lookup"><span data-stu-id="04f50-182">These give more details about the event.</span></span> <span data-ttu-id="04f50-183">Em particular, ambos `EventArgs` têm uma propriedade `address` que representa as linhas ou as colunas movidas como resultado da operação de classificação.</span><span class="sxs-lookup"><span data-stu-id="04f50-183">In particular, both `EventArgs` have an `address` property that represents the rows or columns moved as a result of the sort operation.</span></span> <span data-ttu-id="04f50-184">Qualquer célula com o conteúdo classificado será incluída, mesmo que o valor da célula não seja parte do critério de classificação.</span><span class="sxs-lookup"><span data-stu-id="04f50-184">Any cell with sorted content is included, even if that cell's value was not part of the sorting criteria.</span></span>

<span data-ttu-id="04f50-185">As imagens a seguir mostram os intervalos retornados pela propriedade `address` para eventos de classificação.</span><span class="sxs-lookup"><span data-stu-id="04f50-185">The following images show the ranges returned by the `address` property for sort events.</span></span> <span data-ttu-id="04f50-186">Primeiro, aqui estão os dados de exemplo antes da classificação:</span><span class="sxs-lookup"><span data-stu-id="04f50-186">First, here is the sample data before sorting:</span></span>

![Dados de tabela em Excel antes de serem classificação.](../images/excel-sort-event-before.png)

<span data-ttu-id="04f50-188&quot;>Se uma classificação de cima para baixo for realizada no &quot;**Q1**&quot; (os valores em &quot;**B**"), as seguintes linhas realçadas serão retornadas por `WorksheetRowSortedEventArgs.address`:</span><span class="sxs-lookup"><span data-stu-id="04f50-188">If a top-to-bottom sort is performed on "**Q1**" (the values in "**B**"), the following highlighted rows are returned by `WorksheetRowSortedEventArgs.address`:</span></span>

![Dados da tabela no Excel após uma classificação de cima para baixo.](../images/excel-sort-event-after-row.png)

<span data-ttu-id="04f50-191&quot;>Se uma classificação da esquerda para a direita for executada em &quot;**Quinces**&quot; (os valores em &quot;**4**") nos dados originais, as seguintes colunas realçadas serão retornadas por `WorksheetColumnsSortedEventArgs.address`:</span><span class="sxs-lookup"><span data-stu-id="04f50-191">If a left-to-right sort is performed on "**Quinces**" (the values in "**4**") on the original data, the following highlighted columns are returned by `WorksheetColumnsSortedEventArgs.address`:</span></span>

![Dados da tabela no Excel após uma classificação da esquerda para a direita.](../images/excel-sort-event-after-column.png)

<span data-ttu-id="04f50-194">O exemplo de código a seguir mostra como registrar um manipulador de eventos para o evento `Worksheet.onRowSorted`.</span><span class="sxs-lookup"><span data-stu-id="04f50-194">The following code sample shows how to register an event handler for the `Worksheet.onRowSorted` event.</span></span> <span data-ttu-id="04f50-195">O retorno de chamada do manipulador limpa a cor de preenchimento do intervalo, e depois preenche as células das linhas movidas.</span><span class="sxs-lookup"><span data-stu-id="04f50-195">The handler's callback clears the fill color for the range, then fills the cells of the moved rows.</span></span>

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

## <a name="find-all-cells-with-matching-text"></a><span data-ttu-id="04f50-196">Localizar todas as células com texto correspondente</span><span class="sxs-lookup"><span data-stu-id="04f50-196">Find all cells with matching text</span></span>

<span data-ttu-id="04f50-197">O objeto `Worksheet` tem o método `find` para pesquisar uma cadeia especificada dentro da planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-197">The `Worksheet` object has a `find` method to search for a specified string within the worksheet.</span></span> <span data-ttu-id="04f50-198">Ele retorna um objeto `RangeAreas`, que é um conjunto de objetos `Range` que podem ser editados ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="04f50-198">It returns a `RangeAreas` object, which is a collection of `Range` objects that can be edited all at once.</span></span> <span data-ttu-id="04f50-199">O exemplo de código a seguir localiza todas as células com valores iguais à cadeia de caracteres **Concluída** e os marca de verde.</span><span class="sxs-lookup"><span data-stu-id="04f50-199">The following code sample finds all cells with values equal to the string **Complete** and colors them green.</span></span> <span data-ttu-id="04f50-200">Observe que `findAll` exibirá um erro `ItemNotFound` se a cadeia especificada não existir na planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-200">Note that `findAll` will throw an `ItemNotFound` error if the specified string doesn't exist in the worksheet.</span></span> <span data-ttu-id="04f50-201">Se você acha que a cadeia especificada pode não estar na planilha, use o método [findAllOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) para que seu código manipule normalmente esse cenário.</span><span class="sxs-lookup"><span data-stu-id="04f50-201">If you expect that the specified string may not exist in the worksheet, use the [findAllOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) method instead, so your code gracefully handles that scenario.</span></span>

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
> <span data-ttu-id="04f50-202">Esta seção descreve como localizar as células e intervalos usando as funções do objeto `Worksheet`.</span><span class="sxs-lookup"><span data-stu-id="04f50-202">This section describes how to find cells and ranges using the `Worksheet` object's functions.</span></span> <span data-ttu-id="04f50-203">Encontre mais informações de recuperação de intervalo nos artigos específicos do objeto.</span><span class="sxs-lookup"><span data-stu-id="04f50-203">More range retrieval information can be found in object-specific articles.</span></span>
> - <span data-ttu-id="04f50-204">Para obter exemplos que mostram como obter um intervalo dentro de uma planilha usando o objeto, consulte Obter um intervalo usando o Excel `Range` [API JavaScript](excel-add-ins-ranges-get.md).</span><span class="sxs-lookup"><span data-stu-id="04f50-204">For examples that show how to get a range within a worksheet using the `Range` object, see [Get a range using the Excel JavaScript API](excel-add-ins-ranges-get.md).</span></span>
> - <span data-ttu-id="04f50-205">Para obter exemplos que mostram como obter intervalos de um objeto `Table`, confira [Trabalhar com tabelas usando a API JavaScript do Excel](excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="04f50-205">For examples that show how to get ranges from a `Table` object, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).</span></span>
> - <span data-ttu-id="04f50-206">Para obter exemplos que mostram como pesquisar um grande intervalo para vários subgrupos com base nas características da célula, confira [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="04f50-206">For examples that show how to search a large range for multiple sub-ranges based on cell characteristics, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="filter-data"></a><span data-ttu-id="04f50-207">Filtrar dados</span><span class="sxs-lookup"><span data-stu-id="04f50-207">Filter data</span></span>

<span data-ttu-id="04f50-208">Um [AutoFiltro](/javascript/api/excel/excel.autofilter) aplica filtros de data em um intervalo dentro da planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-208">An [AutoFilter](/javascript/api/excel/excel.autofilter) applies data filters across a range within the worksheet.</span></span> <span data-ttu-id="04f50-209">Isso é criado com `Worksheet.autoFilter.apply`, que possui os seguintes parâmetros:</span><span class="sxs-lookup"><span data-stu-id="04f50-209">This is created with `Worksheet.autoFilter.apply`, which has the following parameters:</span></span>

- <span data-ttu-id="04f50-210">`range`: O intervalo para o qual o filtro é aplicado, especificado como um `Range` objeto ou uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="04f50-210">`range`: The range to which the filter is applied, specified as either a `Range` object or a string.</span></span>
- <span data-ttu-id="04f50-211">`columnIndex`: O índice da coluna com base em zero contra os quais o critério de filtro é avaliado.</span><span class="sxs-lookup"><span data-stu-id="04f50-211">`columnIndex`: The zero-based column index against which the filter criteria is evaluated.</span></span>
- <span data-ttu-id="04f50-212">`criteria`: Um [FilterCriteria](/javascript/api/excel/excel.filtercriteria) objeto determinando quais linhas devem ser filtradas com base na célula da coluna.</span><span class="sxs-lookup"><span data-stu-id="04f50-212">`criteria`: A [FilterCriteria](/javascript/api/excel/excel.filtercriteria) object determining which rows should be filtered based on the column's cell.</span></span>

<span data-ttu-id="04f50-213">O exemplo do primeiro código mostra como adicionar um filtro de intervalo usado na planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-213">The first code sample shows how to add a filter to the worksheet's used range.</span></span> <span data-ttu-id="04f50-214">Esse filtro ocultará as entradas que não estiverem superior a 25%, com base nos valores na coluna **3**.</span><span class="sxs-lookup"><span data-stu-id="04f50-214">This filter will hide entries that are not in the top 25%, based on the values in column **3**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var farmData = sheet.getUsedRange();

    // This filter will only show the rows with the top 25% of values in column 3.
    sheet.autoFilter.apply(farmData, 3, { criterion1: "25", filterOn: Excel.FilterOn.topPercent });
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="04f50-215">O exemplo do código seguinte mostra como atualizar o filtro automático usando o método `reapply`.</span><span class="sxs-lookup"><span data-stu-id="04f50-215">The next code sample shows how to refresh the auto-filter using the `reapply` method.</span></span> <span data-ttu-id="04f50-216">Isso deve ser feito quando os dados no intervalo forem alterados.</span><span class="sxs-lookup"><span data-stu-id="04f50-216">This should be done when the data in the range changes.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.reapply();
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="04f50-217">O exemplo de código final de filtro automático mostra como remover o filtro automático de planilha com o método `remove`.</span><span class="sxs-lookup"><span data-stu-id="04f50-217">The final auto-filter code sample shows how to remove the auto-filter from the worksheet with the `remove` method.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.remove();
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="04f50-218">Um `AutoFilter` também pode ser aplicado em tabelas individuais.</span><span class="sxs-lookup"><span data-stu-id="04f50-218">An `AutoFilter` can also be applied to individual tables.</span></span> <span data-ttu-id="04f50-219">Consulte [Trabalhar com tabelas usando o API JavaScript do Excel](excel-add-ins-tables.md#autofilter) para mais informações.</span><span class="sxs-lookup"><span data-stu-id="04f50-219">See [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md#autofilter) for more information.</span></span>

## <a name="data-protection"></a><span data-ttu-id="04f50-220">Proteção de dados</span><span class="sxs-lookup"><span data-stu-id="04f50-220">Data protection</span></span>

<span data-ttu-id="04f50-221">O suplemento pode controlar a capacidade de um usuário de editar dados em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-221">Your add-in can control a user's ability to edit data in a worksheet.</span></span> <span data-ttu-id="04f50-222">A propriedade `protection` da planilha é um objeto [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) com um método `protect()`.</span><span class="sxs-lookup"><span data-stu-id="04f50-222">The worksheet's `protection` property is a [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) object with a `protect()` method.</span></span> <span data-ttu-id="04f50-223">O exemplo a seguir mostra um cenário básico ativando/desativando a proteção completa da planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="04f50-223">The following example shows a basic scenario toggling the complete protection of the active worksheet.</span></span>

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

<span data-ttu-id="04f50-224">O método `protect` tem dois parâmetros opcionais:</span><span class="sxs-lookup"><span data-stu-id="04f50-224">The `protect` method has two optional parameters:</span></span>

- <span data-ttu-id="04f50-225">`options`: Um objeto [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) definindo restrições de edição de específicas.</span><span class="sxs-lookup"><span data-stu-id="04f50-225">`options`: A [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) object defining specific editing restrictions.</span></span>
- <span data-ttu-id="04f50-226">`password`: Uma cadeia de caracteres que representa a senha necessária para um usuário ignorar a proteção e editar a planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-226">`password`: A string representing the password needed for a user to bypass protection and edit the worksheet.</span></span>

<span data-ttu-id="04f50-227">O artigo [Proteger uma planilha](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) tem mais informações sobre a proteção de planilhas e sobre como alterar na interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="04f50-227">The article [Protect a worksheet](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) has more information about worksheet protection and how to change it through the Excel UI.</span></span>

## <a name="page-layout-and-print-settings"></a><span data-ttu-id="04f50-228">Configurações de impressão e layout da página</span><span class="sxs-lookup"><span data-stu-id="04f50-228">Page layout and print settings</span></span>

<span data-ttu-id="04f50-229">Os suplementos tem acesso às configurações de layout de página em um nível de planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-229">Add-ins have access to page layout settings at a worksheet level.</span></span> <span data-ttu-id="04f50-230">Estes controlam como a planilha é impressa.</span><span class="sxs-lookup"><span data-stu-id="04f50-230">These control how the sheet is printed.</span></span> <span data-ttu-id="04f50-231">Um `Worksheet` objeto tem três propriedades de layout relacionadas: `horizontalPageBreaks`, `verticalPageBreaks`, `pageLayout`.</span><span class="sxs-lookup"><span data-stu-id="04f50-231">A `Worksheet` object has three layout-related properties: `horizontalPageBreaks`, `verticalPageBreaks`, `pageLayout`.</span></span>

<span data-ttu-id="04f50-232">`Worksheet.horizontalPageBreaks` e `Worksheet.verticalPageBreaks` são [PageBreakCollections](/javascript/api/excel/excel.pagebreakcollection).</span><span class="sxs-lookup"><span data-stu-id="04f50-232">`Worksheet.horizontalPageBreaks` and `Worksheet.verticalPageBreaks` are [PageBreakCollections](/javascript/api/excel/excel.pagebreakcollection).</span></span> <span data-ttu-id="04f50-233">Estes são conjuntos [Quebras de página](/javascript/api/excel/excel.pagebreak), que especificam os intervalos em que as quebras de página manuais são inseridas.</span><span class="sxs-lookup"><span data-stu-id="04f50-233">These are collections of [PageBreaks](/javascript/api/excel/excel.pagebreak), which specify ranges where manual page breaks are inserted.</span></span> <span data-ttu-id="04f50-234">O exemplo de código a seguir adiciona uma quebra de página horizontal acima da linha **21**.</span><span class="sxs-lookup"><span data-stu-id="04f50-234">The following code sample adds a horizontal page break above row **21**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.horizontalPageBreaks.add("A21:E21"); // The page break is added above this range.
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="04f50-235">`Worksheet.pageLayout` é um objeto [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="04f50-235">`Worksheet.pageLayout` is a [PageLayout](/javascript/api/excel/excel.pagelayout) object.</span></span> <span data-ttu-id="04f50-236">Esse objeto contém as configurações de layout e impressão que não são dependentes da implementação de qualquer impressora específica.</span><span class="sxs-lookup"><span data-stu-id="04f50-236">This object contains layout and print settings that are not dependent any printer-specific implementation.</span></span> <span data-ttu-id="04f50-237">Essas configurações incluem margens, orientação, numeração de página, linhas de título e a área de impressão.</span><span class="sxs-lookup"><span data-stu-id="04f50-237">These settings include margins, orientation, page numbering, title rows, and print area.</span></span>

<span data-ttu-id="04f50-238">O exemplo de código a seguir centraliza a página (tanto verticalmente quanto horizontalmente), define uma linha de título que será impressa na parte superior de cada página e define a área impressa para a subseção da planilha.</span><span class="sxs-lookup"><span data-stu-id="04f50-238">The following code sample centers the page (both vertically and horizontally), sets a title row that will be printed at the top of every page, and sets the printed area to a subsection of the worksheet.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="04f50-239">Confira também</span><span class="sxs-lookup"><span data-stu-id="04f50-239">See also</span></span>

- [<span data-ttu-id="04f50-240">Modelo de objeto JavaScript do Excel em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="04f50-240">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
