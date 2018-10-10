---
title: Trabalhar com planilhas usando a API JavaScript do Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 9ceb2187cdd7f503fb39171e420adabcc2f13041
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459130"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="51a90-102">Trabalhar com planilhas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="51a90-102">Work with Worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="51a90-103">Este artigo fornece exemplos de código que mostram como executar tarefas comuns com planilhas usando a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="51a90-103">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span> <span data-ttu-id="51a90-104">Para obter a lista completa de propriedades e métodos aos quais os objetos **Worksheet** e **WorksheetCollection** dão suporte, confira [Objeto Worksheet (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) e [Objeto WorksheetCollection (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="51a90-104">For the complete list of properties and methods that the **Worksheet** and **WorksheetCollection** objects support, see [Worksheet Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=office-js) and [WorksheetCollection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection?view=office-js).</span></span>

> [!NOTE]
> <span data-ttu-id="51a90-105">As informações deste artigo se aplicam apenas a planilhas regulares; elas não se aplicam às folhas "gráfico" ou "macro".</span><span class="sxs-lookup"><span data-stu-id="51a90-105">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="51a90-106">Obter planilhas</span><span class="sxs-lookup"><span data-stu-id="51a90-106">Get worksheets</span></span>

<span data-ttu-id="51a90-107">O exemplo de código a seguir obtém a coleção de planilhas, carrega a propriedade **name** de cada planilha e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="51a90-107">The following code sample gets the collection of worksheets, loads the **name** property of each worksheet, and writes a message to the console.</span></span>

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
> <span data-ttu-id="51a90-108">A propriedade **id** de uma planilha identifica exclusivamente a planilha em uma determinada pasta de trabalho e seu valor permanecerá igual, mesmo quando a planilha é renomeada ou movida..</span><span class="sxs-lookup"><span data-stu-id="51a90-108">The **id** property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved.</span></span> <span data-ttu-id="51a90-109">Quando uma planilha é excluída de uma pasta de trabalho no Excel para Mac, a **id** da planilha excluída pode ser reatribuída a uma nova planilha que é subsequentemente criada.</span><span class="sxs-lookup"><span data-stu-id="51a90-109">When a worksheet is deleted from a workbook in Excel for Mac, the **id** of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="51a90-110">Obter a planilha ativa</span><span class="sxs-lookup"><span data-stu-id="51a90-110">Get the active worksheet</span></span>

<span data-ttu-id="51a90-111">O exemplo de código a seguir obtém a planilha ativa, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="51a90-111">The following code sample gets the active worksheet, loads its **name** property, and writes a message to the console.</span></span>

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

## <a name="set-the-active-worksheet"></a><span data-ttu-id="51a90-112">Definir a planilha ativa</span><span class="sxs-lookup"><span data-stu-id="51a90-112">Set the active worksheet</span></span>

<span data-ttu-id="51a90-113">O exemplo de código a seguir define a planilha ativa para a planilha chamada **Sample**, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="51a90-113">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="51a90-114">Se não houver planilha com esse nome, o método **activate()** gerará um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="51a90-114">If there is no worksheet with that name, the **activate()** method throws an **ItemNotFound** error.</span></span>

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

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="51a90-115">Planilhas de referência por posição relativa</span><span class="sxs-lookup"><span data-stu-id="51a90-115">Reference worksheets by relative position</span></span>

<span data-ttu-id="51a90-116">Esses exemplos mostram como fazer referência a uma planilha por sua posição relativa.</span><span class="sxs-lookup"><span data-stu-id="51a90-116">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="51a90-117">Obter a primeira planilha</span><span class="sxs-lookup"><span data-stu-id="51a90-117">Get the first worksheet</span></span>

<span data-ttu-id="51a90-118">O exemplo de código a seguir obtém a primeira planilha na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="51a90-118">The following code sample gets the first worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="get-the-last-worksheet"></a><span data-ttu-id="51a90-119">Obter a última planilha</span><span class="sxs-lookup"><span data-stu-id="51a90-119">Get the last worksheet</span></span>

<span data-ttu-id="51a90-120">O exemplo de código a seguir obtém a última planilha na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="51a90-120">The following code sample gets the last worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="get-the-next-worksheet"></a><span data-ttu-id="51a90-121">Obter a próxima planilha</span><span class="sxs-lookup"><span data-stu-id="51a90-121">Get the next worksheet</span></span>

<span data-ttu-id="51a90-122">O exemplo de código a seguir obtém a planilha que vem depois da planilha ativa na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="51a90-122">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="51a90-123">Se não houver planilha após a planilha ativa, o método **getNext()** gerará um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="51a90-123">If there is no worksheet after the active worksheet, the **getNext()** method throws an **ItemNotFound** error.</span></span>

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

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="51a90-124">Obter a planilha anterior</span><span class="sxs-lookup"><span data-stu-id="51a90-124">Get the previous worksheet</span></span>

<span data-ttu-id="51a90-125">O exemplo de código a seguir obtém a planilha que precede a planilha ativa na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="51a90-125">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="51a90-126">Se não houver planilha antes da planilha ativa, o método **getPrevious()** gerará um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="51a90-126">If there is no worksheet before the active worksheet, the **getPrevious()** method throws an **ItemNotFound** error.</span></span>

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

## <a name="add-a-worksheet"></a><span data-ttu-id="51a90-127">Adicionar uma planilha</span><span class="sxs-lookup"><span data-stu-id="51a90-127">Add a worksheet</span></span>

<span data-ttu-id="51a90-p106">O exemplo de código a seguir adiciona uma nova planilha chamada **Sample** à pasta de trabalho, carrega suas propriedades **name** e **position** e grava uma mensagem no console. A nova planilha é adicionada após todas as planilhas existentes.</span><span class="sxs-lookup"><span data-stu-id="51a90-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its **name** and **position** properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

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

## <a name="delete-a-worksheet"></a><span data-ttu-id="51a90-130">Excluir uma planilha</span><span class="sxs-lookup"><span data-stu-id="51a90-130">Delete a worksheet</span></span>

<span data-ttu-id="51a90-131">O exemplo de código a seguir exclui a planilha final na pasta de trabalho (desde que ela não seja a única folha na pasta de trabalho) e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="51a90-131">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

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

## <a name="rename-a-worksheet"></a><span data-ttu-id="51a90-132">Renomear uma planilha</span><span class="sxs-lookup"><span data-stu-id="51a90-132">Rename a worksheet</span></span>

<span data-ttu-id="51a90-133">O exemplo de código a seguir altera o nome da planilha ativa para **New Name**.</span><span class="sxs-lookup"><span data-stu-id="51a90-133">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="51a90-134">Mover uma planilha</span><span class="sxs-lookup"><span data-stu-id="51a90-134">Move a worksheet</span></span>

<span data-ttu-id="51a90-135">O exemplo de código a seguir move uma planilha da última posição para a primeira posição na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="51a90-135">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

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

## <a name="set-worksheet-visibility"></a><span data-ttu-id="51a90-136">Definir visibilidade da planilha</span><span class="sxs-lookup"><span data-stu-id="51a90-136">Set worksheet visibility</span></span>

<span data-ttu-id="51a90-137">Esses exemplos mostram como definir a visibilidade de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="51a90-137">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="51a90-138">Ocultar uma planilha</span><span class="sxs-lookup"><span data-stu-id="51a90-138">Hide a worksheet</span></span>

<span data-ttu-id="51a90-139">O exemplo de código a seguir define a visibilidade da planilha chamada **Sample** para oculta, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="51a90-139">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="unhide-a-worksheet"></a><span data-ttu-id="51a90-140">Reexibir uma planilha</span><span class="sxs-lookup"><span data-stu-id="51a90-140">Unhide a worksheet</span></span>

<span data-ttu-id="51a90-141">O exemplo de código a seguir define a visibilidade da planilha chamada **Sample** para visível, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="51a90-141">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its **name** property, and writes a message to the console.</span></span>

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

## <a name="get-a-cell-within-a-worksheet"></a><span data-ttu-id="51a90-142">Obter uma célula em uma planilha</span><span class="sxs-lookup"><span data-stu-id="51a90-142">Get a cell within a worksheet</span></span>

<span data-ttu-id="51a90-143">O exemplo de código a seguir obtém a célula que está localizada na linha 2, coluna 5 da planilha chamada **Sample**, carrega suas propriedades **address** e **values** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="51a90-143">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its **address** and **values** properties, and writes a message to the console.</span></span> <span data-ttu-id="51a90-144">Os valores que são passados no método **getCell(row: number, column:number)** são número de linha e número de coluna indexados por zero para a célula que está sendo recuperada.</span><span class="sxs-lookup"><span data-stu-id="51a90-144">The values that are passed into the **getCell(row: number, column:number)** method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

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

## <a name="get-a-range-within-a-worksheet"></a><span data-ttu-id="51a90-145">Obter um intervalo em uma planilha</span><span class="sxs-lookup"><span data-stu-id="51a90-145">Get a range within a worksheet</span></span>

<span data-ttu-id="51a90-146">Confira exemplos que mostram como obter um intervalo em uma planilha em [Trabalhar com intervalos usando a API JavaScript do Excel](excel-add-ins-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="51a90-146">For examples that show how to get a range within a worksheet, see [Work with Ranges using the Excel JavaScript API](excel-add-ins-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="51a90-147">Confira também</span><span class="sxs-lookup"><span data-stu-id="51a90-147">See also</span></span>

- [<span data-ttu-id="51a90-148">Conceitos de programação fundamentais com a API do JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="51a90-148">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

