---
title: Trabalhe com planilhas usando a API JavaScript do Excel
description: ''
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: 825ae88afd98afbcd268716c93ddcb13d24a9a1e
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871553"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="61464-102">Trabalhe com planilhas usando a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="61464-102">Work with worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="61464-p101">Este artigo fornece exemplos de código que mostram como executar tarefas comuns com planilhas usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos aos quais os objetos **Worksheet** e **WorksheetCollection** dão suporte, confira [Objeto Worksheet (API JavaScript para Excel)](/javascript/api/excel/excel.worksheet) e [Objeto WorksheetCollection (API JavaScript para Excel)](/javascript/api/excel/excel.worksheetcollection).</span><span class="sxs-lookup"><span data-stu-id="61464-p101">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API. For the complete list of properties and methods that the **Worksheet** and **WorksheetCollection** objects support, see [Worksheet Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheet) and [WorksheetCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheetcollection).</span></span>

> [!NOTE]
> <span data-ttu-id="61464-105">As informações deste artigo se aplicam apenas a planilhas regulares; elas não se aplicam às folhas "gráfico" ou "macro".</span><span class="sxs-lookup"><span data-stu-id="61464-105">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="61464-106">Obter planilhas</span><span class="sxs-lookup"><span data-stu-id="61464-106">Get worksheets</span></span>

<span data-ttu-id="61464-107">O exemplo de código a seguir obtém a coleção de planilhas, carrega a propriedade **name** de cada planilha e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="61464-107">The following code sample gets the collection of worksheets, loads the **name** property of each worksheet, and writes a message to the console.</span></span>

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
> <span data-ttu-id="61464-p102">A propriedade **id** de uma planilha identifica exclusivamente a planilha em uma determinada pasta de trabalho e seu valor permanecerá igual, mesmo quando a planilha for renomeada ou movida. Quando uma planilha é excluída de uma pasta de trabalho no Excel para Mac, a **id** da planilha excluída pode ser reatribuída a uma nova planilha que é subsequentemente criada.</span><span class="sxs-lookup"><span data-stu-id="61464-p102">The **id** property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved. When a worksheet is deleted from a workbook in Excel for Mac, the **id** of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="61464-110">Obter a planilha ativa</span><span class="sxs-lookup"><span data-stu-id="61464-110">Get the active worksheet</span></span>

<span data-ttu-id="61464-111">O exemplo de código a seguir obtém a planilha ativa, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="61464-111">The following code sample gets the active worksheet, loads its **name** property, and writes a message to the console.</span></span>

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

## <a name="set-the-active-worksheet"></a><span data-ttu-id="61464-112">Definir a planilha ativa</span><span class="sxs-lookup"><span data-stu-id="61464-112">Set the active worksheet</span></span>

<span data-ttu-id="61464-p103">O exemplo de código a seguir define a planilha ativa para a planilha chamada **Amostra**, carrega sua propriedade **name** e grava uma mensagem no console. Se não houver planilha com esse nome, o método **activate()** gerará um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="61464-p103">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its **name** property, and writes a message to the console. If there is no worksheet with that name, the **activate()** method throws an **ItemNotFound** error.</span></span>

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

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="61464-115">Planilhas de referência por posição relativa</span><span class="sxs-lookup"><span data-stu-id="61464-115">Reference worksheets by relative position</span></span>

<span data-ttu-id="61464-116">Esses exemplos mostram como fazer referência a uma planilha por sua posição relativa.</span><span class="sxs-lookup"><span data-stu-id="61464-116">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="61464-117">Obter a primeira planilha</span><span class="sxs-lookup"><span data-stu-id="61464-117">Get the first worksheet</span></span>

<span data-ttu-id="61464-118">O exemplo de código a seguir obtém a primeira planilha na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="61464-118">The following code sample gets the first worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="get-the-last-worksheet"></a><span data-ttu-id="61464-119">Obter a última planilha</span><span class="sxs-lookup"><span data-stu-id="61464-119">Get the last worksheet</span></span>

<span data-ttu-id="61464-120">O exemplo de código a seguir obtém a última planilha na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="61464-120">The following code sample gets the last worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="get-the-next-worksheet"></a><span data-ttu-id="61464-121">Obter a próxima planilha</span><span class="sxs-lookup"><span data-stu-id="61464-121">Get the next worksheet</span></span>

<span data-ttu-id="61464-p104">O exemplo de código a seguir obtém a planilha que vem depois da planilha ativa na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console. Se não houver planilha após a planilha ativa, o método **getNext()** gerará um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="61464-p104">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its **name** property, and writes a message to the console. If there is no worksheet after the active worksheet, the **getNext()** method throws an **ItemNotFound** error.</span></span>

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

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="61464-124">Obter a planilha anterior</span><span class="sxs-lookup"><span data-stu-id="61464-124">Get the previous worksheet</span></span>

<span data-ttu-id="61464-p105">O exemplo de código a seguir obtém a planilha que precede a planilha ativa na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console. Se não houver planilha antes da planilha ativa, o método **getPrevious()** gerará um erro **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="61464-p105">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its **name** property, and writes a message to the console. If there is no worksheet before the active worksheet, the **getPrevious()** method throws an **ItemNotFound** error.</span></span>

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

## <a name="add-a-worksheet"></a><span data-ttu-id="61464-127">Adicionar uma planilha</span><span class="sxs-lookup"><span data-stu-id="61464-127">Add a worksheet</span></span>

<span data-ttu-id="61464-p106">O exemplo de código a seguir adiciona uma nova planilha chamada **Amostra** à pasta de trabalho, carrega suas propriedades **name** e **position** e grava uma mensagem no console. A nova planilha é adicionada após todas as planilhas existentes.</span><span class="sxs-lookup"><span data-stu-id="61464-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its **name** and **position** properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

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

## <a name="delete-a-worksheet"></a><span data-ttu-id="61464-130">Excluir uma planilha</span><span class="sxs-lookup"><span data-stu-id="61464-130">Delete a worksheet</span></span>

<span data-ttu-id="61464-131">O exemplo de código a seguir exclui a planilha final na pasta de trabalho (desde que ela não seja a única folha na pasta de trabalho) e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="61464-131">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

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
> <span data-ttu-id="61464-132">Uma planilha com visibilidade de "[Muito oculta](/javascript/api/excel/excel.sheetvisibility)" não pode ser excluída com o método `delete`.</span><span class="sxs-lookup"><span data-stu-id="61464-132">A worksheet with a visibility of "[Very Hidden](/javascript/api/excel/excel.sheetvisibility)" cannot be deleted with the `delete` method.</span></span> <span data-ttu-id="61464-133">Se você quiser excluir a planilha de qualquer forma, deverá primeiro alterar a visibilidade.</span><span class="sxs-lookup"><span data-stu-id="61464-133">If you wish to delete the worksheet anyway, you must first change the visibility.</span></span>

## <a name="rename-a-worksheet"></a><span data-ttu-id="61464-134">Renomear uma planilha</span><span class="sxs-lookup"><span data-stu-id="61464-134">Rename a worksheet</span></span>

<span data-ttu-id="61464-135">O exemplo de código a seguir altera o nome da planilha ativa para **Novo Nome**.</span><span class="sxs-lookup"><span data-stu-id="61464-135">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="61464-136">Mover uma planilha</span><span class="sxs-lookup"><span data-stu-id="61464-136">Move a worksheet</span></span>

<span data-ttu-id="61464-137">O exemplo de código a seguir move uma planilha da última posição para a primeira posição na pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="61464-137">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

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

## <a name="set-worksheet-visibility"></a><span data-ttu-id="61464-138">Definir visibilidade da planilha</span><span class="sxs-lookup"><span data-stu-id="61464-138">Set worksheet visibility</span></span>

<span data-ttu-id="61464-139">Esses exemplos mostram como definir a visibilidade de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="61464-139">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="61464-140">Ocultar uma planilha</span><span class="sxs-lookup"><span data-stu-id="61464-140">Hide a worksheet</span></span>

<span data-ttu-id="61464-141">O exemplo de código a seguir define a visibilidade da planilha chamada **Amostra** para oculta, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="61464-141">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="unhide-a-worksheet"></a><span data-ttu-id="61464-142">Reexibir uma planilha</span><span class="sxs-lookup"><span data-stu-id="61464-142">Unhide a worksheet</span></span>

<span data-ttu-id="61464-143">O exemplo de código a seguir define a visibilidade da planilha chamada **Amostra** para visível, carrega sua propriedade **name** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="61464-143">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its **name** property, and writes a message to the console.</span></span>

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

## <a name="get-a-single-cell-within-a-worksheet"></a><span data-ttu-id="61464-144">Obter uma única célula em uma planilha</span><span class="sxs-lookup"><span data-stu-id="61464-144">Get a single cell within a worksheet</span></span>

<span data-ttu-id="61464-145">O exemplo de código a seguir obtém a célula que está localizada na linha 2, coluna 5 da planilha chamada **Amostra**, carrega suas propriedades **address** e **values** e grava uma mensagem no console.</span><span class="sxs-lookup"><span data-stu-id="61464-145">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its **address** and **values** properties, and writes a message to the console.</span></span> <span data-ttu-id="61464-146">Os valores que são passados no método `getCell(row: number, column:number)` são número de linha e número de coluna indexados por zero para a célula que está sendo recuperada.</span><span class="sxs-lookup"><span data-stu-id="61464-146">The values that are passed into the `getCell(row: number, column:number)` method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

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

## <a name="find-all-cells-with-matching-text-preview"></a><span data-ttu-id="61464-147">Encontrar todas as células com texto correspondente (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="61464-147">Find all cells with matching text (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="61464-148">A função `findAll` do objeto da planilha só está disponível atualmente na versão prévia pública.</span><span class="sxs-lookup"><span data-stu-id="61464-148">The Worksheet object's `findAll` function is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="61464-149">O objeto `Worksheet` tem o método `find` para pesquisar uma cadeia especificada dentro da planilha.</span><span class="sxs-lookup"><span data-stu-id="61464-149">The `Worksheet` object has a `find` method to search for a specified string within the worksheet.</span></span> <span data-ttu-id="61464-150">Ele retorna um objeto `RangeAreas`, que é um conjunto de objetos `Range` que podem ser editados ao mesmo tempo.</span><span class="sxs-lookup"><span data-stu-id="61464-150">It returns a `RangeAreas` object, which is a collection of `Range` objects that can be edited all at once.</span></span> <span data-ttu-id="61464-151">O exemplo de código a seguir localiza todas as células com valores iguais à cadeia de caracteres **Concluída** e os marca de verde.</span><span class="sxs-lookup"><span data-stu-id="61464-151">The following code sample finds all cells with values equal to the string **Complete** and colors them green.</span></span> <span data-ttu-id="61464-152">Observe que `findAll` exibirá um erro `ItemNotFound` se a cadeia especificada não existir na planilha.</span><span class="sxs-lookup"><span data-stu-id="61464-152">Note that `findAll` will throw an `ItemNotFound` error if the specified string doesn't exist in the worksheet.</span></span> <span data-ttu-id="61464-153">Se você acha que a cadeia especificada pode não estar na planilha, use o método [findAllOrNullObject](excel-add-ins-advanced-concepts.md#ornullobject-methods) para que seu código manipule normalmente esse cenário.</span><span class="sxs-lookup"><span data-stu-id="61464-153">If you expect that the specified string may not exist in the worksheet, use the [findAllOrNullObject](excel-add-ins-advanced-concepts.md#ornullobject-methods) method instead, so your code gracefully handles that scenario.</span></span>

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
> <span data-ttu-id="61464-154">Esta seção descreve como localizar as células e intervalos usando as funções do objeto `Worksheet`.</span><span class="sxs-lookup"><span data-stu-id="61464-154">This section describes how to find cells and ranges using the `Worksheet` object's functions.</span></span> <span data-ttu-id="61464-155">Encontre mais informações de recuperação de intervalo nos artigos específicos do objeto.</span><span class="sxs-lookup"><span data-stu-id="61464-155">More range retrieval information can be found in object-specific articles.</span></span>
> - <span data-ttu-id="61464-156">Confira os exemplos que mostram como obter um intervalo em uma planilha usando o objeto `Range` em [Trabalhar com intervalos usando a API JavaScript do Excel](excel-add-ins-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="61464-156">For examples that show how to get a range within a worksheet using the `Range` object, see [Work with ranges using the Excel JavaScript API](excel-add-ins-ranges.md).</span></span>
> - <span data-ttu-id="61464-157">Para obter exemplos que mostram como obter intervalos de um objeto `Table`, confira [Trabalhar com tabelas usando a API JavaScript do Excel](excel-add-ins-tables.md).</span><span class="sxs-lookup"><span data-stu-id="61464-157">For examples that show how to get ranges from a `Table` object, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).</span></span>
> - <span data-ttu-id="61464-158">Para obter exemplos que mostram como pesquisar um grande intervalo para vários subgrupos com base nas características da célula, confira [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="61464-158">For examples that show how to search a large range for multiple sub-ranges based on cell characteristics, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="data-protection"></a><span data-ttu-id="61464-159">Proteção de dados</span><span class="sxs-lookup"><span data-stu-id="61464-159">Data protection</span></span>

<span data-ttu-id="61464-160">O suplemento pode controlar a capacidade de um usuário de editar dados em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="61464-160">Your add-in can control a user's ability to edit data in a worksheet.</span></span> <span data-ttu-id="61464-161">A propriedade `protection` da planilha é um objeto [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) com um método `protect()`.</span><span class="sxs-lookup"><span data-stu-id="61464-161">The worksheet's `protection` property is a [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) object with a `protect()` method.</span></span> <span data-ttu-id="61464-162">O exemplo a seguir mostra um cenário básico ativando/desativando a proteção completa da planilha ativa.</span><span class="sxs-lookup"><span data-stu-id="61464-162">The following example shows a basic scenario toggling the complete protection of the active worksheet.</span></span>

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

<span data-ttu-id="61464-163">O método `protect` tem dois parâmetros opcionais:</span><span class="sxs-lookup"><span data-stu-id="61464-163">The `protect` method has two optional parameters:</span></span>

- <span data-ttu-id="61464-164">`options`: Um objeto [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) definindo restrições de edição de específicas.</span><span class="sxs-lookup"><span data-stu-id="61464-164">`options`: A [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) object defining specific editing restrictions.</span></span>
- <span data-ttu-id="61464-165">`password`: Uma cadeia de caracteres que representa a senha necessária para um usuário ignorar a proteção e editar a planilha.</span><span class="sxs-lookup"><span data-stu-id="61464-165">`password`: A string representing the password needed for a user to bypass protection and edit the worksheet.</span></span>

<span data-ttu-id="61464-166">O artigo [Proteger uma planilha](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) tem mais informações sobre a proteção de planilhas e sobre como alterar na interface do usuário do Excel.</span><span class="sxs-lookup"><span data-stu-id="61464-166">The article [Protect a worksheet](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) has more information about worksheet protection and how to change it through the Excel UI.</span></span>

## <a name="see-also"></a><span data-ttu-id="61464-167">Confira também</span><span class="sxs-lookup"><span data-stu-id="61464-167">See also</span></span>

- [<span data-ttu-id="61464-168">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="61464-168">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
