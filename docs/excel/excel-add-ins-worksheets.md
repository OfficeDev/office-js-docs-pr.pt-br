---
title: Trabalhe com planilhas usando a API JavaScript do Excel
description: ''
ms.date: 12/28/2018
localization_priority: Priority
ms.openlocfilehash: 62f64beaefcc938f91ee581594922b2c965f2655
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389526"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a>Trabalhe com planilhas usando a API JavaScript do Excel

Este artigo fornece exemplos de código que mostram como executar tarefas comuns com planilhas usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos aos quais os objetos **Worksheet** e **WorksheetCollection** dão suporte, confira [Objeto Worksheet (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) e [Objeto WorksheetCollection (API JavaScript para Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection).

> [!NOTE]
> As informações deste artigo se aplicam apenas a planilhas regulares; elas não se aplicam às folhas "gráfico" ou "macro".

## <a name="get-worksheets"></a>Obter planilhas

O exemplo de código a seguir obtém a coleção de planilhas, carrega a propriedade **name** de cada planilha e grava uma mensagem no console.

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
> A propriedade **id** de uma planilha identifica exclusivamente a planilha em uma determinada pasta de trabalho e seu valor permanecerá igual, mesmo quando a planilha for renomeada ou movida. Quando uma planilha é excluída de uma pasta de trabalho no Excel para Mac, a **id** da planilha excluída pode ser reatribuída a uma nova planilha que é subsequentemente criada.

## <a name="get-the-active-worksheet"></a>Obter a planilha ativa

O exemplo de código a seguir obtém a planilha ativa, carrega sua propriedade **name** e grava uma mensagem no console.

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

## <a name="set-the-active-worksheet"></a>Definir a planilha ativa

O exemplo de código a seguir define a planilha ativa para a planilha chamada **Amostra**, carrega sua propriedade **name** e grava uma mensagem no console. Se não houver planilha com esse nome, o método **activate()** gerará um erro **ItemNotFound**.

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

## <a name="reference-worksheets-by-relative-position"></a>Planilhas de referência por posição relativa

Esses exemplos mostram como fazer referência a uma planilha por sua posição relativa.

### <a name="get-the-first-worksheet"></a>Obter a primeira planilha

O exemplo de código a seguir obtém a primeira planilha na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console.

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

### <a name="get-the-last-worksheet"></a>Obter a última planilha

O exemplo de código a seguir obtém a última planilha na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console.

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

### <a name="get-the-next-worksheet"></a>Obter a próxima planilha

O exemplo de código a seguir obtém a planilha que vem depois da planilha ativa na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console. Se não houver planilha após a planilha ativa, o método **getNext()** gerará um erro **ItemNotFound**.

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

### <a name="get-the-previous-worksheet"></a>Obter a planilha anterior

O exemplo de código a seguir obtém a planilha que precede a planilha ativa na pasta de trabalho, carrega sua propriedade **name** e grava uma mensagem no console. Se não houver planilha antes da planilha ativa, o método **getPrevious()** gerará um erro **ItemNotFound**.

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

## <a name="add-a-worksheet"></a>Adicionar uma planilha

O exemplo de código a seguir adiciona uma nova planilha chamada **Amostra** à pasta de trabalho, carrega suas propriedades **name** e **position** e grava uma mensagem no console. A nova planilha é adicionada após todas as planilhas existentes.

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

## <a name="delete-a-worksheet"></a>Excluir uma planilha

O exemplo de código a seguir exclui a planilha final na pasta de trabalho (desde que ela não seja a única folha na pasta de trabalho) e grava uma mensagem no console.

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

## <a name="rename-a-worksheet"></a>Renomear uma planilha

O exemplo de código a seguir altera o nome da planilha ativa para **Novo Nome**.

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a>Mover uma planilha

O exemplo de código a seguir move uma planilha da última posição para a primeira posição na pasta de trabalho.

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

## <a name="set-worksheet-visibility"></a>Definir visibilidade da planilha

Esses exemplos mostram como definir a visibilidade de uma planilha.

### <a name="hide-a-worksheet"></a>Ocultar uma planilha

O exemplo de código a seguir define a visibilidade da planilha chamada **Amostra** para oculta, carrega sua propriedade **name** e grava uma mensagem no console.

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

### <a name="unhide-a-worksheet"></a>Reexibir uma planilha

O exemplo de código a seguir define a visibilidade da planilha chamada **Amostra** para visível, carrega sua propriedade **name** e grava uma mensagem no console.

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

## <a name="get-a-single-cell-within-a-worksheet"></a>Obter uma única célula em uma planilha

O exemplo de código a seguir obtém a célula que está localizada na linha 2, coluna 5 da planilha chamada **Amostra**, carrega suas propriedades **address** e **values** e grava uma mensagem no console. Os valores que são passados no método `getCell(row: number, column:number)` são número de linha e número de coluna indexados por zero para a célula que está sendo recuperada.

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

## <a name="find-all-cells-with-matching-text-preview"></a>Encontrar todas as células com texto correspondente (versão prévia)

> [!NOTE]
> A função `findAll` do objeto da planilha só está disponível atualmente na versão prévia pública (beta). Para usar esse recurso, você deve usar a biblioteca beta do CDN do Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> Se você estiver usando o TypeScript ou se seu editor de código usar arquivos de definição de tipo do TypeScript do IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.

O objeto `Worksheet` tem o método `find` para pesquisar uma cadeia especificada dentro da planilha. Ele retorna um objeto `RangeAreas`, que é um conjunto de objetos `Range` que podem ser editados ao mesmo tempo. O exemplo de código a seguir localiza todas as células com valores iguais à cadeia de caracteres **Concluída** e os marca de verde. Observe que `findAll` exibirá um erro `ItemNotFound` se a cadeia especificada não existir na planilha. Se você acha que a cadeia especificada pode não estar na planilha, use o método [findAllOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) para que seu código manipule normalmente esse cenário.

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
> Esta seção descreve como localizar as células e intervalos usando as funções do objeto `Worksheet`. Encontre mais informações de recuperação de intervalo nos artigos específicos do objeto.
> - Confira os exemplos que mostram como obter um intervalo em uma planilha usando o objeto `Range` em [Trabalhar com intervalos usando a API JavaScript do Excel](excel-add-ins-ranges.md).
> - Para obter exemplos que mostram como obter intervalos de um objeto `Table`, confira [Trabalhar com tabelas usando a API JavaScript do Excel](excel-add-ins-tables.md).
> - Para obter exemplos que mostram como pesquisar um grande intervalo para vários subgrupos com base nas características da célula, confira [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md).

## <a name="data-protection"></a>Proteção de dados

O suplemento pode controlar a capacidade de um usuário de editar dados em uma planilha. A propriedade `protection` da planilha é um objeto [WorksheetProtection](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotection) com um método `protect()`. O exemplo a seguir mostra um cenário básico ativando/desativando a proteção completa da planilha ativa.

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

O método `protect` tem dois parâmetros opcionais:

- `options`: Um objeto [WorksheetProtectionOptions](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotectionoptions) definindo restrições de edição de específicas.
- `password`: Uma cadeia de caracteres que representa a senha necessária para um usuário ignorar a proteção e editar a planilha.

O artigo [Proteger uma planilha](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) tem mais informações sobre a proteção de planilhas e sobre como alterar na interface do usuário do Excel.

## <a name="see-also"></a>Confira também

- [Conceitos fundamentais de programação com a API JavaScript do Excel](excel-add-ins-core-concepts.md)
