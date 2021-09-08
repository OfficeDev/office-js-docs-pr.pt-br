---
title: Obter um intervalo usando a EXCEL JavaScript
description: Saiba como recuperar um intervalo usando a EXCEL JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6aa9bb00bc9d24aeee5f1fef9e8d1531525e9d1f
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938173"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a>Obter um intervalo usando a EXCEL JavaScript

Este artigo fornece exemplos que mostram diferentes maneiras de obter um intervalo dentro de uma planilha usando a api javascript Excel javascript. Para ver a lista completa de propriedades e métodos que o `Range` objeto oferece suporte, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a>Obter intervalo por endereço

O exemplo de código a seguir obtém o intervalo com o endereço **B2:C5** da planilha denominada **Exemplo**, carrega sua propriedade e grava uma mensagem `address` no console.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-range-by-name"></a>Obter intervalo por nome

O exemplo de código a seguir obtém o intervalo nomeado da planilha denominada Exemplo , carrega sua propriedade e grava `MyRange` uma mensagem no  `address` console.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-used-range"></a>Obter intervalo usado

O exemplo de código a seguir obtém o intervalo usado da planilha denominada **Exemplo**, carrega sua propriedade e grava `address` uma mensagem no console. O intervalo usado é o menor intervalo que abrange todas as células na planilha que têm um valor ou uma formatação atribuída a elas. Se a planilha inteira estiver em branco, o método retornará um intervalo que `getUsedRange()` consiste apenas na célula superior esquerda.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-entire-range"></a>Obter intervalo inteiro

O exemplo de código a seguir obtém todo o intervalo de planilhas da planilha denominada **Exemplo**, carrega sua propriedade e grava `address` uma mensagem no console.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Inserir um intervalo usando a EXCEL JavaScript](excel-add-ins-ranges-insert.md)
