---
title: Obter um intervalo usando a EXCEL JavaScript
description: Saiba como recuperar um intervalo usando a Excel JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 16c42ccf8f3496316fbf7b52e4d8139f819c6da1
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340936"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a>Obter um intervalo usando a EXCEL JavaScript

Este artigo fornece exemplos que mostram diferentes maneiras de obter um intervalo dentro de uma planilha usando o Excel API JavaScript. Para ver a lista completa de propriedades e métodos compatíveis `Range` com o objeto, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a>Obter intervalo por endereço

O exemplo de código a seguir obtém o intervalo com o endereço **B2:C5** da planilha chamada **Sample**, `address` carrega sua propriedade e grava uma mensagem no console.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    
    let range = sheet.getRange("B2:C5");
    range.load("address");
    await context.sync();
    
    console.log(`The address of the range B2:C5 is "${range.address}"`);
});
```

## <a name="get-range-by-name"></a>Obter intervalo por nome

O exemplo de código a seguir obtém o intervalo `MyRange` nomeado da planilha denominada **Exemplo**, `address` carrega sua propriedade e grava uma mensagem no console.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("MyRange");
    range.load("address");
    await context.sync();

    console.log(`The address of the range "MyRange" is "${range.address}"`);
});
```

## <a name="get-used-range"></a>Obter intervalo usado

O exemplo de código a seguir obtém o intervalo usado da planilha denominada **Exemplo**, `address` carrega sua propriedade e grava uma mensagem no console. O intervalo usado é o menor intervalo que abrange todas as células na planilha que têm um valor ou uma formatação atribuída a elas. Se a planilha inteira estiver em branco, o `getUsedRange()` método retornará um intervalo que consiste apenas na célula superior esquerda.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getUsedRange();
    range.load("address");
    await context.sync();
    
    console.log(`The address of the used range in the worksheet is "${range.address}"`);
});
```

## <a name="get-entire-range"></a>Obter intervalo inteiro

O exemplo de código a seguir obtém todo o intervalo de planilhas da planilha denominada **Exemplo**, `address` carrega sua propriedade e grava uma mensagem no console.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange();
    range.load("address");
    await context.sync();
    
    console.log(`The address of the entire worksheet range is "${range.address}"`);
});
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Inserir um intervalo usando a EXCEL JavaScript](excel-add-ins-ranges-insert.md)
