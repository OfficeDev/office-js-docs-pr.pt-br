---
title: Manipular matrizes dinâmicas e vazamento de intervalo usando a API JavaScript Excel JavaScript
description: Saiba como lidar com o vazamento de matrizes dinâmicas e intervalos com Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 4601cd6a901243cadab0e7c5ead6061e28806377
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149011"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a>Manipular matrizes dinâmicas e vazamento usando a api javascript Excel javascript

Este artigo fornece um exemplo de código que lida com matrizes dinâmicas e vazamento de intervalo usando a API JavaScript Excel javascript. Para ver a lista completa de propriedades e métodos que o `Range` objeto oferece suporte, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

## <a name="dynamic-arrays"></a>Matrizes dinâmicas

Algumas Excel retornam [matrizes dinâmicas.](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) Eles preenchem os valores de várias células fora da célula original da fórmula. Esse estouro de valor é chamado de "vazamento". Seu complemento pode encontrar o intervalo usado para um vazamento com o [método Range.getSpillingToRange.](/javascript/api/excel/excel.range#getSpillingToRange__) Há também uma [versão *OrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject` .

O exemplo a seguir mostra uma fórmula básica que copia o conteúdo de um intervalo em uma célula, que se espalha em células vizinhas. Em seguida, o complemento registra o intervalo que contém o vazamento.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

## <a name="range-spilling"></a>Vazamento de intervalo

Encontre a célula responsável pelo vazamento em uma determinada célula usando o [método Range.getSpillParent.](/javascript/api/excel/excel.range#getSpillParent__) Observe que `getSpillParent` só funciona quando o objeto range é uma única célula. Chamar em um intervalo com várias células resultará em um erro sendo lançado (ou um intervalo `getSpillParent` nulo sendo retornado para `Range.getSpillParentOrNullObject` ).

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
