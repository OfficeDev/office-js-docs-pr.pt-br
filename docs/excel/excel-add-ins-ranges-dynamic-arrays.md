---
title: Manipular matrizes dinâmicas e vazamento de intervalo usando a API JavaScript do Excel
description: Saiba como lidar com o vazamento de matrizes dinâmicas e intervalos com a API JavaScript do Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c224fc336791440911519a6d24aee6c208d90c9e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652774"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a>Manipular matrizes dinâmicas e vazamento usando a API JavaScript do Excel

Este artigo fornece um exemplo de código que lida com matrizes dinâmicas e vazamento de intervalo usando a API JavaScript do Excel. Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).

## <a name="dynamic-arrays"></a>Matrizes dinâmicas

Algumas fórmulas do Excel [retornam matrizes dinâmicas.](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531) Eles preenchem os valores de várias células fora da célula original da fórmula. Esse estouro de valor é chamado de "vazamento". Seu complemento pode encontrar o intervalo usado para um vazamento com o [método Range.getSpillingToRange.](/javascript/api/excel/excel.range#getspillingtorange--) Há também uma [versão *OrNullObject](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject` .

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

Encontre a célula responsável pelo vazamento em uma determinada célula usando o [método Range.getSpillParent.](/javascript/api/excel/excel.range#getspillparent--) Observe que `getSpillParent` só funciona quando o objeto range é uma única célula. Chamar em um intervalo com várias células resultará em um erro sendo lançado (ou um intervalo `getSpillParent` nulo sendo retornado para `Range.getSpillParentOrNullObject` ).

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a API JavaScript do Excel](excel-add-ins-cells.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
