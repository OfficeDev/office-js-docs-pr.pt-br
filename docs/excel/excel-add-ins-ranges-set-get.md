---
title: Definir e obter o intervalo selecionado usando a API JavaScript do Excel
description: Saiba como usar a API JavaScript do Excel para definir e obter intervalos usando a API JavaScript do Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 06b6219924f0667ecef57d608cb417a76ef8031d
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652755"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a>Definir e obter intervalos usando a API JavaScript do Excel

Este artigo fornece exemplos de código que definir e obter intervalos com a API JavaScript do Excel. Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a>Definir o intervalo selecionado

O exemplo de código a seguir seleciona o intervalo **B2:E6** na planilha ativa.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a>Intervalo selecionado B2:E6

![Intervalo selecionado no Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a>Obter o intervalo selecionado

O exemplo de código a seguir obtém o intervalo selecionado, carrega `address` sua propriedade e grava uma mensagem no console.

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a API JavaScript do Excel](excel-add-ins-cells.md)
- [Definir e obter valores de intervalo, texto ou fórmulas usando a API JavaScript do Excel](excel-add-ins-ranges-set-get-values.md)
- [Definir o formato de intervalo usando a API JavaScript do Excel](excel-add-ins-ranges-set-format.md)
