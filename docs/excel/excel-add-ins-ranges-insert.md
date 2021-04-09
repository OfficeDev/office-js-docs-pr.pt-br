---
title: Inserir intervalos usando a API JavaScript do Excel
description: Saiba como inserir um intervalo de células com a API JavaScript do Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 401a08dd10b3775012738ab9c80ec6ab367555ec
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652768"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>Inserir um intervalo de células usando a API JavaScript do Excel

Este artigo fornece um exemplo de código que insere um intervalo de células com a API JavaScript do Excel. Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte a [classe Excel.Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a>Inserir um intervalo de células

O exemplo de código a seguir insere um intervalo de células no local **B4:E4** e desloca outras células para baixo a fim de fornecer espaço para as novas células.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a>Dados antes da inserção do intervalo

![Dados no Excel antes da inserção do intervalo](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>Dados após a inserção do intervalo

![Dados no Excel após a inserção do intervalo](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a API JavaScript do Excel](excel-add-ins-cells.md)
- [Limpar ou excluir intervalos usando a API JavaScript do Excel](excel-add-ins-ranges-clear-delete.md)
