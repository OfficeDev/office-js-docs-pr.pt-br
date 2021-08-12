---
title: Inserir intervalos usando a EXCEL JavaScript
description: Saiba como inserir um intervalo de células com a EXCEL JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad75b8c107005777047418ff9a1824665552cb5cca06c1e858f3645172f12e7c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084642"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>Inserir um intervalo de células usando a EXCEL JavaScript

Este artigo fornece um exemplo de código que insere um intervalo de células com a EXCEL JavaScript. Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte o [Excel. Classe Range](/javascript/api/excel/excel.range).

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

![Dados na Excel antes da inserção do intervalo.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>Dados após a inserção do intervalo

![Dados na Excel após a inserção do intervalo.](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Limpar ou excluir intervalos usando a EXCEL JavaScript](excel-add-ins-ranges-clear-delete.md)
