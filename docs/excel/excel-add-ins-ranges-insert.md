---
title: Inserir intervalos usando a EXCEL JavaScript
description: Saiba como inserir um intervalo de células com a EXCEL JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 0e1ed6d2302bcdb4a11688cd6d77448811f8a93b
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340544"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>Inserir um intervalo de células usando a EXCEL JavaScript

Este artigo fornece um exemplo de código que insere um intervalo de células com a EXCEL JavaScript. Para ver a lista completa de propriedades e métodos `Range` compatíveis com o objeto, consulte o Excel[. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a>Inserir um intervalo de células

O exemplo de código a seguir insere um intervalo de células no local **B4:E4** e desloca outras células para baixo a fim de fornecer espaço para as novas células.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    await context.sync();
});
```

### <a name="data-before-range-is-inserted"></a>Dados antes da inserção do intervalo

![Dados em Excel antes da inserção do intervalo.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>Dados após a inserção do intervalo

![Dados no Excel após a inserção do intervalo.](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Limpar ou excluir intervalos usando a EXCEL JavaScript](excel-add-ins-ranges-clear-delete.md)
