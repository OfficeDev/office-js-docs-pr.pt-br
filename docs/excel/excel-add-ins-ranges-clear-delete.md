---
title: Limpar ou excluir intervalos usando a EXCEL JavaScript
description: Saiba como limpar ou excluir intervalos usando a EXCEL JavaScript.
ms.date: 02/16/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7336a0e6485ce502216818b4a8cd077fed0069c3
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340705"
---
# <a name="clear-or-delete-ranges-using-the-excel-javascript-api"></a>Limpar ou excluir intervalos usando a EXCEL JavaScript

Este artigo fornece exemplos de código que limpam e excluem intervalos com Excel API JavaScript. Para ver a lista completa de propriedades e métodos suportados pelo `Range` objeto, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="clear-a-range-of-cells"></a>Limpar um intervalo de células

O exemplo de código a seguir limpa todo o conteúdo e a formatação das células no intervalo **E2:E5**.  

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("E2:E5");

    range.clear();

    await context.sync();
});
```

### <a name="data-before-range-is-cleared"></a>Dados antes da limpeza do intervalo

![Dados em Excel antes que o intervalo seja limpo.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a>Dados após a limpeza do intervalo

![Os dados Excel depois que o intervalo for limpo.](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a>Excluir um intervalo de células

O exemplo de código a seguir exclui as células no intervalo **B4:E4** e desloca outras células para cima para preencher o espaço que foi desocupado pelas células excluídas.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    await context.sync();
});
```

### <a name="data-before-range-is-deleted"></a>Dados antes da exclusão do intervalo

![Dados no Excel antes que o intervalo seja excluído.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a>Dados após a exclusão do intervalo

![Os dados Excel depois que o intervalo for excluído.](../images/excel-ranges-after-delete.png)

## <a name="see-also"></a>Confira também

- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Definir e obter intervalos usando a EXCEL JavaScript](excel-add-ins-ranges-set-get.md)
- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
