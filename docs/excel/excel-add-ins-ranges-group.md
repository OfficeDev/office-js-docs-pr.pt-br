---
title: Intervalos de grupo usando a EXCEL JavaScript
description: Saiba como agrupar linhas ou colunas de um intervalo para criar um contorno usando Excel API JavaScript.
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5e5ba5e5d539d4391b65a792d05ae2f0d2ff49fc63725bc3681cb9015873bb45
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084050"
---
# <a name="group-ranges-for-an-outline-using-the-excel-javascript-api"></a>Intervalos de grupo para um contorno usando a EXCEL JavaScript

Este artigo fornece um exemplo de código que mostra como agrupar intervalos para um contorno usando a API JavaScript Excel JavaScript. Para ver a lista completa de propriedades e métodos que o `Range` objeto oferece suporte, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

## <a name="group-rows-or-columns-of-a-range-for-an-outline"></a>Linhas de grupo ou colunas de um intervalo para um contorno

Linhas ou colunas de um intervalo podem ser agrupadas para criar um [contorno](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF). Esses grupos podem ser recolhidos e expandidos para ocultar e mostrar as células correspondentes. Isso facilita a análise rápida dos dados de linha superior. Use [Range.group](/javascript/api/excel/excel.range#group_groupOption_) para fazer esses grupos de contornos.

Um contorno pode ter uma hierarquia, onde grupos menores são aninhados em grupos maiores. Isso permite que o contorno seja exibido em diferentes níveis. Alterar o nível de contorno visível pode ser feito programaticamente por meio do [método Worksheet.showOutlineLevels.](/javascript/api/excel/excel.worksheet#showOutlineLevels_rowLevels__columnLevels_) Observe que Excel suporta apenas oito níveis de grupos de contornos.

O exemplo de código a seguir cria um contorno com dois níveis de grupos para as linhas e colunas. A imagem subsequente mostra os agrupamentos desse contorno. No exemplo de código, os intervalos que estão sendo agrupados não incluem a linha ou coluna do controle de contorno (os "Totais" deste exemplo). Um grupo define o que será recolhido, não a linha ou coluna com o controle.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);
```

![Intervalo com um contorno de dois níveis e duas dimensões.](../images/excel-outline.png)

## <a name="remove-grouping-from-rows-or-columns-of-a-range"></a>Remover o agrupamento de linhas ou colunas de um intervalo

Para desagrupar um grupo de linhas ou colunas, use o [método Range.ungroup.](/javascript/api/excel/excel.range#ungroup_groupOption_) Isso remove o nível mais externo do contorno. Se vários grupos do mesmo tipo de linha ou coluna estão no mesmo nível dentro do intervalo especificado, todos esses grupos serão desagrupados.

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
