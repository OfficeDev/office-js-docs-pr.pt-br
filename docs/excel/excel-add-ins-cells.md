---
title: Trabalhe com células usando a API JavaScript do Excel.
description: Aprenda a definição da API JavaScript do Excel de uma célula e saiba como trabalhar com células.
ms.date: 04/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5fcfeeef52f17c22d13ed3c1a10851f1d8e69204
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652813"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>Trabalhar com células usando a API JavaScript do Excel

A API JavaScript do Excel não tem um objeto ou classe "Cell". Em vez disso, todas as células do Excel são `Range` objetos. Uma célula individual na interface do usuário do Excel é traduzida para um `Range` objeto com uma célula na API JavaScript do Excel.

Um `Range` objeto também pode conter várias células contíguas. Células contíguas formam um retângulo ininterrupto (incluindo linhas ou colunas simples). Para saber mais sobre como trabalhar com células que não são contíguas, consulte Trabalhar com células [descontíguas usando o objeto RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object).

Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).

## <a name="excel-javascript-apis-that-mention-cells"></a>APIs JavaScript do Excel que mencionam células

Mesmo que a API JavaScript do Excel não tenha um objeto ou classe "Cell", vários nomes de API mencionam células. Essas APIs controlam propriedades de célula, como cor, formatação de texto e fonte.

A lista a seguir das APIs JavaScript do Excel referem-se a células.

- [CellBorder](/javascript/api/excel/excel.cellborder)
- [CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)
- [CellProperties](/javascript/api/excel/excel.cellproperties)
- [CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)
- [CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)
- [CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)
- [CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)
- [CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)
- [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)
- [SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>Trabalhar com células desconsiguadas usando o objeto RangeAreas

O [objeto RangeAreas](/javascript/api/excel/excel.rangeareas) permite que o seu complemento execute operações em vários intervalos de uma só vez. Esses intervalos podem ser contíguos, mas não precisam ser. `RangeAreas` são descritas ainda mais no artigo [Trabalhar com vários intervalos simultaneamente em suplementos do Excel](excel-add-ins-multiple-ranges.md).

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Obter um intervalo usando a API JavaScript do Excel](excel-add-ins-ranges-get.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
