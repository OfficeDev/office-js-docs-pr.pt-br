---
title: Trabalhe com células usando Excel API JavaScript.
description: Aprenda a Excel definição da API JavaScript de uma célula e saiba como trabalhar com células.
ms.date: 04/16/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f9ce806fa9478835ddf009596315108c88c4f1b4
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744639"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>Trabalhar com células usando a EXCEL JavaScript

A API JavaScript do Excel não tem um objeto ou classe "Célula". Em vez disso, Excel células são `Range` objetos. Uma célula individual na interface do usuário do Excel se traduz em um objeto `Range` com uma célula na API JavaScript do Excel.

Um `Range` objeto também pode conter várias células contíguas. Células contíguas formam um retângulo ininterrupto (incluindo linhas ou colunas simples). Para saber mais sobre como trabalhar com células que não são contíguas, consulte Trabalhar com células [descontíguas usando o objeto RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object).

Para ver a lista completa de propriedades e `Range` métodos que o objeto oferece suporte, consulte [Range Object (API JavaScript para Excel)](/javascript/api/excel/excel.range).

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>Trabalhar com células desconsiguadas usando o objeto RangeAreas

O [objeto RangeAreas](/javascript/api/excel/excel.rangeareas) permite que o seu complemento execute operações em vários intervalos de uma só vez. Esses intervalos podem ser contíguos, mas não precisam ser. `RangeAreas` são descritas ainda mais no artigo [Trabalhar com vários intervalos simultaneamente em suplementos do Excel](excel-add-ins-multiple-ranges.md).

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Obter um intervalo usando a EXCEL JavaScript](excel-add-ins-ranges-get.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
