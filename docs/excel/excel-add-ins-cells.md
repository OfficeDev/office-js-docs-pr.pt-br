---
title: Trabalhe com células usando a API JavaScript do Excel.
description: Aprenda a definição da API JavaScript do Excel de uma célula e saiba como trabalhar com células.
ms.date: 04/16/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad8ca985b6bbdcf19920c36c371e690f61639f16
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917097"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>Trabalhar com células usando a API JavaScript do Excel

A API JavaScript do Excel não tem um objeto ou classe "Célula". Em vez disso, todas as células do Excel são `Range` objetos. Uma célula individual na interface do usuário do Excel se traduz em um objeto `Range` com uma célula na API JavaScript do Excel.

Um `Range` objeto também pode conter várias células contíguas. Células contíguas formam um retângulo ininterrupto (incluindo linhas ou colunas simples). Para saber mais sobre como trabalhar com células que não são contíguas, consulte Trabalhar com células [descontíguas usando o objeto RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object).

Para ver a lista completa de propriedades e métodos compatíveis com o objeto, consulte `Range` [Range Object (API JavaScript para Excel)](/javascript/api/excel/excel.range).

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>Trabalhar com células desconsiguadas usando o objeto RangeAreas

O [objeto RangeAreas](/javascript/api/excel/excel.rangeareas) permite que o seu complemento execute operações em vários intervalos de uma só vez. Esses intervalos podem ser contíguos, mas não precisam ser. `RangeAreas` são descritas ainda mais no artigo [Trabalhar com vários intervalos simultaneamente em suplementos do Excel](excel-add-ins-multiple-ranges.md).

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Obter um intervalo usando a API JavaScript do Excel](excel-add-ins-ranges-get.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
