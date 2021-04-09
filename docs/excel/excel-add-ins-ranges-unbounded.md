---
title: Ler ou gravar em um intervalo não-rebote usando a API JavaScript do Excel
description: Saiba como usar a API JavaScript do Excel para ler ou gravar em um intervalo não-rebote.
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f7be2efc3e069ea3451088608ca5255a632ef863
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652756"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a>Ler ou gravar em um intervalo não-rebote usando a API JavaScript do Excel

Este artigo descreve como ler e gravar em um intervalo não-rebote com a API JavaScript do Excel. Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).

Um endereço de intervalo não rebotado é um endereço de intervalo que especifica colunas inteiras ou linhas inteiras. Por exemplo:

- Endereços de intervalo compostos por colunas inteiras:<ul><li>`C:C`</li><li>`A:F`</li></ul>
- Endereços de intervalo compostos por linhas inteiras:<ul><li>`2:2`</li><li>`1:4`</li></ul>

## <a name="read-an-unbounded-range"></a>Ler um intervalo não limitado

Quando uma API faz uma solicitação para recuperar um intervalo não limitado (por exemplo, `getRange('C:C')`), a resposta conterá valores `null` para as propriedades no nível de célula, como `values`, `text`, `numberFormat` e `formula`. Outras propriedades do intervalo, como `address` e `cellCount`, conterão valores válidos para o intervalo não limitado.

## <a name="write-to-an-unbounded-range"></a>Gravar em um intervalo não limitado

Não é possível definir propriedades no nível da célula, como , e em um intervalo não rebotado porque a solicitação de `values` `numberFormat` entrada é muito `formula` grande. Por exemplo, o exemplo de código a seguir não é válido porque ele tenta especificar para um `values` intervalo não-rebote. A API retornará um erro se você tentar definir propriedades no nível da célula para um intervalo não-rebote.

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a API JavaScript do Excel](excel-add-ins-cells.md)
- [Ler ou gravar em um intervalo grande usando a API JavaScript do Excel](excel-add-ins-ranges-large.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
