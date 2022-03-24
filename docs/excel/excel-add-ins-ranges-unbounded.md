---
title: Ler ou gravar em um intervalo não ressalvado usando Excel API JavaScript
description: Saiba como usar a EXCEL JavaScript para ler ou gravar em um intervalo não-rebote.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 5ef9b6a385db5b1de90e1bd61802d20ef7864533
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745499"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a>Ler ou gravar em um intervalo não ressalvado usando Excel API JavaScript

Este artigo descreve como ler e gravar em um intervalo não-rebote com a API JavaScript Excel JavaScript. Para ver a lista completa de propriedades e métodos compatíveis `Range` com o objeto, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

Um endereço de intervalo não rebotado é um endereço de intervalo que especifica colunas inteiras ou linhas inteiras. Por exemplo:

- Endereços de intervalo compostos por colunas inteiras.
  - `C:C`
  - `A:F`
- Endereços de intervalo compostos por linhas inteiras.
  - `2:2`
  - `1:4`

## <a name="read-an-unbounded-range"></a>Ler um intervalo não limitado

Quando uma API faz uma solicitação para recuperar um intervalo não limitado (por exemplo, `getRange('C:C')`), a resposta conterá valores `null` para as propriedades no nível de célula, como `values`, `text`, `numberFormat` e `formula`. Outras propriedades do intervalo, como `address` e `cellCount`, conterão valores válidos para o intervalo não limitado.

## <a name="write-to-an-unbounded-range"></a>Gravar em um intervalo não limitado

Não é possível definir propriedades no nível `values`da célula, como , e `numberFormat``formula` em um intervalo não rebotado porque a solicitação de entrada é muito grande. Por exemplo, o exemplo de código a seguir não é válido porque ele tenta especificar para `values` um intervalo não-rebote. A API retornará um erro se você tentar definir propriedades no nível da célula para um intervalo não-rebote.

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
let range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Ler ou gravar em um intervalo grande usando a EXCEL JavaScript](excel-add-ins-ranges-large.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
