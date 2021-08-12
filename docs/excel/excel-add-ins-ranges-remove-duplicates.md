---
title: Remover duplicatas usando a EXCEL JavaScript
description: Saiba como usar a EXCEL JavaScript para remover duplicatas.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 626676dae57854baf5e8fe195c669a0432eee832c8099715144c2b569ae1055a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57086130"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a>Remover duplicatas usando a EXCEL JavaScript

Este artigo fornece um exemplo de código que remove entradas duplicadas em um intervalo usando Excel API JavaScript. Para ver a lista completa de propriedades e métodos que o `Range` objeto oferece suporte, [consulte Excel. Classe Range](/javascript/api/excel/excel.range).

## <a name="remove-rows-with-duplicate-entries"></a>Remover linhas com entradas duplicadas

O [método Range.removeDuplicates](/javascript/api/excel/excel.range#removeDuplicates_columns__includesHeader_) remove linhas com entradas duplicadas nas colunas especificadas. O método passa por cada linha no intervalo do índice de menor valor até o índice de maior valor no intervalo (de cima para baixo). Uma linha é excluída se um valor em sua coluna ou colunas especificadas aparecer mais cedo no intervalo. Linhas no intervalo abaixo da linha excluída são deslocadas para cima. `removeDuplicates` não afeta a posição de células fora do intervalo.

`removeDuplicates` leva um `number[]` representando os índices da coluna que são verificados para duplicatas. Essa matriz é baseada em zero e relativa ao intervalo, não à planilha. O método também recebe um parâmetro booleano que especifica se a primeira linha é um header. Quando **verdadeiro**, a primeira linha será ignorada ao considerar duplicatas. O método retorna um objeto que especifica o número de linhas removidas e `removeDuplicates` o número de linhas `RemoveDuplicatesResult` exclusivas restantes.

Ao usar o método de `removeDuplicates` um intervalo, lembre-se do seguinte.

- `removeDuplicates` considera valores de célula, não resultados de função. Se as duas funções diferentes forem avaliadas como o mesmo resultado, os valores de célula não são considerados duplicatas.
- Células vazias não serão ignoradas por `removeDuplicates`. O valor de uma célula vazia é tratado como qualquer outro valor. Isso significa que as linhas vazias contidas no intervalo serão incluídas em `RemoveDuplicatesResult`.

O exemplo de código a seguir mostra a remoção de entradas com valores duplicados na primeira coluna.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

### <a name="data-before-duplicate-entries-are-removed"></a>Dados antes que entradas duplicadas sejam removidas

![Dados em Excel antes que o método remove duplicatas do intervalo tenha sido executado.](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a>Dados após entradas duplicadas são removidos

![Dados em Excel após a executar o método remove duplicates do intervalo.](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a EXCEL JavaScript](excel-add-ins-cells.md)
- [Intervalos de corte, cópia e colar usando a API JavaScript Excel JavaScript](excel-add-ins-ranges-cut-copy-paste.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
