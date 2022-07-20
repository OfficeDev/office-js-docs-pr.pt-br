---
title: Remover duplicatas usando a API JavaScript do Excel
description: Saiba como usar a API JavaScript do Excel para remover duplicatas.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9ece7c9f35b341dbb8d0d90e8ca4bda5215580ed
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889139"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a>Remover duplicatas usando a API JavaScript do Excel

Este artigo fornece um exemplo de código que remove entradas duplicadas em um intervalo usando a API JavaScript do Excel. Para obter a lista completa de propriedades e métodos compatíveis `Range` com o objeto, consulte a [classe Excel.Range](/javascript/api/excel/excel.range).

## <a name="remove-rows-with-duplicate-entries"></a>Remover linhas com entradas duplicadas

O [método Range.removeDuplicates](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1)) remove linhas com entradas duplicadas nas colunas especificadas. O método passa por cada linha no intervalo do índice de menor valor até o índice de valor mais alto no intervalo (de cima para baixo). Uma linha é excluída se um valor em sua coluna ou colunas especificadas aparecer mais cedo no intervalo. Linhas no intervalo abaixo da linha excluída são deslocadas para cima. `removeDuplicates` não afeta a posição de células fora do intervalo.

`removeDuplicates` leva um `number[]` representando os índices da coluna que são verificados para duplicatas. Essa matriz é baseada em zero e relativa ao intervalo, não à planilha. O método também usa um parâmetro booliano que especifica se a primeira linha é um cabeçalho. Quando `true`, a linha superior é ignorada ao considerar duplicatas. O `removeDuplicates` método retorna um `RemoveDuplicatesResult` objeto que especifica o número de linhas removidas e o número de linhas exclusivas restantes.

Ao usar o método de um `removeDuplicates` intervalo, lembre-se do seguinte.

- `removeDuplicates` considera valores de célula, não resultados de função. Se as duas funções diferentes forem avaliadas como o mesmo resultado, os valores de célula não são considerados duplicatas.
- Células vazias não serão ignoradas por `removeDuplicates`. O valor de uma célula vazia é tratado como qualquer outro valor. Isso significa que as linhas vazias contidas no intervalo serão incluídas em `RemoveDuplicatesResult`.

O exemplo de código a seguir mostra a remoção de entradas com valores duplicados na primeira coluna.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B2:D11");

    let deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    await context.sync();

    console.log(deleteResult.removed + " entries with duplicate names removed.");
    console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
});
```

### <a name="data-before-duplicate-entries-are-removed"></a>Dados antes de entradas duplicadas serem removidas

![Dados no Excel antes que o método remove duplicatas do intervalo tenha sido executado.](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a>Dados após a remoção de entradas duplicadas

![Dados no Excel após a execução do método remove duplicates do intervalo.](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a API JavaScript do Excel](excel-add-ins-cells.md)
- [Recortar, copiar e colar intervalos usando a API JavaScript do Excel](excel-add-ins-ranges-cut-copy-paste.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
