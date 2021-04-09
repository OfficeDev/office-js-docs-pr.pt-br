---
title: Recorte, copie e colar intervalos usando a API JavaScript do Excel
description: Saiba como cortar, copiar e colar intervalos usando a API JavaScript do Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 8cf92ef148c24613674930140cec762c9cd8c4a4
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652782"
---
# <a name="cut-copy-and-paste-ranges-using-the-excel-javascript-api"></a>Recorte, copie e colar intervalos usando a API JavaScript do Excel

Este artigo fornece exemplos de código que cortam, copiam e colaram intervalos usando a API JavaScript do Excel. Para ver a lista completa de propriedades e métodos compatíveis com o `Range` objeto, consulte [Classe Excel.Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="copy-and-paste"></a>Copy and paste

O [método Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) replica as ações **Copiar** e **Colar** da interface do usuário do Excel. O destino é `Range` o objeto `copyFrom` chamado. A fonte a ser copiada é passada como um intervalo ou um endereço de cadeia de caracteres que representa um intervalo.

O exemplo a seguir copia dados de **A1:E1** para o intervalo que começa em **G1** (que acaba sendo colado em **G1:K1**).

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

`Range.copyFrom` tem três parâmetros opcionais.

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` especifica quais dados são copiados da origem para o destino.

- `Excel.RangeCopyType.formulas` transfere as fórmulas nas células de origem e preserva o posicionamento relativo dos intervalos dessas fórmulas. As entradas que não sejam uma fórmula são copiadas no seu estado original.
- `Excel.RangeCopyType.values` copia os valores dos dados e, no caso de fórmulas, o resultado da fórmula.
- `Excel.RangeCopyType.formats` copia a formatação do intervalo incluindo a fonte, cor e outras configurações de formato, mas nenhum valor.
- `Excel.RangeCopyType.all` (a opção padrão) copia os dados e a formatação, preservando as fórmulas das células, se encontradas.

`skipBlanks` define se as células em branco são copiadas para o destino. Quando for verdadeiro, `copyFrom` ignora células em branco no intervalo de origem.
As células ignoradas não substituem os dados existentes de suas células correspondentes no intervalo de destino. O padrão é false.

`transpose` determina se os dados são transpostos, ou seja, suas linhas e colunas são alternadas para o local de origem.
Um intervalo transposto invertido na diagonal principal, portanto as linhas **1**, **2** e **3** se tornarão as colunas **A**, **B** e **C**.

O exemplo de código e as imagens a seguir demonstram esse comportamento em um cenário simples.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-copied-and-pasted"></a>Dados antes que o intervalo seja copiado e passado

![Dados no Excel antes que o método de cópia do intervalo tenha sido executado](../images/excel-range-copyfrom-skipblanks-before.png)

### <a name="data-after-range-is-copied-and-pasted"></a>Dados após o intervalo são copiados e copiados

![Dados no Excel após o método de cópia do intervalo ter sido executado](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="cut-and-paste-move-cells"></a>Cortar e colar células (mover)

O [método Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) move células para um novo local na workbook. Esse comportamento de movimento de célula funciona [](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) da mesma forma que quando as células são movidas arrastando a borda do intervalo ou ao tomar as ações **Cortar** **e Colar.** Tanto a formatação quanto os valores do intervalo são movidos para o local especificado como o `destinationRange` parâmetro.

O exemplo de código a seguir move um intervalo com o `Range.moveTo` método. Observe que, se o intervalo de destino for menor que a fonte, ele será expandido para abranger o conteúdo de origem.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="see-also"></a>Confira também

- [Modelo de objeto JavaScript do Excel em Suplementos do Office](excel-add-ins-core-concepts.md)
- [Trabalhar com células usando a API JavaScript do Excel](excel-add-ins-cells.md)
- [Remover duplicatas usando a API JavaScript do Excel](excel-add-ins-ranges-remove-duplicates.md)
- [Trabalhar simultaneamente com vários intervalos em suplementos do Excel](excel-add-ins-multiple-ranges.md)
